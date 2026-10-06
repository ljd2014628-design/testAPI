# -*- coding: utf-8 -*-
"""接口测试工具 · 核心逻辑（无界面依赖）"""

import os
import re
import sys
import json
import time
import copy
import threading

import requests
import urllib3

import excel_io
from excel_io import (MAX_HISTORY, now_str, dump_params_text, parse_params_text,
                      dump_memos_text, parse_memos_text, dump_asserts_text,
                      parse_asserts_text)

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

FROZEN = getattr(sys, "frozen", False)
RES_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
APP_DIR = os.path.dirname(sys.executable) if FROZEN else os.path.dirname(os.path.abspath(__file__))


def _pick_data_dir():
    cands = [os.path.join(APP_DIR, "data"),
             os.path.join(os.path.expanduser("~"), "ApiTestTool", "data")]
    for d in cands:
        try:
            os.makedirs(d, exist_ok=True)
            p = os.path.join(d, ".wtest")
            with open(p, "w") as f:
                f.write("1")
            os.remove(p)
            return d
        except Exception:
            continue
    import tempfile
    d = os.path.join(tempfile.gettempdir(), "ApiTestTool")
    os.makedirs(d, exist_ok=True)
    return d


DATA_DIR = _pick_data_dir()
STATE_FILE = os.path.join(DATA_DIR, "state.json")


def default_state():
    return {
        "config": {
            "baseUrl": "",
            "encoding": "utf-8",
            "timeout": 30,
            "verifySsl": False,
            "stopOnAssertFail": False,
            "headers": [
                {"key": "Content-Type", "value": "application/json;charset=utf-8", "enabled": True},
                {"key": "Accept", "value": "application/json", "enabled": True},
            ],
        },
        "cases": [],
        "scenarios": [],
        "fileName": "",
        "vars": {},
    }


STATE = default_state()


_save_lock = threading.Lock()
_save_timer = None
SAVE_DELAY = 1.2          # 秒：把连续多次改动合并成一次写盘


def _write_state_now():
    try:
        tmp = STATE_FILE + ".tmp"
        with _save_lock:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(STATE, f, ensure_ascii=False)
            os.replace(tmp, STATE_FILE)
    except Exception as e:
        print("保存失败:", e)


def save_state(now=False):
    """默认延迟合并写盘；now=True 立即写（退出、导出前用）。"""
    global _save_timer
    if _save_timer is not None:
        try:
            _save_timer.cancel()
        except Exception:
            pass
        _save_timer = None
    if now:
        _write_state_now()
        return
    _save_timer = threading.Timer(SAVE_DELAY, _write_state_now)
    _save_timer.daemon = True
    _save_timer.start()


def flush_state():
    save_state(now=True)


def state_size_mb():
    try:
        return os.path.getsize(STATE_FILE) / 1048576.0
    except Exception:
        return 0.0


def load_state():
    global STATE
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, encoding="utf-8") as f:
                data = json.load(f)
            base = default_state()
            base.update(data)
            STATE = base
        except Exception as e:
            print("读取历史数据失败:", e)


# ---------------------------------------------------------------- 变量
VAR_PATTERN = re.compile(r"\$\{\s*([^}\s]+)\s*\}")


def render(text, variables):
    if not text:
        return text

    def rep(m):
        key = m.group(1)
        if key in variables:
            v = variables[key]
            if isinstance(v, (dict, list)):
                return json.dumps(v, ensure_ascii=False)
            return "" if v is None else str(v)
        return m.group(0)

    return VAR_PATTERN.sub(rep, text)


MISSING = object()


def _tokens(path):
    """把 $.a.b[0].c / a.list[*].id 解析成 token 列表，* 表示通配所有元素。"""
    p = (path or "").strip()
    if p.startswith("$"):
        p = p[1:]
    p = p.lstrip(".")
    tokens, token, i = [], "", 0
    while i < len(p):
        ch = p[i]
        if ch == ".":
            if token:
                tokens.append(token)
                token = ""
        elif ch == "[":
            if token:
                tokens.append(token)
                token = ""
            j = p.find("]", i)
            if j < 0:
                break
            inner = p[i + 1:j].strip().strip('"\'')
            if inner == "*":
                tokens.append("*")
            elif inner.lstrip("-").isdigit():
                tokens.append(int(inner))
            else:
                tokens.append(inner)
            i = j
        else:
            token += ch
        i += 1
    if token:
        tokens.append(token)
    return tokens


def _resolve(cur, tokens):
    if not tokens:
        return cur
    t, rest = tokens[0], tokens[1:]
    if t == "*":
        if isinstance(cur, list):
            vals = [_resolve(x, rest) for x in cur]
        elif isinstance(cur, dict):
            vals = [_resolve(v, rest) for v in cur.values()]
        else:
            return MISSING
        if vals and all(v is MISSING for v in vals):
            return MISSING
        return [None if v is MISSING else v for v in vals]
    if isinstance(t, int):
        if isinstance(cur, list) and -len(cur) <= t < len(cur):
            return _resolve(cur[t], rest)
        if isinstance(cur, dict) and str(t) in cur:
            return _resolve(cur[str(t)], rest)
        return MISSING
    if isinstance(cur, dict):
        if t not in cur:
            return MISSING
        return _resolve(cur[t], rest)
    if isinstance(cur, list):
        # 在数组上直接取字段：等价于 [*].字段
        vals = [_resolve(x, [t] + rest) for x in cur]
        if vals and all(v is MISSING for v in vals):
            return MISSING
        return [None if v is MISSING else v for v in vals]
    return MISSING


def json_path_ex(data, path):
    """取值；路径不存在返回 MISSING（用于区分“字段不存在”和“值就是 null”）。"""
    return _resolve(data, _tokens(path))


def json_path(data, path):
    v = json_path_ex(data, path)
    return None if v is MISSING else v


def parse_extract_detail(text):
    """逐行解析提取配置，返回每行的变量名/路径/格式错误，供调试窗口使用。"""
    out = []
    for no, line in enumerate((text or "").splitlines(), start=1):
        raw = line.strip()
        if not raw or raw.startswith("#"):
            continue
        if "=" not in raw:
            out.append({"line": no, "raw": raw, "var": "", "path": raw,
                        "error": "格式错误：缺少「变量名=」，应写成  变量名=%s" % raw})
            continue
        k, v = raw.split("=", 1)
        k, v = k.strip(), v.strip()
        if not k:
            out.append({"line": no, "raw": raw, "var": "", "path": v,
                        "error": "格式错误：等号左边的变量名为空"})
        elif not v:
            out.append({"line": no, "raw": raw, "var": k, "path": "",
                        "error": "格式错误：等号右边的 JSON 路径为空"})
        else:
            out.append({"line": no, "raw": raw, "var": k, "path": v, "error": ""})
    return out


def parse_extract(text):
    return [(d["var"], d["path"]) for d in parse_extract_detail(text) if not d["error"]]


def parse_headers(text):
    text = (text or "").strip()
    if not text:
        return {}
    if text.startswith("{"):
        try:
            obj = json.loads(text)
            return {str(k): ("" if v is None else str(v)) for k, v in obj.items()}
        except Exception:
            pass
    out = {}
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if ":" in line:
            k, v = line.split(":", 1)
        elif "=" in line:
            k, v = line.split("=", 1)
        else:
            continue
        out[k.strip()] = v.strip()
    return out


def merge_headers(cfg, case, variables):
    """公共头打底，接口级同名覆盖（不区分大小写）。"""
    merged = {}
    for h in cfg.get("headers", []):
        if h.get("enabled", True) and h.get("key"):
            merged[h["key"].strip()] = render(str(h.get("value", "")), variables)
    for k, v in parse_headers(case.get("headers", "")).items():
        for exist in list(merged.keys()):
            if exist.lower() == k.lower():
                merged.pop(exist)
        merged[k] = render(v, variables)
    return merged


def encode_headers(headers, enc):
    out = {}
    for k, v in headers.items():
        v = "" if v is None else str(v)
        try:
            k.encode("ascii")
            v.encode("ascii")
            out[k] = v
        except UnicodeEncodeError:
            out[k.encode(enc, errors="replace")] = v.encode(enc, errors="replace")
    return out


def build_url(cfg, raw_url, variables, params=None):
    url = render((raw_url or "").strip(), variables)
    if not url.lower().startswith(("http://", "https://")):
        base = (cfg.get("baseUrl") or "").strip().rstrip("/")
        if not url.startswith("/"):
            url = "/" + url
        url = base + url
    return join_query(url, params, variables)


def split_query(url):
    """把 URL 里的查询串拆成参数表，返回 (不带参数的URL, 参数列表)。"""
    url = url or ""
    if "?" not in url:
        return url, []
    base, q = url.split("?", 1)
    params = []
    for part in q.split("&"):
        if not part:
            continue
        k, v = part.split("=", 1) if "=" in part else (part, "")
        if k:
            params.append({"key": k, "value": v, "memo": "", "on": True})
    return base, params


def join_query(url, params, variables=None):
    """把参数表拼回 URL；停用的行不发送，备注永远不发送。"""
    items = []
    for p in params or []:
        if not p.get("on", True) or not p.get("key"):
            continue
        v = p.get("value", "") or ""
        if variables:
            v = render(str(v), variables)
        items.append("%s=%s" % (p["key"], v))
    if not items:
        return url
    return url + ("&" if "?" in url else "?") + "&".join(items)


# ---------------------------------------------------------------- Body 字段
def flatten_fields(data, prefix=""):
    """把 JSON 摊平成 [(字段路径, 值)]，用于逐字段加备注。"""
    out = []
    if isinstance(data, dict):
        for k, v in data.items():
            out.extend(flatten_fields(v, ("%s.%s" % (prefix, k)) if prefix else str(k)))
    elif isinstance(data, list):
        for i, v in enumerate(data):
            out.extend(flatten_fields(v, "%s[%d]" % (prefix, i)))
    else:
        out.append((prefix, data))
    return out


def set_field(data, path, value):
    """按字段路径写回值，成功返回 True。"""
    tokens = _tokens(path)
    if not tokens:
        return False
    cur = data
    try:
        for t in tokens[:-1]:
            cur = cur[t]
        cur[tokens[-1]] = value
        return True
    except Exception:
        return False


def shorten(text, limit=60000):
    if text is None:
        return ""
    text = str(text)
    return text if len(text) <= limit else text[:limit] + "\n...(内容过长，已截断)"


# ---------------------------------------------------------------- 断言
# 比较方式按用途分组，界面上的下拉框按这个顺序展示
ASSERT_GROUPS = [
    ("数值 / 相等", ["等于", "不等于", "文本等于", "约等于", "大于", "大于等于",
                     "小于", "小于等于", "介于"]),
    ("文本", ["包含", "不包含", "开头是", "结尾是", "正则匹配", "正则不匹配"]),
    ("存在 / 空值", ["存在", "不存在", "为空", "不为空", "为null", "不为null"]),
    ("类型 / 结构", ["类型是", "必含字段", "是有效JSON"]),
    ("数组 / 集合", ["长度等于", "长度大于", "长度小于", "属于", "不属于",
                     "包含元素", "不包含元素", "每项都不为空", "每项都含字段",
                     "无重复", "升序", "降序"]),
    ("与上次比较（推荐用于领券/库存这类每次都变的数据）",
     ["较上次增加", "较上次减少", "较上次不变", "较上次变化", "较上次变化不超过"]),
]
ASSERT_OPS = [op for _g, ops in ASSERT_GROUPS for op in ops]
ASSERT_LEVELS = ["必须通过", "仅警告"]
REL_OPS = ("较上次增加", "较上次减少", "较上次不变", "较上次变化", "较上次变化不超过")
TYPE_NAMES = {"字符串": str, "数字": (int, float), "布尔": bool,
              "数组": list, "对象": dict, "null": type(None)}

ASSERT_TARGET_HINTS = ["状态码", "耗时ms", "响应文本", "响应头.Content-Type",
                       "$.code", "$.msg", "$.data"]


def _to_num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except Exception:
        return None


def _calc(expr):
    """把 '10-1'、'${n}+2' 这类算式求值；不是算式就原样返回。"""
    import ast
    txt = str(expr).strip()
    if not txt or not re.fullmatch(r"[0-9\.\s\+\-\*/\(\)]+", txt):
        return txt
    try:
        node = ast.parse(txt, mode="eval")
        allowed = (ast.Expression, ast.BinOp, ast.UnaryOp, ast.Constant,
                   ast.Add, ast.Sub, ast.Mult, ast.Div, ast.USub, ast.UAdd)
        for n in ast.walk(node):
            if not isinstance(n, allowed):
                return txt
        val = eval(compile(node, "<calc>", "eval"), {"__builtins__": {}}, {})
        return int(val) if float(val).is_integer() else val
    except Exception:
        return txt


def _fmt_val(v):
    if v is MISSING:
        return "（字段不存在）"
    if isinstance(v, str):
        return v if len(v) <= 120 else v[:120] + "…"
    try:
        t = json.dumps(v, ensure_ascii=False)
    except Exception:
        t = str(v)
    return t if len(t) <= 120 else t[:120] + "…"


def _split_list(expected):
    """把 'a,b,c' 或 JSON 数组拆成列表。"""
    txt = str(expected).strip()
    if txt.startswith("["):
        try:
            v = json.loads(txt)
            if isinstance(v, list):
                return [str(x) for x in v]
        except Exception:
            pass
    return [x.strip() for x in re.split(r"[,，|]", txt) if x.strip()]


def _as_list(actual):
    if isinstance(actual, list):
        return actual
    if isinstance(actual, dict):
        return list(actual.values())
    return None


def assert_actual(target, parsed, resp_text, status, elapsed, headers=None):
    t = (target or "").strip()
    if t in ("状态码", "状态", "status"):
        return status
    if t in ("耗时ms", "耗时", "elapsed"):
        return elapsed
    if t in ("响应文本", "响应", "body"):
        return resp_text
    if t.startswith("响应头."):
        name = t[4:].strip().lower()
        for k, v in (headers or {}).items():
            if str(k).lower() == name:
                return v
        return MISSING
    if parsed is None:
        return MISSING
    return json_path_ex(parsed, t)


def eval_asserts(asserts, parsed, resp_text, status, elapsed, variables,
                 prev_asserts=None, headers=None):
    """逐条判定断言。prev_asserts 是上一次执行的断言结果，供「较上次」类比较使用。"""
    prev_map = {}
    for p in prev_asserts or []:
        prev_map.setdefault(p.get("target"), p.get("actualRaw"))

    out = []
    for a in asserts or []:
        if not a.get("on", True) or not a.get("target"):
            continue
        target = a["target"]
        op = a.get("op") or "等于"
        level = a.get("level") or "必须通过"
        expected_raw = render(str(a.get("expected", "") or ""), variables or {})
        expected = _calc(expected_raw)
        actual = assert_actual(target, parsed, resp_text, status, elapsed, headers)
        try:
            ok, msg = _compare(op, actual, expected, prev_map.get(target))
        except Exception as e:
            ok, msg = False, "断言执行出错：%s: %s" % (type(e).__name__, e)
        out.append({"target": target, "op": op, "expected": expected_raw,
                    "level": level, "memo": a.get("memo", ""),
                    "actual": _fmt_val(actual),
                    "actualRaw": None if actual is MISSING else actual,
                    "ok": ok, "msg": msg})
    return out


def assert_stats(rows):
    """统计：(失败数, 警告数, 通过数)。仅警告的失败不算失败。"""
    fail = warn = passed = 0
    for r in rows or []:
        if r.get("ok"):
            passed += 1
        elif r.get("level") == "仅警告":
            warn += 1
        else:
            fail += 1
    return fail, warn, passed


def _compare(op, actual, expected, prev):
    a_num, e_num = _to_num(actual), _to_num(expected)
    a_txt = "" if actual is MISSING else (
        actual if isinstance(actual, str) else json.dumps(actual, ensure_ascii=False, default=str))
    e_txt = str(expected)

    # ---- 存在性：允许 actual 为 MISSING ----
    if op == "存在":
        return (actual is not MISSING), ("字段存在，值 %s" % _fmt_val(actual)
                                         if actual is not MISSING else "字段不存在")
    if op == "不存在":
        return (actual is MISSING), ("确实不存在" if actual is MISSING
                                     else "字段存在，实际 %s" % _fmt_val(actual))
    if actual is MISSING:
        return False, "取不到这个字段，请检查路径是否写对"

    # ---- 空值 ----
    if op == "为空":
        empty = actual in (None, "", [], {})
        return empty, ("确实为空" if empty else "实际 %s" % _fmt_val(actual))
    if op == "不为空":
        empty = actual in (None, "", [], {})
        return (not empty), ("非空：%s" % _fmt_val(actual) if not empty else "实际为空")
    if op == "为null":
        return actual is None, ("确实是 null" if actual is None else "实际 %s" % _fmt_val(actual))
    if op == "不为null":
        return actual is not None, ("不是 null：%s" % _fmt_val(actual)
                                    if actual is not None else "实际是 null")

    # ---- 类型 / 结构 ----
    if op == "类型是":
        want = TYPE_NAMES.get(e_txt.strip())
        if want is None:
            return False, "类型只能填：%s" % "、".join(TYPE_NAMES)
        if e_txt.strip() == "数字" and isinstance(actual, bool):
            return False, "实际是布尔，不是数字"
        ok = isinstance(actual, want)
        return ok, "实际类型 %s，期望 %s" % (type(actual).__name__, e_txt)
    if op == "必含字段":
        if not isinstance(actual, dict):
            return False, "这个值不是对象，无法检查字段，实际 %s" % _fmt_val(actual)
        need = _split_list(expected)
        miss = [k for k in need if k not in actual]
        return (not miss), ("字段齐全：%s" % "、".join(need) if not miss
                            else "缺少字段：%s" % "、".join(miss))
    if op == "是有效JSON":
        try:
            json.loads(a_txt)
            return True, "是合法 JSON"
        except Exception as e:
            return False, "不是合法 JSON：%s" % e

    # ---- 数组 / 集合 ----
    if op in ("长度等于", "长度大于", "长度小于"):
        try:
            ln = len(actual)
        except Exception:
            return False, "这个值没有长度概念，实际 %s" % _fmt_val(actual)
        want = e_num if e_num is not None else 0
        ok = {"长度等于": ln == want, "长度大于": ln > want, "长度小于": ln < want}[op]
        return ok, "实际长度 %d，期望 %s %s" % (ln, op[2:], e_txt)
    if op in ("属于", "不属于"):
        opts = _split_list(expected)
        inside = str(actual) in opts or a_txt in opts
        ok = inside if op == "属于" else (not inside)
        return ok, "实际 %s，候选 [%s]" % (_fmt_val(actual), "、".join(opts))
    if op in ("包含元素", "不包含元素"):
        lst = _as_list(actual)
        if lst is None:
            return False, "这个值不是数组，实际 %s" % _fmt_val(actual)
        vals = [json.dumps(x, ensure_ascii=False) if not isinstance(x, str) else x for x in lst]
        has = (e_txt in vals) or (e_txt in [str(x) for x in lst])
        ok = has if op == "包含元素" else (not has)
        return ok, "数组共 %d 项，%s「%s」" % (len(lst), "含有" if has else "没有", e_txt)
    if op == "每项都不为空":
        lst = _as_list(actual)
        if lst is None:
            return False, "这个值不是数组"
        bad = [i for i, x in enumerate(lst) if x in (None, "", [], {})]
        return (not bad), ("%d 项都不为空" % len(lst) if not bad
                           else "第 %s 项为空" % "、".join(str(i + 1) for i in bad[:10]))
    if op == "每项都含字段":
        lst = _as_list(actual)
        if lst is None:
            return False, "这个值不是数组"
        need = _split_list(expected)
        bad = []
        for i, x in enumerate(lst):
            if not isinstance(x, dict) or any(k not in x for k in need):
                bad.append(i + 1)
        return (not bad), ("%d 项都含 %s" % (len(lst), "、".join(need)) if not bad
                           else "第 %s 项缺字段" % "、".join(str(i) for i in bad[:10]))
    if op == "无重复":
        lst = _as_list(actual)
        if lst is None:
            return False, "这个值不是数组"
        keys = [json.dumps(x, ensure_ascii=False, sort_keys=True) if not isinstance(x, str) else x
                for x in lst]
        dup = {k for k in keys if keys.count(k) > 1}
        return (not dup), ("%d 项无重复" % len(lst) if not dup
                           else "存在重复：%s" % "、".join(list(dup)[:5]))
    if op in ("升序", "降序"):
        lst = _as_list(actual)
        if lst is None:
            return False, "这个值不是数组"
        field = str(expected).strip()
        try:
            vals = [(_to_num(x.get(field)) if isinstance(x, dict) else _to_num(x))
                    if field else _to_num(x) for x in lst]
            if any(v is None for v in vals):
                vals = [str(x.get(field)) if isinstance(x, dict) and field else str(x) for x in lst]
        except Exception:
            return False, "无法取到排序字段 %s" % field
        ok = all(vals[i] <= vals[i + 1] for i in range(len(vals) - 1)) if op == "升序" else \
             all(vals[i] >= vals[i + 1] for i in range(len(vals) - 1))
        return ok, "%d 项%s%s" % (len(vals), "确实" if ok else "不是", op)

    # ---- 与上次比较 ----
    if op in REL_OPS:
        if prev is None:
            return True, "首次执行，记录基线 %s；下次起按「%s」比较" % (_fmt_val(actual), op)
        p_num = _to_num(prev)
        if p_num is None or a_num is None:
            same = (str(prev) == str(actual))
            if op == "较上次不变":
                return same, ("与上次一致：%s" % _fmt_val(actual) if same
                              else "上次 %s → 本次 %s" % (_fmt_val(prev), _fmt_val(actual)))
            if op == "较上次变化":
                return (not same), "上次 %s → 本次 %s" % (_fmt_val(prev), _fmt_val(actual))
            return False, "非数字，无法比较增减：上次 %s → 本次 %s" % (_fmt_val(prev), _fmt_val(actual))
        diff = a_num - p_num
        tip = "上次 %s → 本次 %s（变化 %+g）" % (_fmt_val(prev), _fmt_val(actual), diff)
        if op == "较上次不变":
            return diff == 0, tip
        if op == "较上次变化":
            return diff != 0, tip
        if op == "较上次变化不超过":
            lim = abs(e_num) if e_num is not None else 0
            return abs(diff) <= lim, tip + "，允许范围 ±%g" % lim
        step = e_num if e_num is not None else 1
        want = step if op == "较上次增加" else -step
        return diff == want, tip + "，期望%s%g" % ("增加" if want > 0 else "减少", abs(want))

    # ---- 数值 / 相等 ----
    if op == "文本等于":
        same = (str(actual) == e_txt)
        return same, "实际「%s」，期望「%s」" % (_fmt_val(actual), e_txt)
    if op in ("等于", "不等于"):
        if a_num is not None and e_num is not None:
            same = (a_num == e_num)
        else:
            same = (a_txt == e_txt) or (str(actual) == e_txt)
        ok = same if op == "等于" else (not same)
        return ok, "实际 %s，期望%s %s" % (_fmt_val(actual), "" if op == "等于" else "不等于", e_txt)
    if op == "约等于":
        m = re.match(r"^\s*([-\d\.]+)\s*(?:[±~]\s*([\d\.]+))?\s*$", e_txt)
        if not m or a_num is None:
            return False, "写法：期望值填「100±5」或「100~5」；实际 %s" % _fmt_val(actual)
        base = float(m.group(1))
        tol = float(m.group(2) or 0)
        ok = abs(a_num - base) <= tol
        return ok, "实际 %s，期望 %g ± %g" % (_fmt_val(actual), base, tol)
    if op == "介于":
        m = re.match(r"^\s*([-\d\.]+)\s*[~至\-]\s*([-\d\.]+)\s*$", e_txt)
        if not m or a_num is None:
            return False, "写法：期望值填「1~100」；实际 %s" % _fmt_val(actual)
        lo, hi = float(m.group(1)), float(m.group(2))
        ok = lo <= a_num <= hi
        return ok, "实际 %s，期望在 %g ~ %g 之间" % (_fmt_val(actual), lo, hi)
    if op in ("大于", "大于等于", "小于", "小于等于"):
        if a_num is None or e_num is None:
            return False, "不是数字，无法比较大小：实际 %s" % _fmt_val(actual)
        ok = {"大于": a_num > e_num, "大于等于": a_num >= e_num,
              "小于": a_num < e_num, "小于等于": a_num <= e_num}[op]
        return ok, "实际 %s，期望 %s %s" % (_fmt_val(actual), op, e_txt)

    # ---- 文本 ----
    if op in ("包含", "不包含"):
        has = e_txt in a_txt
        ok = has if op == "包含" else (not has)
        return ok, "%s「%s」" % ("含有" if has else "没有", e_txt)
    if op == "开头是":
        ok = a_txt.startswith(e_txt)
        return ok, "实际 %s" % _fmt_val(actual)
    if op == "结尾是":
        ok = a_txt.endswith(e_txt)
        return ok, "实际 %s" % _fmt_val(actual)
    if op in ("正则匹配", "正则不匹配"):
        try:
            hit = re.search(e_txt, a_txt) is not None
        except Exception as e:
            return False, "正则有误：%s" % e
        ok = hit if op == "正则匹配" else (not hit)
        return ok, ("匹配成功" if hit else "没匹配上") + "，实际 %s" % _fmt_val(actual)
    return False, "不认识的比较方式：%s" % op


def probe_asserts(case, asserts, use_cached=True):
    """调试断言：取一份响应逐条判定，不写入执行历史，也不影响「较上次」基线。"""
    resp_text, source, _ = probe_extract(case, "", use_cached)
    try:
        parsed = json.loads(resp_text)
    except Exception:
        parsed = None
    prev = None
    for r in (case.get("results") or []):
        if r.get("asserts"):
            prev = r["asserts"]
            break
    rows = eval_asserts(asserts, parsed, resp_text, "(调试)", 0,
                        STATE.get("vars") or {}, prev)
    return resp_text, source, rows


# ---------------------------------------------------------------- 执行
def run_case(case, cfg, variables, log, prev_asserts=None):
    enc = cfg.get("encoding") or "utf-8"
    method = (case.get("method") or "GET").upper()
    url = build_url(cfg, case.get("url"), variables, case.get("params"))
    headers = merge_headers(cfg, case, variables)
    body_raw = render(case.get("body", ""), variables)

    data = params = json_body = None
    if body_raw.strip():
        if method in ("GET", "DELETE", "HEAD"):
            try:
                params = json.loads(body_raw)
            except Exception:
                data = body_raw.encode(enc, errors="replace")
        else:
            try:
                json_body = json.loads(body_raw)
            except Exception:
                data = body_raw.encode(enc, errors="replace")

    res = {"time": now_str(), "method": method, "url": url,
           "reqHeaders": json.dumps(headers, ensure_ascii=False),
           "reqBody": shorten(body_raw), "caseId": case.get("id", ""),
           "caseName": case.get("name", "")}

    t0 = time.time()
    try:
        if json_body is not None and not any(k.lower() == "content-type" for k in headers):
            headers["Content-Type"] = "application/json;charset=%s" % enc
        kwargs = dict(headers=encode_headers(headers, enc),
                      timeout=cfg.get("timeout", 30), verify=bool(cfg.get("verifySsl")))
        if params is not None:
            kwargs["params"] = params
        if json_body is not None:
            kwargs["data"] = json.dumps(json_body, ensure_ascii=False).encode(enc, errors="replace")
        elif data is not None:
            kwargs["data"] = data
        r = requests.request(method, url, **kwargs)
        elapsed = int((time.time() - t0) * 1000)
        if not r.encoding or r.encoding.lower() == "iso-8859-1":
            r.encoding = enc
        try:
            parsed = r.json()
            pretty = json.dumps(parsed, ensure_ascii=False, indent=2)
        except Exception:
            parsed, pretty = None, r.text
        res.update({"ok": 200 <= r.status_code < 300, "status": r.status_code,
                    "elapsed": elapsed, "respBody": shorten(pretty),
                    "respHeaders": json.dumps(dict(r.headers), ensure_ascii=False)})
        picked = {}
        for k, path in parse_extract(case.get("extract", "")):
            val = json_path(parsed, path) if parsed is not None else None
            picked[k] = val
            variables[k] = val
        if picked:
            res["vars"] = json.dumps(picked, ensure_ascii=False)
        try:
            rows = eval_asserts(case.get("asserts"), parsed, r.text, r.status_code,
                                elapsed, variables, prev_asserts, dict(r.headers))
        except Exception as ae:          # 断言自身出错不能被当成请求失败
            rows = [{"target": "(断言执行出错)", "op": "", "expected": "", "memo": "",
                     "actual": "", "actualRaw": None, "ok": False,
                     "msg": "%s: %s" % (type(ae).__name__, ae)}]
        fail, warn, passed = assert_stats(rows)
        res["asserts"] = rows
        res["assertOk"] = (fail == 0) if rows else None
        res["assertWarn"] = warn
        log("[%s] %s %s -> %s  (%sms)%s" % (
            case.get("id"), method, url, r.status_code, elapsed,
            "" if not rows else "  断言 通过%d/失败%d/警告%d" % (passed, fail, warn)))
        for x in rows:
            if not x["ok"]:
                log("      %s 断言%s：%s [%s] %s —— %s"
                    % ("⚠" if x.get("level") == "仅警告" else "✗",
                       "警告" if x.get("level") == "仅警告" else "不通过",
                       x["target"], x["op"], x["expected"], x["msg"]))
    except Exception as e:
        elapsed = int((time.time() - t0) * 1000)
        res.update({"ok": False, "status": "ERR", "elapsed": elapsed,
                    "error": "%s: %s" % (type(e).__name__, e), "respBody": "",
                    "asserts": ([{"target": "(请求失败)", "op": "", "expected": "", "memo": "",
                                  "actual": "", "actualRaw": None, "ok": False,
                                  "msg": "请求没有成功，断言未执行"}] if case.get("asserts") else []),
                    "assertOk": (False if case.get("asserts") else None)})
        log("[%s] %s %s -> 异常 %s" % (case.get("id"), method, url, e))
    return res


def push_result(container, res):
    container.setdefault("results", []).insert(0, res)
    del container["results"][MAX_HISTORY:]


def find_case(cid):
    for c in STATE["cases"]:
        if c.get("id") == cid:
            return c
    return None


def find_scene(sid):
    for s in STATE["scenarios"]:
        if s.get("id") == sid:
            return s
    return None


def run_scenario(sc, cfg, log):
    variables = dict(STATE.get("vars") or {})
    prev_steps = (sc.get("results") or [{}])[0].get("steps") or []
    steps_out, ok_all = [], True
    for si, st in enumerate(sc.get("steps", [])):
        case = find_case(st.get("caseId"))
        if not case:
            steps_out.append({"caseId": st.get("caseId"), "caseName": "(用例不存在)",
                              "status": "ERR", "error": "用例ID不存在：%s" % st.get("caseId"),
                              "url": "", "reqBody": "", "method": "", "elapsed": 0, "ok": False})
            ok_all = False
            continue
        step_case = copy.deepcopy(case)
        if (st.get("extract") or "").strip():
            step_case["extract"] = st["extract"]
        prev = prev_steps[si].get("asserts") if si < len(prev_steps) else None
        res = run_case(step_case, cfg, variables, log, prev)
        steps_out.append(res)
        if not res.get("ok") or res.get("assertOk") is False:
            ok_all = False
            if cfg.get("stopOnAssertFail") and res.get("assertOk") is False:
                log("      ⏹ 断言失败，按配置停止该场景后续步骤")
                break
    return {"time": now_str(), "ok": ok_all, "steps": steps_out,
            "vars": json.dumps(variables, ensure_ascii=False, default=str)}


def execute(case_ids=None, scenario_ids=None, run_all=False, log=None):
    log = log or (lambda s: None)
    cfg = STATE["config"]
    if run_all:
        cases = [c for c in STATE["cases"] if c.get("enabled", True)]
        scenes = list(STATE["scenarios"])
    else:
        cases = [c for c in STATE["cases"] if c["id"] in (case_ids or [])]
        scenes = [s for s in STATE["scenarios"] if s["id"] in (scenario_ids or [])]

    variables = dict(STATE.get("vars") or {})
    for c in cases:
        prev = (c.get("results") or [{}])[0].get("asserts") if c.get("results") else None
        push_result(c, run_case(c, cfg, variables, log, prev))
    STATE["vars"] = {k: v for k, v in variables.items()
                     if isinstance(v, (str, int, float, bool, type(None)))}
    for sc in scenes:
        sc.setdefault("results", []).insert(0, run_scenario(sc, cfg, log))
        del sc["results"][MAX_HISTORY:]
    save_state()

    okc = sum(1 for c in cases if c.get("results") and c["results"][0].get("ok"))
    oks = sum(1 for s in scenes if s.get("results") and s["results"][0].get("ok"))
    a_pass = a_fail = a_warn = 0
    for c in cases:
        if c.get("results"):
            f, w, p = assert_stats(c["results"][0].get("asserts"))
            a_fail += f; a_warn += w; a_pass += p
    for sc in scenes:
        if sc.get("results"):
            for st in sc["results"][0].get("steps", []):
                f, w, p = assert_stats(st.get("asserts"))
                a_fail += f; a_warn += w; a_pass += p
    msg = "执行完成：接口 %d 个（成功 %d），场景 %d 个（成功 %d）" % (
        len(cases), okc, len(scenes), oks)
    if a_pass or a_fail or a_warn:
        msg += "；断言 通过 %d / 失败 %d / 警告 %d" % (a_pass, a_fail, a_warn)
    return msg


# ---------------------------------------------------------------- 清空
def clear_results(target, obj_id=None, index="all"):
    def clear_one(obj):
        rs = obj.get("results", [])
        if index == "all" or index is None:
            obj["results"] = []
        else:
            i = int(index)
            if 0 <= i < len(rs):
                rs.pop(i)

    if target == "all":
        for c in STATE["cases"]:
            clear_one(c)
        for s in STATE["scenarios"]:
            clear_one(s)
        STATE["vars"] = {}
    elif target == "case":
        c = find_case(obj_id)
        if c:
            clear_one(c)
    elif target == "scenario":
        s = find_scene(obj_id)
        if s:
            clear_one(s)
    save_state()


# ---------------------------------------------------------------- 导入导出
def import_excel(path, mode="replace"):
    parsed = excel_io.read_workbook(path)
    if mode == "append":
        exist = {c["id"] for c in STATE["cases"]}
        for c in parsed["cases"]:
            while c["id"] in exist:
                c["id"] += "_1"
            exist.add(c["id"])
            STATE["cases"].append(c)
        STATE["scenarios"].extend(parsed["scenarios"])
    else:
        STATE["cases"] = parsed["cases"]
        STATE["scenarios"] = parsed["scenarios"]
        if parsed["config"]:
            STATE["config"].update(parsed["config"])
    STATE["fileName"] = os.path.basename(path)
    save_state()
    return "导入成功：%d 个接口用例，%d 个组合场景" % (
        len(parsed["cases"]), len(parsed["scenarios"]))


def export_excel(path):
    excel_io.write_workbook(path, STATE)
    return path


def build_template(path):
    return excel_io.build_template(path)


def probe_extract(case, extract_text, use_cached=True, log=None):
    """调试用：拿一份响应，逐条检查提取配置能否取到值。

    返回 (响应文本, 来源说明, 行结果列表)。本函数不会写入执行历史。
    """
    log = log or (lambda s: None)
    source = ""
    resp_text = None
    if use_cached:
        for r in (case.get("results") or []):
            if r.get("imported"):
                continue
            if r.get("respBody"):
                resp_text = r["respBody"]
                source = "最近一次执行结果（%s）" % r.get("time", "")
                break
    if resp_text is None:
        variables = dict(STATE.get("vars") or {})
        res = run_case(copy.deepcopy(case), STATE["config"], variables, log)
        resp_text = res.get("respBody") or res.get("error") or ""
        source = "刚刚发起的请求（状态码 %s，未计入结果历史）" % res.get("status", "-")

    try:
        parsed = json.loads(resp_text)
    except Exception:
        parsed = None
    return resp_text, source, evaluate_extract(parsed, extract_text)


def evaluate_extract(parsed, extract_text):
    """对已解析的响应逐行计算提取结果。parsed 为 None 表示响应不是 JSON。"""
    rows = []
    for item in parse_extract_detail(extract_text):
        row = {"line": item["line"], "var": item["var"], "path": item["path"]}
        if item["error"]:
            row.update({"ok": False, "status": item["error"], "value": ""})
        elif parsed is None:
            row.update({"ok": False, "status": "响应不是合法 JSON，无法按路径提取", "value": ""})
        else:
            v = json_path_ex(parsed, item["path"])
            if v is MISSING:
                row.update({"ok": False, "status": "✗ 没取到：路径在响应里不存在", "value": ""})
            elif v is None:
                row.update({"ok": True, "status": "△ 取到了，但值是 null", "value": "null"})
            elif isinstance(v, list) and not v:
                row.update({"ok": True, "status": "△ 取到了，但是空数组", "value": "[]"})
            else:
                txt = v if isinstance(v, str) else json.dumps(v, ensure_ascii=False)
                extra = "（数组，共 %d 项）" % len(v) if isinstance(v, list) else ""
                row.update({"ok": True, "status": "✓ 取到了" + extra, "value": txt})
        rows.append(row)
    if not rows:
        rows.append({"line": 0, "var": "", "path": "", "ok": False,
                     "status": "还没有填写提取配置（在下方响应里点一个值即可自动生成）", "value": ""})
    return rows


_IDENT = re.compile(r"^[^.\[\]\s'\"]+$")


def build_path(keys):
    """由键序列生成路径，例如 ['data','list',0,'id'] -> $.data.list[0].id；'*' 表示整列。"""
    p = "$"
    for k in keys:
        if k == "*":
            p += "[*]"
        elif isinstance(k, int):
            p += "[%d]" % k
        elif _IDENT.match(str(k)):
            p += "." + str(k)
        else:
            p += '["%s"]' % str(k).replace('"', "")
    return p


# ---------------------------------------------------------------- 复制用例
def next_copy_name(name, existing):
    """xxx -> xxx_copy1；已存在 _copy1 则 _copy2，以此类推。"""
    base = re.sub(r"_copy\d+$", "", name or "")
    pat = re.compile(re.escape(base) + r"_copy(\d+)$")
    mx = 0
    for n in existing:
        m = pat.match(n or "")
        if m:
            mx = max(mx, int(m.group(1)))
    return "%s_copy%d" % (base, mx + 1)


def duplicate_case(cid):
    src = find_case(cid)
    if not src:
        return None
    new = copy.deepcopy(src)
    new["id"] = next_case_id()
    new["results"] = []
    base = src.get("caseName") or src.get("name") or src["id"]
    new["caseName"] = next_copy_name(base, [c.get("caseName", "") for c in STATE["cases"]])
    # 插到原用例及其已有副本的后面
    root = re.sub(r"_copy\d+$", "", base)
    pat = re.compile(re.escape(root) + r"(_copy\d+)?$")
    pos = STATE["cases"].index(src)
    for i, c in enumerate(STATE["cases"]):
        if i > pos and pat.match(c.get("caseName", "") or ""):
            pos = i
    STATE["cases"].insert(pos + 1, new)
    save_state()
    return new


# ---------------------------------------------------------------- cURL 解析
_CURL_ARG_OPTS = {  # 需要跟一个参数、但我们不关心的选项
    "-o", "--output", "-m", "--max-time", "--connect-timeout", "-x", "--proxy",
    "--cacert", "--capath", "--cert", "-E", "--key", "-w", "--write-out", "--retry",
    "--retry-delay", "--retry-max-time", "-T", "--upload-file", "--resolve",
    "--limit-rate", "-c", "--cookie-jar", "-r", "--range", "-K", "--config",
    "--interface", "--proxy-user", "-U", "--max-redirs", "--dns-servers",
    "--connect-to", "--unix-socket", "--cert-type", "--key-type", "--pass",
}
_ALWAYS_DROP = {"content-length", "host", "accept-encoding"}
_NOISE_EXACT = {"priority", "upgrade-insecure-requests", "dnt", "sec-gpc"}


def _curl_tokenize(text):
    """兼容 bash（反斜杠续行、单引号、双引号、$'...'）的分词器。"""
    tokens, cur, i, n = [], None, 0, len(text)
    esc_map = {"n": "\n", "t": "\t", "r": "\r", "\\": "\\", "'": "'", '"': '"'}
    while i < n:
        c = text[i]
        if c == "\\" and i + 1 < n and text[i + 1] in "\r\n":   # 续行
            i += 2
            if i < n and text[i - 1] == "\r" and text[i] == "\n":
                i += 1
            continue
        if c in " \t\r\n":
            if cur is not None:
                tokens.append(cur)
                cur = None
            i += 1
            continue
        if cur is None:
            cur = ""
        if c == "\\":
            if i + 1 < n:
                cur += text[i + 1]
            i += 2
        elif c == "$" and i + 1 < n and text[i + 1] == "'":
            i += 2
            while i < n and text[i] != "'":
                if text[i] == "\\" and i + 1 < n:
                    e = text[i + 1]
                    cur += esc_map.get(e, "\\" + e)
                    i += 2
                else:
                    cur += text[i]
                    i += 1
            i += 1
        elif c == "'":
            j = text.find("'", i + 1)
            j = n if j < 0 else j
            cur += text[i + 1:j]
            i = j + 1
        elif c == '"':
            i += 1
            while i < n and text[i] != '"':
                if text[i] == "\\" and i + 1 < n and text[i + 1] in '"\\$`':
                    cur += text[i + 1]
                    i += 2
                else:
                    cur += text[i]
                    i += 1
            i += 1
        else:
            cur += c
            i += 1
    if cur is not None:
        tokens.append(cur)
    return tokens


def _preprocess_cmd_style(text):
    """浏览器「复制为 cURL (cmd)」使用 ^ 转义，先还原成普通写法。"""
    if "^\"" in text or re.search(r"\^\s*\r?\n", text):
        text = re.sub(r"\^\s*\r?\n", " ", text)
        text = re.sub(r"\^(.)", r"\1", text)
    return text


def parse_curl(text, base_url="", to_relative=True, slim=True):
    """解析一条或多条 cURL 命令，返回 (用例列表, 提示信息列表)。"""
    from urllib.parse import urlparse, quote

    warnings = []
    text = (text or "").strip()
    if not text:
        return [], ["没有内容可解析"]
    low = text.lower()
    if "invoke-webrequest" in low or "invoke-restmethod" in low:
        return [], ["这是 PowerShell 格式，请在浏览器里选「复制为 cURL (bash)」或「复制为 cURL (cmd)」"]
    tokens = _curl_tokenize(_preprocess_cmd_style(text))

    groups, cur = [], None
    for t in tokens:
        if t.lower() in ("curl", "curl.exe"):
            cur = []
            groups.append(cur)
        elif cur is not None:
            cur.append(t)
    if not groups:
        return [], ["没有找到 curl 命令，请确认粘贴的内容以 curl 开头"]

    cases = []
    for gi, args in enumerate(groups, start=1):
        tag = "第%d条：" % gi if len(groups) > 1 else ""
        method, url, headers, data_parts = None, None, [], []
        force_get, is_json_flag = False, False
        i = 0
        while i < len(args):
            a = args[i]
            val = None
            if a.startswith("--") and "=" in a:
                a, val = a.split("=", 1)

            def take():
                nonlocal i
                if val is not None:
                    return val
                i += 1
                return args[i] if i < len(args) else ""

            if a in ("-X", "--request"):
                method = take().upper()
            elif a.startswith("-X") and len(a) > 2:
                method = a[2:].upper()
            elif a in ("-H", "--header"):
                headers.append(take())
            elif a in ("-d", "--data", "--data-raw", "--data-binary", "--data-ascii"):
                d = take()
                if d.startswith("@"):
                    warnings.append(tag + "Body 引用了本地文件 %s，无法读取，请手动粘贴内容" % d)
                data_parts.append(d)
            elif a == "--data-urlencode":
                d = take()
                if "=" in d:
                    k, v = d.split("=", 1)
                    data_parts.append("%s=%s" % (k, quote(v, safe="")))
                else:
                    data_parts.append(quote(d, safe=""))
            elif a == "--json":
                data_parts.append(take())
                is_json_flag = True
            elif a in ("-u", "--user"):
                import base64
                token = base64.b64encode(take().encode("utf-8")).decode("ascii")
                headers.append("Authorization: Basic " + token)
            elif a in ("-b", "--cookie"):
                ck = take()
                if "=" in ck:
                    headers.append("Cookie: " + ck)
                else:
                    warnings.append(tag + "Cookie 引用了文件 %s，已忽略" % ck)
            elif a in ("-A", "--user-agent"):
                headers.append("User-Agent: " + take())
            elif a in ("-e", "--referer"):
                headers.append("Referer: " + take())
            elif a in ("-G", "--get"):
                force_get = True
            elif a == "--url":
                url = take()
            elif a in ("-F", "--form", "--form-string"):
                take()
                warnings.append(tag + "检测到 -F 表单/文件上传，本工具暂不支持 multipart，已跳过该字段")
            elif a in _CURL_ARG_OPTS:
                take()
            elif a.startswith("-"):
                pass  # -k -L -s -v -i --compressed --insecure --location 等开关，忽略
            elif url is None:
                url = a
            i += 1

        if not url:
            warnings.append(tag + "没有找到请求地址 URL，已跳过")
            continue
        if not re.match(r"^https?://", url, re.I):
            url = "http://" + url

        body = "&".join(data_parts)
        if force_get and body:
            url += ("&" if "?" in url else "?") + body
            body = ""
        if not method:
            method = "POST" if body else "GET"

        # 请求头整理
        hdr_lines, has_ct = [], False
        dropped = 0
        for h in headers:
            if ":" not in h:
                continue
            k, v = h.split(":", 1)
            k, v = k.strip(), v.strip()
            kl = k.lower()
            if kl in _ALWAYS_DROP:
                continue
            if slim and (kl.startswith("sec-ch-") or kl.startswith("sec-fetch-") or kl in _NOISE_EXACT):
                dropped += 1
                continue
            if kl == "content-type":
                has_ct = True
            hdr_lines.append("%s: %s" % (k, v))
        if is_json_flag and not has_ct:
            hdr_lines.append("Content-Type: application/json")
            has_ct = True

        # Body：JSON 美化；非 JSON 且没有 Content-Type 时按 curl 默认表单格式
        if body:
            try:
                body = json.dumps(json.loads(body), ensure_ascii=False, indent=2)
            except Exception:
                if not has_ct:
                    hdr_lines.append("Content-Type: application/x-www-form-urlencoded")

        # URL 转相对路径
        full_url = url
        b = (base_url or "").strip().rstrip("/")
        if to_relative and b and url.lower().startswith(b.lower()):
            rest = url[len(b):]
            if rest == "" or rest[0] in "/?":
                url = rest if rest.startswith("/") else "/" + rest

        path = urlparse(full_url).path
        segs = [x for x in path.split("/") if x]
        name = "/".join(segs[-2:]) if segs else (urlparse(full_url).netloc or "cURL导入接口")

        url, params = split_query(url)
        cases.append({
            "id": "", "name": name, "caseName": "cURL导入", "url": url, "method": method,
            "headers": "\n".join(hdr_lines), "body": body, "params": params, "memos": {},
            "enabled": True, "extract": "", "results": [], "_dropped": dropped,
        })
    return cases, warnings


def next_case_id():
    n = 1
    ids = {c["id"] for c in STATE["cases"]}
    while "CASE_%03d" % n in ids:
        n += 1
    return "CASE_%03d" % n


def next_scene_id():
    n = 1
    ids = {s["id"] for s in STATE["scenarios"]}
    while "SC_%03d" % n in ids:
        n += 1
    return "SC_%03d" % n
