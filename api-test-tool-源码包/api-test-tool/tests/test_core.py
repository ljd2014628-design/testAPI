# -*- coding: utf-8 -*-
"""核心逻辑自测：不依赖网络和界面，直接 `python tests/test_core.py` 运行。

覆盖：JSON 路径解析、变量替换、请求头合并优先级、URL 参数拼接、
cURL 解析、断言比较器、Excel 模板与导入导出往返。
"""

import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))

import core
import excel_io

PASS = FAIL = 0


def check(name, got, want):
    global PASS, FAIL
    if got == want:
        PASS += 1
    else:
        FAIL += 1
        print("  [FAIL] %s\n         实际: %r\n         期望: %r" % (name, got, want))


def ok(name, cond):
    check(name, bool(cond), True)


# ---------------------------------------------------------------- JSON 路径
def test_json_path():
    data = {"code": 0, "data": {"total": 2, "nil": None,
                                "list": [{"id": 1, "no": "A"}, {"id": 2, "no": "B"}]}}
    check("取单值", core.json_path(data, "$.data.total"), 2)
    check("取数组下标", core.json_path(data, "$.data.list[0].no"), "A")
    check("取整列", core.json_path(data, "$.data.list[*].id"), [1, 2])
    check("数组省略[*]", core.json_path(data, "$.data.list.no"), ["A", "B"])
    check("null 值", core.json_path(data, "$.data.nil"), None)
    ok("字段不存在返回 MISSING", core.json_path_ex(data, "$.data.nope") is core.MISSING)
    check("生成路径", core.build_path(["data", "list", 0, "id"]), "$.data.list[0].id")
    check("生成整列路径", core.build_path(["data", "list", "*", "id"]), "$.data.list[*].id")


# ---------------------------------------------------------------- 变量 / 请求头
def test_vars_and_headers():
    check("变量替换", core.render("/q/${id}?b=${b}", {"id": 7, "b": "x"}), "/q/7?b=x")
    check("未定义变量保持原样", core.render("/q/${no}", {}), "/q/${no}")

    cfg = {"headers": [{"key": "Content-Type", "value": "application/json", "enabled": True},
                       {"key": "brandCode", "value": "K001", "enabled": True},
                       {"key": "Off", "value": "x", "enabled": False}]}
    merged = core.merge_headers(cfg, {"headers": '{"content-type":"text/plain","tenant":"T1"}'}, {})
    check("接口级覆盖公共（不分大小写）", merged.get("content-type"), "text/plain")
    ok("公共头保留", merged.get("brandCode") == "K001")
    ok("未启用的公共头不发送", "Off" not in merged)
    ok("接口级新增头", merged.get("tenant") == "T1")


# ---------------------------------------------------------------- URL 参数
def test_params():
    base, params = core.split_query("/get_list?a=&b=1&c=2")
    check("拆出 URL", base, "/get_list")
    check("拆出参数个数", len(params), 3)
    url = core.build_url({"baseUrl": "https://x.com"}, "/get_list", {"p": "13800138000"},
                         [{"key": "a", "value": "", "on": True},
                          {"key": "m", "value": "${p}", "on": True},
                          {"key": "skip", "value": "1", "on": False}])
    check("拼回 URL（停用的不发、变量已替换）", url,
          "https://x.com/get_list?a=&m=13800138000")

    body = {"orderNo": "LS1", "lines": [{"qty": 1}]}
    check("字段摊平", core.flatten_fields(body), [("orderNo", "LS1"), ("lines[0].qty", 1)])
    core.set_field(body, "lines[0].qty", 5)
    check("按路径写回", body["lines"][0]["qty"], 5)


# ---------------------------------------------------------------- cURL
def test_curl():
    bash = ("curl 'https://api.test.com/coupon/list?page=1&size=20' "
            "-H 'authorization: Bearer abc' -H 'content-type: application/json' "
            "-H 'sec-ch-ua: \"Chromium\"' -H 'accept-encoding: gzip, br' "
            "--data-raw '{\"brandCode\":\"K001\",\"名称\":\"中文\"}'")
    cases, warns = core.parse_curl(bash, base_url="https://api.test.com")
    c = cases[0]
    check("cURL 条数", len(cases), 1)
    check("方法", c["method"], "POST")
    check("URL 转相对路径", c["url"], "/coupon/list")
    check("查询参数拆表", [p["key"] for p in c["params"]], ["page", "size"])
    ok("去掉浏览器噪声头", "sec-ch-ua" not in c["headers"])
    ok("去掉 accept-encoding", "accept-encoding" not in c["headers"].lower())
    ok("保留 authorization", "authorization" in c["headers"].lower())
    ok("Body 中文正常", "中文" in c["body"])
    check("无警告", warns, [])

    cmd = ('curl ^"https://api.test.com/a^" ^\n  -H ^"accept: application/json^" ^\n'
           '  --data-raw ^"^{^\\^"k^\\^":1^}^"')
    cases2, _ = core.parse_curl(cmd)
    check("cmd 格式方法", cases2[0]["method"], "POST")
    check("cmd 格式 Body", json.loads(cases2[0]["body"]), {"k": 1})

    multi, _ = core.parse_curl("curl https://x.com/a\ncurl -X DELETE https://x.com/b")
    check("一次解析多条", [m["method"] for m in multi], ["GET", "DELETE"])


# ---------------------------------------------------------------- 断言
def test_asserts():
    resp = {"code": 0, "msg": "success",
            "data": {"total": 3, "remain": 9, "amount": 99.5, "name": "优惠券A",
                     "flag": True, "nil": None,
                     "list": [{"id": 3, "no": "CP0003"}, {"id": 2, "no": "CP0002"},
                              {"id": 1, "no": "CP0001"}]}}
    text = json.dumps(resp, ensure_ascii=False)
    headers = {"Content-Type": "application/json;charset=utf-8"}

    table = [
        ("$.code", "等于", "0", True), ("$.code", "等于", "1", False),
        ("$.data.name", "文本等于", "优惠券A", True),
        ("$.data.amount", "约等于", "100±1", True),
        ("$.data.amount", "约等于", "100±0.1", False),
        ("$.data.total", "介于", "1~10", True),
        ("$.data.name", "包含", "优惠", True),
        ("$.data.name", "开头是", "优惠", True),
        ("$.data.list[0].no", "正则匹配", r"^CP\d{4}$", True),
        ("$.data.nope", "不存在", "", True),
        ("$.data.nil", "为null", "", True),
        ("$.data.total", "类型是", "数字", True),
        ("$.data.flag", "类型是", "数字", False),
        ("$.data", "必含字段", "total,remain,list", True),
        ("$.data", "必含字段", "total,nothing", False),
        ("$.data.list", "长度等于", "3", True),
        ("$.msg", "属于", "success,ok", True),
        ("$.data.list[*].no", "包含元素", "CP0002", True),
        ("$.data.list", "每项都含字段", "id,no", True),
        ("$.data.list[*].id", "无重复", "", True),
        ("$.data.list", "降序", "id", True),
        ("$.data.list", "升序", "id", False),
        ("状态码", "等于", "200", True),
        ("耗时ms", "小于", "2000", True),
        ("响应头.Content-Type", "包含", "application/json", True),
        ("响应文本", "是有效JSON", "", True),
    ]
    for target, op, expected, want in table:
        row = core.eval_asserts([{"on": True, "target": target, "op": op, "expected": expected}],
                                resp, text, 200, 30, {}, None, headers)[0]
        check("断言 %s [%s] %s" % (target, op, expected), row["ok"], want)

    # 领券场景：断言写死一次，连续执行都能通过
    a = [{"on": True, "target": "$.data.remain", "op": "较上次减少", "expected": "1"}]
    prev = None
    got = []
    for remain in (10, 9, 8, 6):
        rows = core.eval_asserts(a, {"data": {"remain": remain}}, "", 200, 10, {}, prev)
        got.append(rows[0]["ok"])
        prev = rows
    check("较上次减少（首次记基线，之后逐次校验）", got, [True, True, True, False])

    # 期望值支持变量与算式
    row = core.eval_asserts([{"on": True, "target": "$.data.remain", "op": "等于",
                              "expected": "${before}-1"}],
                            {"data": {"remain": 9}}, "", 200, 10, {"before": 10}, None)[0]
    ok("期望值算式 ${before}-1", row["ok"])

    # 仅警告不影响整体结论
    rows = core.eval_asserts([{"on": True, "target": "$.code", "op": "等于", "expected": "0"},
                              {"on": True, "target": "耗时ms", "op": "小于", "expected": "1",
                               "level": "仅警告"}], resp, text, 200, 30, {}, None, headers)
    check("统计(失败,警告,通过)", core.assert_stats(rows), (0, 1, 1))


# ---------------------------------------------------------------- Excel
def test_excel():
    tmp = tempfile.mkdtemp()
    tpl = os.path.join(tmp, "template.xlsx")
    excel_io.build_template(tpl)
    parsed = excel_io.read_workbook(tpl)
    ok("模板能生成并解析", len(parsed["cases"]) >= 3)
    ok("模板带组合场景示例", len(parsed["scenarios"]) >= 1)

    saved = dict(core.STATE)
    core.STATE["cases"] = [{
        "id": "CASE_001", "name": "领券", "caseName": "领一张", "url": "/coupon/receive",
        "method": "POST", "headers": "token: abc", "body": '{"batchNo":"B001"}',
        "params": [{"key": "ch", "value": "app", "memo": "渠道", "on": True},
                   {"key": "dbg", "value": "1", "memo": "", "on": False}],
        "memos": {"batchNo": "批次号"},
        "asserts": [{"on": True, "target": "$.code", "op": "等于", "expected": "0",
                     "level": "必须通过", "memo": "业务码"},
                    {"on": True, "target": "耗时ms", "op": "小于", "expected": "2000",
                     "level": "仅警告", "memo": "性能参考"}],
        "enabled": True, "extract": "no=$.data.couponNo", "results": []}]
    core.STATE["scenarios"] = []
    out = os.path.join(tmp, "out.xlsx")
    excel_io.write_workbook(out, core.STATE)
    back = excel_io.read_workbook(out)[0] if isinstance(excel_io.read_workbook(out), list) \
        else excel_io.read_workbook(out)["cases"][0]
    src = core.STATE["cases"][0]
    for key in ("id", "name", "caseName", "url", "method", "headers", "body",
                "extract", "params", "memos", "asserts"):
        check("导出→导入往返 %s" % key, back[key], src[key])
    core.STATE.clear()
    core.STATE.update(saved)


def main():
    print("接口测试工具 · 核心自测")
    for fn in (test_json_path, test_vars_and_headers, test_params,
               test_curl, test_asserts, test_excel):
        print("- %s" % fn.__name__)
        fn()
    print("\n通过 %d 项，失败 %d 项" % (PASS, FAIL))
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
