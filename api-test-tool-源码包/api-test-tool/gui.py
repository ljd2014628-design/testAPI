# -*- coding: utf-8 -*-
"""接口测试工具 · Windows 桌面应用界面（Tkinter）"""

import os
import re
import sys
import json
import queue
import threading
import traceback

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

import core

# 打包成无控制台窗口程序时 sys.stdout/stderr 为 None，任何 print 都会抛异常，这里兜底
class _Null:
    def write(self, *a):
        pass

    def flush(self):
        pass


if sys.stdout is None:
    sys.stdout = _Null()
if sys.stderr is None:
    sys.stderr = _Null()

APP_TITLE = "接口测试工具"
CHK_ON, CHK_OFF = "☑", "☐"

BG = "#f4f6fa"
BRAND = "#2f5597"
OKC = "#15803d"
ERRC = "#b91c1c"


# ====================================================================== 通用
def center(win, w, h):
    win.update_idletasks()
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    win.geometry("%dx%d+%d+%d" % (w, h, max(0, (sw - w) // 2), max(0, (sh - h) // 3)))


PREVIEW_LIMIT = 2000


def _clip(text, full):
    """长文本在摘要模式下截断，并提示可展开。"""
    text = "" if text is None else str(text)
    if full or len(text) <= PREVIEW_LIMIT:
        return text
    return (text[:PREVIEW_LIMIT]
            + "\n\n……（此处省略 %d 个字符，点窗口上方的「展开全部内容」查看完整内容）"
            % (len(text) - PREVIEW_LIMIT))


def fmt_result(r, indent="", full=False):
    """把一次执行结果格式化成可读文本。full=False 时长内容折叠。"""
    if r.get("imported"):
        return indent + _clip(r.get("text") or "", full)
    lines = []
    flag = "成功" if r.get("ok") else "失败"
    lines.append("%s【%s】 %s   状态码：%s   耗时：%s ms" % (
        indent, r.get("time", ""), flag, r.get("status", "-"), r.get("elapsed", "-")))
    lines.append("%s请求：%s %s" % (indent, r.get("method", ""), r.get("url", "")))
    if r.get("reqHeaders"):
        try:
            hd = json.loads(r["reqHeaders"])
            htxt = "\n".join("    %s: %s" % (k, v) for k, v in hd.items())
            lines.append("%s请求头（%d 个）：\n%s" % (indent, len(hd), _clip(htxt, full)))
        except Exception:
            lines.append("%s请求头：%s" % (indent, _clip(r["reqHeaders"], full)))
    if r.get("reqBody"):
        lines.append("%s请求Body：\n%s" % (indent, _clip(r["reqBody"], full)))
    if r.get("vars"):
        lines.append("%s提取变量：%s" % (indent, r["vars"]))
    rows = r.get("asserts") or []
    if rows:
        fail = [x for x in rows if not x.get("ok") and x.get("level") != "仅警告"]
        warn = [x for x in rows if not x.get("ok") and x.get("level") == "仅警告"]
        head = "全部通过（%d 条）" % len(rows) if not fail and not warn else (
            "不符合预期（%d 条未通过%s）" % (len(fail), "，%d 条警告" % len(warn) if warn else "")
            if fail else "通过，但有 %d 条警告" % len(warn))
        lines.append("%s断言：%s" % (indent, head))
        for x in rows:
            mark = "√" if x.get("ok") else ("⚠" if x.get("level") == "仅警告" else "×")
            lines.append("%s    %s %s [%s] %s —— %s" % (
                indent, mark, x.get("target", ""), x.get("op", ""),
                x.get("expected", ""), x.get("msg", "")))
    if r.get("error"):
        lines.append("%s异常：%s" % (indent, r["error"]))
    else:
        body = r.get("respBody", "") or ""
        head = "%s响应结果：" % indent
        if not full and len(body) > PREVIEW_LIMIT:
            head = "%s响应结果（共 %d 个字符，已折叠）：" % (indent, len(body))
        lines.append("%s\n%s" % (head, _clip(body, full)))
    return "\n".join(lines)


def brief(results):
    if not results:
        return "未执行"
    r = results[0]
    if r.get("imported"):
        return "导入结果"
    txt = "%s %s" % ("成功" if r.get("ok") else "失败", r.get("status", ""))
    rows = r.get("asserts") or []
    if rows:
        fail = sum(1 for x in rows if not x.get("ok") and x.get("level") != "仅警告")
        warn = sum(1 for x in rows if not x.get("ok") and x.get("level") == "仅警告")
        if fail:
            txt += " · 断言不通过%d条" % fail
        elif warn:
            txt += " · 断言通过(警告%d)" % warn
        else:
            txt += " · 断言通过"
    return "%s (%s次)" % (txt, len(results))


class TextDialog(tk.Toplevel):
    """只读长文本查看窗口。pages = [(页名, 摘要文本, 完整文本), ...]"""

    def __init__(self, master, title, pages):
        super().__init__(master)
        self.title(title)
        self.transient(master)
        center(self, 1000, 660)
        self.pages = pages
        self.expanded = False

        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=8, pady=(8, 0))
        self.btn_toggle = ttk.Button(bar, text="展开全部内容", style="Brand.TButton",
                                     command=self.toggle)
        self.btn_toggle.pack(side="left")
        ttk.Button(bar, text="把当前页完整内容另存为 txt",
                   command=self.save_txt).pack(side="left", padx=6)
        self.lb_tip = ttk.Label(bar, foreground="#6b7280",
                                text="响应过长时默认折叠，点左边按钮可看全文")
        self.lb_tip.pack(side="left", padx=8)

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.texts = []
        for name, short, _full in pages:
            f = ttk.Frame(self.nb)
            self.nb.add(f, text=name)
            st = ScrolledText(f, wrap="word", font=("Consolas", 10))
            st.pack(fill="both", expand=True)
            st.insert("1.0", short)
            st.configure(state="disabled")
            self.texts.append(st)

        ttk.Button(self, text="关闭", command=self.destroy).pack(pady=(0, 8))
        self.grab_set()

    def toggle(self):
        self.expanded = not self.expanded
        self.btn_toggle.config(text="仅显示摘要" if self.expanded else "展开全部内容")
        for st, (_n, short, full) in zip(self.texts, self.pages):
            st.configure(state="normal")
            st.delete("1.0", "end")
            st.insert("1.0", full if self.expanded else short)
            st.configure(state="disabled")

    def save_txt(self):
        i = self.nb.index("current")
        path = filedialog.asksaveasfilename(
            title="保存完整内容", defaultextension=".txt",
            initialfile="接口响应_%s.txt" % self.pages[i][0],
            filetypes=[("文本文件", "*.txt")], parent=self)
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.pages[i][2])
            messagebox.showinfo("完成", "已保存到：\n%s" % path, parent=self)
        except Exception as e:
            messagebox.showerror("保存失败", str(e), parent=self)


class ClearDialog(tk.Toplevel):
    """清空结果：存在多次结果时让用户选第几次。"""

    def __init__(self, master, label, count, on_ok):
        super().__init__(master)
        self.title("清空执行结果")
        self.transient(master)
        center(self, 430, 240 if count > 1 else 150)
        self.on_ok = on_ok
        ttk.Label(self, text="将清空：%s" % label, font=("Microsoft YaHei", 10, "bold")
                  ).pack(anchor="w", padx=16, pady=(14, 6))
        self.var = tk.StringVar(value="all")
        if count > 1:
            ttk.Label(self, text="检测到有 %d 次执行结果，请选择要清空哪一次：" % count
                      ).pack(anchor="w", padx=16)
            box = ttk.Frame(self)
            box.pack(fill="both", expand=True, padx=26, pady=4)
            ttk.Radiobutton(box, text="全部 %d 次结果" % count, value="all",
                            variable=self.var).pack(anchor="w")
            for i in range(count):
                ttk.Radiobutton(box, text="第 %d 次%s" % (i + 1, "（最新）" if i == 0 else ""),
                                value=str(i), variable=self.var).pack(anchor="w")
        else:
            ttk.Label(self, text="当前只有 1 次执行结果，将直接清空。").pack(anchor="w", padx=16, pady=6)
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=16, pady=10)
        ttk.Button(bar, text="取消", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(bar, text="确认清空", command=self._ok).pack(side="right")
        self.grab_set()

    def _ok(self):
        v = self.var.get()
        self.on_ok("all" if v == "all" else int(v))
        self.destroy()


def scroll_text(parent, height=6, wrap="none"):
    """带横向 + 纵向滚动条的文本框，返回 (外层Frame, Text)。"""
    frm = ttk.Frame(parent)
    t = tk.Text(frm, height=height, wrap=wrap, font=("Consolas", 10), undo=True,
                relief="solid", borderwidth=1)
    ys = ttk.Scrollbar(frm, orient="vertical", command=t.yview)
    xs = ttk.Scrollbar(frm, orient="horizontal", command=t.xview)
    t.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
    t.grid(row=0, column=0, sticky="nsew")
    ys.grid(row=0, column=1, sticky="ns")
    if wrap == "none":
        xs.grid(row=1, column=0, sticky="ew")
    frm.rowconfigure(0, weight=1)
    frm.columnconfigure(0, weight=1)
    return frm, t


def fit_height(win, want):
    return max(480, min(want, win.winfo_screenheight() - 80))


def inline_edit(tv, iid, col, on_commit, on_next=None):
    """在 Treeview 单元格上盖输入框：回车保存，Esc 取消，Tab 去下一格。"""
    tv.update_idletasks()
    tv.see(iid)
    tv.update_idletasks()
    bbox = tv.bbox(iid, col)
    if not bbox:
        return None
    x, y, w, h = bbox
    ent = ttk.Entry(tv, font=("Microsoft YaHei", 9))
    ent.place(x=x, y=y, width=w, height=h)
    ent.insert(0, tv.set(iid, col))
    ent.select_range(0, "end")
    ent.icursor("end")
    ent.focus_set()
    done = {"v": False}

    def finish(save, jump=False):
        if done["v"]:
            return
        done["v"] = True
        text = ent.get()
        ent.destroy()
        if save:
            on_commit(text.strip())
        if jump and on_next:
            tv.after(10, on_next)

    ent.bind("<Return>", lambda e: finish(True))
    ent.bind("<KP_Enter>", lambda e: finish(True))
    ent.bind("<Escape>", lambda e: finish(False))
    ent.bind("<FocusOut>", lambda e: finish(True))
    ent.bind("<Tab>", lambda e: (finish(True, True), "break")[1])
    return ent


class ParamTable(ttk.Frame):
    """URL 参数表：启用 / 参数名 / 参数值 / 备注（备注不参与请求）。"""

    COLS = (("on", "启用", 46, "center"), ("key", "参数名", 200, "w"),
            ("value", "参数值", 330, "w"), ("memo", "备注（只给自己看，不发送）", 300, "w"))

    def __init__(self, master, on_change=None):
        super().__init__(master)
        self.on_change = on_change
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(0, 3))
        self._btns = []
        for txt, cmd in (("+ 添加一行", self.add_row), ("删除选中", self.del_row),
                         ("↑", lambda: self.move(-1)), ("↓", lambda: self.move(1)),
                         ("粘贴 URL / 参数串解析", self.paste_url), ("清空", self.clear)):
            b = ttk.Button(bar, text=txt, command=cmd)
            b.pack(side="left", padx=2)
            self._btns.append(b)
        self.lb_tip = ttk.Label(bar, foreground="#6b7280",
                                text="  双击单元格修改，回车保存，Tab 到下一格；去掉「启用」这次就不发送该参数")
        self.lb_tip.pack(side="left")
        self.locked = False

        box = ttk.Frame(self)
        box.pack(fill="both", expand=True)
        self.tv = ttk.Treeview(box, columns=[c[0] for c in self.COLS], show="headings", height=6)
        for key, title, w, anchor in self.COLS:
            self.tv.heading(key, text=title)
            self.tv.column(key, width=w, anchor=anchor)
        self.tv.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tv.yview)
        sb.pack(side="right", fill="y")
        self.tv.configure(yscrollcommand=sb.set)
        self.tv.bind("<Button-1>", self._click)
        self.tv.bind("<Double-1>", self._dbl)
        self._seq = 0

    # ---- 数据 ----
    def load(self, params):
        self.tv.delete(*self.tv.get_children())
        for p in params or []:
            self._insert(p)

    def _insert(self, p):
        self._seq += 1
        iid = "row%d" % self._seq
        self.tv.insert("", "end", iid=iid,
                       values=(CHK_ON if p.get("on", True) else CHK_OFF,
                               p.get("key", ""), p.get("value", ""), p.get("memo", "")))
        return iid

    def dump(self):
        out = []
        for iid in self.tv.get_children():
            v = self.tv.item(iid, "values")
            if not str(v[1]).strip():
                continue
            out.append({"on": v[0] == CHK_ON, "key": str(v[1]).strip(),
                        "value": str(v[2]), "memo": str(v[3])})
        return out

    def count(self):
        return len(self.tv.get_children())

    def set_locked(self, flag):
        """Body 里有内容时锁定参数表（两者互斥）。"""
        self.locked = flag
        for b in self._btns:
            b.configure(state="disabled" if flag else "normal")
        self.lb_tip.config(
            text=("  已锁定：当前这个接口用的是请求 Body，参数表不可编辑"
                  if flag else
                  "  双击单元格修改，回车保存，Tab 到下一格；去掉「启用」这次就不发送该参数"),
            foreground=(ERRC if flag else "#6b7280"))

    # ---- 交互 ----
    def _changed(self):
        if self.on_change:
            self.on_change()

    def _click(self, event):
        if self.locked:
            return "break"
        if self.tv.identify_region(event.x, event.y) != "cell":
            return
        if self.tv.identify_column(event.x) != "#1":
            return
        iid = self.tv.identify_row(event.y)
        if not iid:
            return
        cur = self.tv.set(iid, "on")
        self.tv.set(iid, "on", CHK_OFF if cur == CHK_ON else CHK_ON)
        self._changed()
        return "break"

    def _dbl(self, event):
        if self.locked:
            return "break"
        col = self.tv.identify_column(event.x)
        iid = self.tv.identify_row(event.y)
        if not iid or col == "#1":
            return "break"
        self.edit(iid, col)
        return "break"

    def edit(self, iid, col):
        def commit(text):
            self.tv.set(iid, col, text)
            self._changed()

        nxt = None
        if col in ("#2", "#3"):
            nxt = lambda: self.edit(iid, "#%d" % (int(col[1:]) + 1))
        inline_edit(self.tv, iid, col, commit, nxt)

    def add_row(self):
        iid = self._insert({"on": True, "key": "", "value": "", "memo": ""})
        self.tv.selection_set(iid)
        self.after(20, lambda: self.edit(iid, "#2"))

    def del_row(self):
        for iid in self.tv.selection():
            self.tv.delete(iid)
        self._changed()

    def move(self, d):
        sel = self.tv.selection()
        if not sel:
            return
        iid = sel[0]
        idx = self.tv.index(iid) + d
        if 0 <= idx < len(self.tv.get_children()):
            self.tv.move(iid, "", idx)
            self._changed()

    def clear(self):
        self.tv.delete(*self.tv.get_children())
        self._changed()

    def paste_url(self):
        PasteParamDialog(self, self._append_parsed)

    def _append_parsed(self, text):
        text = (text or "").strip()
        if not text:
            return
        if "?" in text:
            text = text.split("?", 1)[1]
        _b, params = core.split_query("?" + text.lstrip("&?"))
        for p in params:
            self._insert(p)
        self._changed()


class PasteParamDialog(tk.Toplevel):
    def __init__(self, master, on_ok):
        super().__init__(master)
        self.title("粘贴 URL 或参数串")
        self.transient(master)
        center(self, 680, 260)
        ttk.Label(self, text="把完整 URL 或 a=1&b=2 形式的参数串粘进来，会自动拆成一行一个参数追加到表里"
                  ).pack(anchor="w", padx=12, pady=(12, 4))
        fr, self.t = scroll_text(self, 6, wrap="word")
        fr.pack(fill="both", expand=True, padx=12)
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=12, pady=10)
        ttk.Button(bar, text="取消", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(bar, text="解析并添加", style="Brand.TButton",
                   command=lambda: (on_ok(self.t.get("1.0", "end")), self.destroy())
                   ).pack(side="right")
        self.grab_set()
        self.t.focus_set()


class FieldTable(ttk.Frame):
    """Body 字段表：字段路径 / 值 / 备注（备注不参与请求）。"""

    def __init__(self, master, get_body, set_body):
        super().__init__(master)
        self.get_body = get_body
        self.set_body = set_body
        self.memos = {}
        self.paths = {}

        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(0, 3))
        ttk.Button(bar, text="重新解析 Body", command=self.reload).pack(side="left", padx=2)
        ttk.Button(bar, text="清空全部备注", command=self.clear_memos).pack(side="left", padx=2)
        self.lb = ttk.Label(bar, foreground="#6b7280",
                            text="  双击「值」可直接改并写回 Body；双击「备注」写注释，备注不会发送")
        self.lb.pack(side="left")

        box = ttk.Frame(self)
        box.pack(fill="both", expand=True)
        cols = ("path", "value", "memo")
        self.tv = ttk.Treeview(box, columns=cols, show="headings", height=6)
        for key, title, w in (("path", "字段路径", 260), ("value", "值", 330),
                              ("memo", "备注（只给自己看，不发送）", 300)):
            self.tv.heading(key, text=title)
            self.tv.column(key, width=w, anchor="w")
        self.tv.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tv.yview)
        sb.pack(side="right", fill="y")
        self.tv.configure(yscrollcommand=sb.set)
        self.tv.bind("<Double-1>", self._dbl)

    def load_memos(self, memos):
        self.memos = dict(memos or {})

    def reload(self):
        self.tv.delete(*self.tv.get_children())
        self.paths = {}
        raw = (self.get_body() or "").strip()
        if not raw:
            self.lb.config(text="  Body 为空，没有可备注的字段")
            return
        try:
            data = json.loads(raw)
        except Exception:
            self.lb.config(text="  Body 不是合法 JSON，无法按字段拆分（备注仍可在 Excel 的「字段备注」列维护）")
            return
        for path, value in core.flatten_fields(data):
            disp = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False)
            iid = self.tv.insert("", "end", values=(path, disp, self.memos.get(path, "")))
            self.paths[iid] = path
        self.lb.config(text="  共 %d 个字段；双击「值」改内容，双击「备注」写注释（备注不发送）"
                            % len(self.paths))

    def dump_memos(self):
        out = {}
        for iid in self.tv.get_children():
            path, _v, memo = self.tv.item(iid, "values")
            if str(memo).strip():
                out[path] = str(memo).strip()
        # 保留当前 Body 里已不存在、但之前写过的备注
        for k, v in self.memos.items():
            out.setdefault(k, v)
        return out

    def clear_memos(self):
        for iid in self.tv.get_children():
            self.tv.set(iid, "memo", "")
        self.memos = {}

    def _dbl(self, event):
        col = self.tv.identify_column(event.x)
        iid = self.tv.identify_row(event.y)
        if not iid or col == "#1":
            return "break"
        if col == "#2":
            inline_edit(self.tv, iid, col, lambda t: self._set_value(iid, t))
        else:
            def commit(t):
                self.tv.set(iid, "memo", t)
                self.memos[self.paths.get(iid, "")] = t
            inline_edit(self.tv, iid, col, commit)
        return "break"

    def _set_value(self, iid, text):
        raw = (self.get_body() or "").strip()
        try:
            data = json.loads(raw)
        except Exception:
            return
        try:
            value = json.loads(text)      # 数字 / true / null 保持原类型
        except Exception:
            value = text
        if core.set_field(data, self.paths.get(iid, ""), value):
            self.set_body(json.dumps(data, ensure_ascii=False, indent=2))
            self.tv.set(iid, "value", text)


def inline_combo(tv, iid, col, values, on_commit):
    tv.update_idletasks()
    tv.see(iid)
    bbox = tv.bbox(iid, col)
    if not bbox:
        return None
    x, y, w, h = bbox
    cb = ttk.Combobox(tv, values=values, state="readonly", font=("Microsoft YaHei", 9))
    cb.place(x=x, y=y, width=max(w, 120), height=h)
    cur = tv.set(iid, col)
    if cur in values:
        cb.current(values.index(cur))
    cb.focus_set()
    done = {"v": False}

    def finish(save):
        if done["v"]:
            return
        done["v"] = True
        val = cb.get()
        cb.destroy()
        if save and val:
            on_commit(val)

    cb.bind("<<ComboboxSelected>>", lambda e: finish(True))
    cb.bind("<Return>", lambda e: finish(True))
    cb.bind("<Escape>", lambda e: finish(False))
    cb.bind("<FocusOut>", lambda e: finish(True))
    return cb


class JsonPickDialog(tk.Toplevel):
    """从接口响应里点选字段，拿到 JSON 路径。"""

    def __init__(self, master, case_provider, on_pick):
        super().__init__(master)
        self.title("从响应里选字段")
        self.transient(master)
        center(self, 900, fit_height(self, 640))
        self.case_provider = case_provider
        self.on_pick = on_pick
        self.node_keys = {}
        self.busy = False

        top = ttk.Frame(self)
        top.pack(fill="x", padx=12, pady=(10, 4))
        self.src = tk.StringVar(value="cached")
        ttk.Radiobutton(top, text="最近一次执行结果", value="cached", variable=self.src).pack(side="left")
        ttk.Radiobutton(top, text="立即请求一次", value="live", variable=self.src).pack(side="left", padx=8)
        self.btn = ttk.Button(top, text="获取响应", command=self.fetch)
        self.btn.pack(side="left", padx=8)
        self.lb = ttk.Label(top, foreground="#6b7280", text="")
        self.lb.pack(side="left")

        box = ttk.Frame(self)
        box.pack(fill="both", expand=True, padx=12)
        self.jt = ttk.Treeview(box, columns=("value", "type"), show="tree headings")
        self.jt.heading("#0", text="字段")
        self.jt.heading("value", text="值")
        self.jt.heading("type", text="类型")
        self.jt.column("#0", width=280, stretch=False)
        self.jt.column("value", width=420)
        self.jt.column("type", width=70, anchor="center", stretch=False)
        ys = ttk.Scrollbar(box, orient="vertical", command=self.jt.yview)
        self.jt.configure(yscrollcommand=ys.set)
        self.jt.pack(side="left", fill="both", expand=True)
        ys.pack(side="right", fill="y")
        self.jt.bind("<<TreeviewSelect>>", self._pick)
        self.jt.bind("<Double-1>", lambda e: self.use())

        row = ttk.Frame(self)
        row.pack(side="bottom", fill="x", padx=12, pady=8)
        ttk.Label(row, text="路径").pack(side="left")
        self.e_path = ttk.Entry(row, font=("Consolas", 10))
        self.e_path.pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="取整列 [*]", command=self.whole).pack(side="left", padx=2)
        ttk.Button(row, text="使用这个路径", style="Brand.TButton", command=self.use).pack(side="left")
        ttk.Button(row, text="关闭", command=self.destroy).pack(side="left", padx=4)
        self.grab_set()
        self.after(200, self.fetch)

    def fetch(self):
        if self.busy:
            return
        case = self.case_provider()
        if not case:
            return
        self.busy = True
        self.btn.config(text="获取中…", state="disabled")
        use_cached = self.src.get() == "cached"

        def worker():
            try:
                resp, source, _ = core.probe_extract(case, "", use_cached)
                err = None
            except Exception as e:
                resp, source, err = "", "", str(e)
            self.after(0, done, resp, source, err)

        def done(resp, source, err):
            self.busy = False
            self.btn.config(text="获取响应", state="normal")
            if err:
                self.lb.config(text="获取失败：%s" % err)
                return
            self.lb.config(text=source)
            try:
                parsed = json.loads(resp)
            except Exception:
                parsed = None
            self.jt.delete(*self.jt.get_children())
            self.node_keys = {}
            if parsed is None:
                self.jt.insert("", "end", text="响应不是 JSON，无法点选")
                return
            tname = {str: "字符串", bool: "布尔", int: "数字", float: "数字", type(None): "null"}

            def add(parent, label, value, keys, depth):
                if isinstance(value, dict):
                    iid = self.jt.insert(parent, "end", text=label,
                                         values=("{ %d 个字段 }" % len(value), "对象"), open=depth < 2)
                    self.node_keys[iid] = keys
                    for k, v in value.items():
                        add(iid, str(k), v, keys + [k], depth + 1)
                elif isinstance(value, list):
                    iid = self.jt.insert(parent, "end", text=label,
                                         values=("[ %d 项 ]" % len(value), "数组"), open=depth < 2)
                    self.node_keys[iid] = keys
                    for i, v in enumerate(value[:200]):
                        add(iid, "[%d]" % i, v, keys + [i], depth + 1)
                else:
                    disp = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False)
                    iid = self.jt.insert(parent, "end", text=label,
                                         values=(str(disp)[:200], tname.get(type(value), "")))
                    self.node_keys[iid] = keys

            add("", "$（响应根节点）", parsed, [], 0)

        threading.Thread(target=worker, daemon=True).start()

    def _pick(self, _e=None):
        sel = self.jt.selection()
        if sel and sel[0] in self.node_keys:
            self.e_path.delete(0, "end")
            self.e_path.insert(0, core.build_path(self.node_keys[sel[0]]))

    def whole(self):
        sel = self.jt.selection()
        if not sel or sel[0] not in self.node_keys:
            return
        keys = list(self.node_keys[sel[0]])
        idx = max((i for i, k in enumerate(keys) if isinstance(k, int)), default=None)
        if idx is None:
            messagebox.showinfo("提示", "这个字段不在数组里", parent=self)
            return
        keys[idx] = "*"
        self.e_path.delete(0, "end")
        self.e_path.insert(0, core.build_path(keys))

    def use(self):
        p = self.e_path.get().strip()
        if p:
            self.on_pick(p)
            self.destroy()


class AssertTable(ttk.Frame):
    """断言表：启用 / 目标 / 比较方式 / 期望值 / 备注。"""

    def __init__(self, master, case_provider=None):
        super().__init__(master)
        self.pick_provider = case_provider
        self.case_provider = case_provider
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(0, 3))
        for txt, cmd in (("+ 添加一条", self.add_row), ("删除选中", self.del_row),
                         ("↑", lambda: self.move(-1)), ("↓", lambda: self.move(1)),
                         ("从响应选字段", self.pick_field), ("常用断言", self.add_common),
                         ("▶ 试跑断言", self.test_run), ("清空", self.clear)):
            ttk.Button(bar, text=txt, command=cmd).pack(side="left", padx=2)
        ttk.Label(bar, foreground="#6b7280",
                  text="  双击单元格修改；「较上次…」适合领券这类每跑一次就变的场景，不用每次改期望值"
                  ).pack(side="left")

        box = ttk.Frame(self)
        box.pack(fill="both", expand=True)
        cols = ("on", "target", "op", "expected", "level", "memo")
        self.tv = ttk.Treeview(box, columns=cols, show="headings", height=5)
        for key, title, w, a in (("on", "启用", 46, "center"),
                                 ("target", "目标（JSON路径 / 状态码 / 耗时ms / 响应头.X）", 250, "w"),
                                 ("op", "比较方式", 110, "center"), ("expected", "期望值", 160, "w"),
                                 ("level", "级别", 80, "center"), ("memo", "备注", 180, "w")):
            self.tv.heading(key, text=title)
            self.tv.column(key, width=w, anchor=a)
        self.tv.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tv.yview)
        sb.pack(side="right", fill="y")
        self.tv.configure(yscrollcommand=sb.set)
        self.tv.bind("<Button-1>", self._click)
        self.tv.bind("<Double-1>", self._dbl)
        self._seq = 0

    def load(self, asserts):
        self.tv.delete(*self.tv.get_children())
        for a in asserts or []:
            self._insert(a)

    def _insert(self, a):
        self._seq += 1
        iid = "a%d" % self._seq
        self.tv.insert("", "end", iid=iid,
                       values=(CHK_ON if a.get("on", True) else CHK_OFF, a.get("target", ""),
                               a.get("op", "等于"), a.get("expected", ""),
                               a.get("level", "必须通过"), a.get("memo", "")))
        return iid

    def dump(self):
        out = []
        for iid in self.tv.get_children():
            v = self.tv.item(iid, "values")
            if not str(v[1]).strip():
                continue
            out.append({"on": v[0] == CHK_ON, "target": str(v[1]).strip(),
                        "op": str(v[2]).strip() or "等于", "expected": str(v[3]).strip(),
                        "level": str(v[4]).strip() or "必须通过", "memo": str(v[5]).strip()})
        return out

    def count(self):
        return len(self.tv.get_children())

    def _click(self, event):
        if self.tv.identify_region(event.x, event.y) != "cell":
            return
        if self.tv.identify_column(event.x) != "#1":
            return
        iid = self.tv.identify_row(event.y)
        if not iid:
            return
        cur = self.tv.set(iid, "on")
        self.tv.set(iid, "on", CHK_OFF if cur == CHK_ON else CHK_ON)
        return "break"

    def _dbl(self, event):
        col = self.tv.identify_column(event.x)
        iid = self.tv.identify_row(event.y)
        if not iid or col == "#1":
            return "break"
        self.edit(iid, col)
        return "break"

    def edit(self, iid, col):
        if col == "#3":
            inline_combo(self.tv, iid, col, core.ASSERT_OPS,
                         lambda t: self.tv.set(iid, col, t))
            return
        if col == "#5":
            inline_combo(self.tv, iid, col, core.ASSERT_LEVELS,
                         lambda t: self.tv.set(iid, col, t))
            return
        nxt = None
        if col in ("#2", "#4"):
            nxt = lambda: self.edit(iid, "#%d" % (int(col[1:]) + 1))
        inline_edit(self.tv, iid, col, lambda t: self.tv.set(iid, col, t), nxt)

    def add_row(self):
        iid = self._insert({"on": True, "target": "", "op": "等于", "expected": "",
                            "level": "必须通过", "memo": ""})
        self.tv.selection_set(iid)
        self.after(20, lambda: self.edit(iid, "#2"))

    ASSERT_LIB = [
        ("接口健康检查（最常用）", [
            {"target": "状态码", "op": "等于", "expected": "200", "memo": "HTTP 正常"},
            {"target": "$.code", "op": "等于", "expected": "0", "memo": "业务码正常"},
            {"target": "$.data", "op": "不为空", "expected": "", "memo": "有返回数据"},
            {"target": "耗时ms", "op": "小于", "expected": "3000", "level": "仅警告",
             "memo": "响应慢只提醒不判失败"},
        ]),
        ("列表 / 分页接口", [
            {"target": "$.data.list", "op": "类型是", "expected": "数组", "memo": "列表字段类型"},
            {"target": "$.data.list", "op": "长度大于", "expected": "0", "memo": "至少一条"},
            {"target": "$.data.total", "op": "类型是", "expected": "数字", "memo": "总数是数字"},
            {"target": "$.data.list[*].id", "op": "无重复", "expected": "", "memo": "主键不重复"},
        ]),
        ("领券 / 库存（每跑一次就变）", [
            {"target": "$.data.remain", "op": "较上次减少", "expected": "1",
             "memo": "每领一张少一张，不用改期望值"},
        ]),
        ("字段结构校验", [
            {"target": "$.data", "op": "必含字段", "expected": "id,name,status",
             "memo": "关键字段不能少"},
            {"target": "$.data.list", "op": "每项都含字段", "expected": "id,name",
             "memo": "每条记录字段齐全"},
        ]),
        ("异常 / 反向用例", [
            {"target": "$.code", "op": "不等于", "expected": "0", "memo": "预期业务失败"},
            {"target": "$.msg", "op": "包含", "expected": "不存在", "memo": "错误文案符合预期"},
            {"target": "$.data", "op": "为空", "expected": "", "memo": "失败时不返回数据"},
        ]),
    ]

    def add_common(self):
        menu = tk.Menu(self, tearoff=0)
        for name, items in self.ASSERT_LIB:
            menu.add_command(label="%s（%d 条）" % (name, len(items)),
                             command=lambda it=items: self._add_group(it))
        try:
            menu.tk_popup(self.winfo_pointerx(), self.winfo_pointery())
        finally:
            menu.grab_release()

    def _add_group(self, items):
        for a in items:
            self._insert(dict(a, on=True, level=a.get("level", "必须通过")))

    def pick_field(self):
        if not self.pick_provider:
            return
        sel = self.tv.selection()

        def on_pick(path):
            iid = sel[0] if sel else self._insert(
                {"on": True, "target": "", "op": "等于", "expected": "",
                 "level": "必须通过", "memo": ""})
            self.tv.set(iid, "target", path)
            self.tv.selection_set(iid)
        JsonPickDialog(self.winfo_toplevel(), self.pick_provider, on_pick)

    def test_run(self):
        """拿一份响应把当前断言跑一遍，不写入执行历史、不影响「较上次」基线。"""
        if not self.case_provider:
            return
        rows_cfg = self.dump()
        if not rows_cfg:
            messagebox.showinfo("提示", "还没有断言可以试跑", parent=self.winfo_toplevel())
            return
        case = self.case_provider()
        top = self.winfo_toplevel()

        def worker():
            try:
                resp, source, rows = core.probe_asserts(case, rows_cfg, True)
                err = None
            except Exception as e:
                resp, source, rows, err = "", "", [], str(e)
            top.after(0, done, resp, source, rows, err)

        def done(resp, source, rows, err):
            if err:
                messagebox.showerror("试跑失败", err, parent=top)
                return
            fail, warn, passed = core.assert_stats(rows)
            head = ["响应来源：%s" % source,
                    "结论：通过 %d 条 / 失败 %d 条 / 警告 %d 条" % (passed, fail, warn), ""]
            for x in rows:
                mark = "√" if x["ok"] else ("⚠" if x.get("level") == "仅警告" else "×")
                head.append("%s %s [%s] %s" % (mark, x["target"], x["op"], x["expected"]))
                head.append("      实际：%s" % x["actual"])
                head.append("      说明：%s" % x["msg"])
                if x.get("memo"):
                    head.append("      备注：%s" % x["memo"])
                head.append("")
            text = "\n".join(head)
            full = text + "\n---------- 本次使用的响应 ----------\n" + (resp or "")
            TextDialog(top, "断言试跑结果", [("结果", text, full)])

        threading.Thread(target=worker, daemon=True).start()

    def del_row(self):
        for iid in self.tv.selection():
            self.tv.delete(iid)

    def move(self, d):
        sel = self.tv.selection()
        if not sel:
            return
        idx = self.tv.index(sel[0]) + d
        if 0 <= idx < len(self.tv.get_children()):
            self.tv.move(sel[0], "", idx)

    def clear(self):
        self.tv.delete(*self.tv.get_children())


class CaseDialog(tk.Toplevel):
    """新增 / 编辑接口用例。"""

    def __init__(self, master, case, on_save, is_new=None, on_copy=None):
        super().__init__(master)
        if is_new is None:
            is_new = case is None
        self.is_new = is_new
        self.on_save = on_save
        self.on_copy = on_copy
        self.title("新增接口用例" if is_new else "编辑接口用例")
        self.transient(master)
        center(self, 960, fit_height(self, 820))
        self.minsize(780, 480)

        c = dict(case) if case else {"name": "", "caseName": "", "url": "", "method": "POST",
                                     "headers": "", "body": "", "enabled": True,
                                     "extract": "", "results": []}
        if not c.get("id"):
            c["id"] = core.next_case_id()
        c.setdefault("params", [])
        c.setdefault("memos", {})
        c.setdefault("asserts", [])
        # URL 里带着 ?a=1&b=2 时，自动拆成参数表，一行一个
        moved = 0
        if "?" in (c.get("url") or ""):
            base, extra = core.split_query(c["url"])
            if extra:
                c["url"] = base
                c["params"] = list(c["params"]) + extra
                moved = len(extra)
        self.orig = c
        self._moved = moved

        form = ttk.Frame(self)
        form.pack(fill="x", padx=12, pady=(10, 2))
        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=2)
        ttk.Label(form, text="用例ID").grid(row=0, column=0, sticky="w", pady=3)
        self.e_id = ttk.Entry(form)
        self.e_id.grid(row=0, column=1, sticky="ew", padx=(6, 16))
        ttk.Label(form, text="接口名称").grid(row=0, column=2, sticky="w")
        self.e_name = ttk.Entry(form)
        self.e_name.grid(row=0, column=3, sticky="ew", padx=(6, 0))

        ttk.Label(form, text="接口用例").grid(row=1, column=0, sticky="w", pady=3)
        self.e_case = ttk.Entry(form)
        self.e_case.grid(row=1, column=1, sticky="ew", padx=(6, 16))
        self.v_en = tk.BooleanVar(value=c.get("enabled", True))
        ttk.Checkbutton(form, text="参与「全部执行」", variable=self.v_en
                        ).grid(row=1, column=3, sticky="w", padx=(6, 0))

        ttk.Label(form, text="请求方法").grid(row=2, column=0, sticky="w", pady=3)
        mrow = ttk.Frame(form)
        mrow.grid(row=2, column=1, columnspan=3, sticky="ew", padx=(6, 0))
        self.cb_m = ttk.Combobox(mrow, values=["GET", "POST", "PUT", "DELETE", "PATCH"],
                                 width=9, state="readonly")
        self.cb_m.pack(side="left")
        ttk.Label(mrow, text="  接口URL").pack(side="left")
        self.e_url = ttk.Entry(mrow)
        self.e_url.pack(side="left", fill="x", expand=True, padx=(6, 0))
        ttk.Label(form, foreground="#6b7280",
                  text="相对路径会自动拼接公共配置的 BaseURL；URL / 参数 / Body / 请求头里都可以用 ${变量名} 引用上一个接口提取的值"
                  ).grid(row=3, column=0, columnspan=4, sticky="w")
        self.lb_preview = ttk.Label(form, foreground=BRAND, wraplength=900, justify="left")
        self.lb_preview.grid(row=4, column=0, columnspan=4, sticky="w", pady=(2, 0))

        # 底部按钮条先 pack 到底部，保证任何屏幕高度下都不会被内容挤掉
        bar = ttk.Frame(self)
        bar.pack(side="bottom", fill="x", padx=12, pady=(2, 10))

        # 三个可拖动调整大小的区域
        paned = ttk.PanedWindow(self, orient="vertical")
        paned.pack(fill="both", expand=True, padx=12, pady=6)

        f1 = ttk.Frame(paned)
        h1 = ttk.Frame(f1); h1.pack(fill="x")
        ttk.Label(h1, text="接口级请求头 / 信息头（优先于公共配置，同名覆盖；每行一个 名称: 值，也支持 JSON）",
                  foreground="#374151").pack(side="left")
        self.lb_hcount = ttk.Label(h1, foreground=BRAND)
        self.lb_hcount.pack(side="right")
        fr, self.t_head = scroll_text(f1, 6)
        fr.pack(fill="both", expand=True, pady=(2, 4))

        f2 = ttk.Frame(paned)
        self.nb_body = ttk.Notebook(f2)
        self.nb_body.pack(fill="both", expand=True, pady=(2, 4))

        tab_p = ttk.Frame(self.nb_body)
        self.nb_body.add(tab_p, text="  请求参数（URL）  ")
        prow = ttk.Frame(tab_p); prow.pack(fill="x", padx=4, pady=(4, 0))
        self.lb_plock = ttk.Label(prow, foreground=ERRC)
        self.lb_plock.pack(side="left")
        self.btn_use_param = ttk.Button(prow, text="清空 Body，改用 URL 参数",
                                        command=self._use_params)
        self.tb_params = ParamTable(tab_p, on_change=self._on_params_change)
        self.tb_params.pack(fill="both", expand=True, padx=4, pady=4)

        tab_b = ttk.Frame(self.nb_body)
        self.nb_body.add(tab_b, text="  请求 Body  ")
        h2 = ttk.Frame(tab_b); h2.pack(fill="x", padx=4, pady=(4, 0))
        self.lb_block = ttk.Label(h2, foreground="#6b7280",
                                  text="JSON 格式；GET 请求的查询参数请放到「请求参数（URL）」页")
        self.lb_block.pack(side="left")
        self.btn_fmt = ttk.Button(h2, text="格式化 JSON", command=self._fmt_body)
        self.btn_fmt.pack(side="right")
        self.btn_use_body = ttk.Button(h2, text="清空 URL 参数，改用 Body",
                                       command=self._use_body)
        fr, self.t_body = scroll_text(tab_b, 8)
        fr.pack(fill="both", expand=True, padx=4, pady=4)

        tab_f = ttk.Frame(self.nb_body)
        self.nb_body.add(tab_f, text="  Body 字段备注  ")
        self.tb_fields = FieldTable(tab_f,
                                    lambda: self.t_body.get("1.0", "end"),
                                    self._set_body_text)
        self.tb_fields.pack(fill="both", expand=True, padx=4, pady=4)
        self.nb_body.bind("<<NotebookTabChanged>>", self._on_body_tab)

        f3 = ttk.Frame(paned)
        self.nb_bottom = ttk.Notebook(f3)
        self.nb_bottom.pack(fill="both", expand=True, pady=(2, 4))

        tab_e = ttk.Frame(self.nb_bottom)
        self.nb_bottom.add(tab_e, text="  提取变量  ")
        h3 = ttk.Frame(tab_e); h3.pack(fill="x", padx=4, pady=(4, 0))
        ttk.Label(h3, text="每行一个 变量名=$.路径，供后续接口用 ${变量名} 引用",
                  foreground="#6b7280").pack(side="left")
        ttk.Button(h3, text="测试提取 / 点响应值生成路径", style="Brand.TButton",
                   command=self._test_extract).pack(side="right")
        fr, self.t_ext = scroll_text(tab_e, 4)
        fr.pack(fill="both", expand=True, padx=4, pady=4)

        tab_a = ttk.Frame(self.nb_bottom)
        self.nb_bottom.add(tab_a, text="  断言（预期结果）  ")
        self.tb_asserts = AssertTable(tab_a, case_provider=self._probe_case)
        self.tb_asserts.pack(fill="both", expand=True, padx=4, pady=4)

        paned.add(f1, weight=1)
        paned.add(f2, weight=2)
        paned.add(f3, weight=1)

        ttk.Label(bar, foreground="#6b7280",
                  text="提示：拖动三个区域之间的分隔线可以调整大小").pack(side="left")
        ttk.Button(bar, text="取消", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(bar, text="保存", style="Brand.TButton", command=self._save).pack(side="right")
        if not is_new and on_copy:
            ttk.Button(bar, text="保存并复制为新用例", command=self._save_copy).pack(side="right", padx=8)

        self.e_id.insert(0, c["id"])
        self.e_name.insert(0, c.get("name", ""))
        self.e_case.insert(0, c.get("caseName", ""))
        self.cb_m.set(c.get("method", "POST"))
        self.e_url.insert(0, c.get("url", ""))
        self.t_head.insert("1.0", c.get("headers", "") or "")
        self.t_body.insert("1.0", c.get("body", "") or "")
        self.t_ext.insert("1.0", c.get("extract", "") or "")
        self.t_head.bind("<KeyRelease>", lambda e: self._count_headers())
        self._count_headers()
        self.tb_params.load(c.get("params"))
        self.tb_fields.load_memos(c.get("memos"))
        self.tb_asserts.load(c.get("asserts"))
        self.t_body.bind("<KeyRelease>", lambda e: self.after(150, self._sync_exclusive))
        self._sync_exclusive()
        self.e_url.bind("<KeyRelease>", lambda e: self._refresh_preview())
        self.cb_m.bind("<<ComboboxSelected>>", lambda e: self._refresh_preview())
        self._refresh_preview()
        # GET 默认停在参数页，其它方法停在 Body 页
        if (c.get("body") or "").strip():
            self.nb_body.select(1)
        elif self.tb_params.count() or (c.get("method") or "").upper() in ("GET", "DELETE", "HEAD"):
            self.nb_body.select(0)
        else:
            self.nb_body.select(1)
        self.grab_set()
        self.e_name.focus_set()

    def _set_body_text(self, text):
        st = str(self.t_body.cget("state"))
        self.t_body.configure(state="normal")
        self.t_body.delete("1.0", "end")
        self.t_body.insert("1.0", text)
        self.t_body.configure(state=st)

    def _on_params_change(self):
        self._refresh_preview()
        self._sync_exclusive()

    def _use_params(self, confirm=True):
        """清空 Body，改用 URL 参数。"""
        if confirm and self.t_body.get("1.0", "end").strip() and not messagebox.askyesno(
                "确认", "将清空请求 Body 的内容，改用 URL 参数，确定？", parent=self):
            return
        self._set_body_text("")
        self._sync_exclusive()

    def _use_body(self, confirm=True):
        """清空参数表，改用 Body。"""
        if confirm and self.tb_params.count() and not messagebox.askyesno(
                "确认", "将清空 URL 参数表，改用请求 Body，确定？", parent=self):
            return
        self.tb_params.clear()
        self._sync_exclusive()

    def _body_enabled(self, flag):
        self.t_body.configure(state="normal" if flag else "disabled",
                              background="#ffffff" if flag else "#eef1f5")
        self.btn_fmt.configure(state="normal" if flag else "disabled")

    def _sync_exclusive(self):
        """请求参数与请求 Body 互斥；Body 有内容时才能用「Body 字段备注」。"""
        has_p = len(self.tb_params.dump()) > 0
        has_b = bool(self.t_body.get("1.0", "end").strip())

        # 字段备注页：只有 Body 有内容时可用
        try:
            self.nb_body.tab(2, state="normal" if has_b else "disabled")
        except Exception:
            pass
        if not has_b and self.nb_body.index("current") == 2:
            self.nb_body.select(1)

        if has_p and has_b:      # 历史数据可能两者都有，提示用户二选一
            self._body_enabled(True)
            self.tb_params.set_locked(False)
            self.lb_plock.config(text="⚠ URL 参数和 Body 同时有内容，请清空其中一个")
            self.lb_block.config(text="⚠ URL 参数和 Body 同时有内容，请清空其中一个", foreground=ERRC)
            self.btn_use_param.pack(side="left", padx=8)
            self.btn_use_body.pack(side="right", padx=8)
        elif has_p:              # 用参数 → 锁 Body
            self._body_enabled(False)
            self.tb_params.set_locked(False)
            self.lb_plock.config(text="")
            self.btn_use_param.pack_forget()
            self.lb_block.config(text="已锁定：这个接口用的是 URL 参数，Body 不可填写",
                                 foreground=ERRC)
            self.btn_use_body.pack(side="right", padx=8)
        elif has_b:              # 用 Body → 锁参数表
            self._body_enabled(True)
            self.tb_params.set_locked(True)
            self.lb_plock.config(text="已锁定：这个接口用的是请求 Body")
            self.btn_use_param.pack(side="left", padx=8)
            self.lb_block.config(text="JSON 格式；备注请到「Body 字段备注」页逐字段添加",
                                 foreground="#6b7280")
            self.btn_use_body.pack_forget()
        else:                    # 都是空的 → 两边都能填
            self._body_enabled(True)
            self.tb_params.set_locked(False)
            self.lb_plock.config(text="")
            self.btn_use_param.pack_forget()
            self.lb_block.config(text="JSON 格式；GET 请求的查询参数请放到「请求参数（URL）」页",
                                 foreground="#6b7280")
            self.btn_use_body.pack_forget()

    def _on_body_tab(self, _e=None):
        if self.nb_body.index("current") == 2:
            self.tb_fields.reload()

    def _refresh_preview(self):
        url = self.e_url.get().strip()
        params = self.tb_params.dump()
        full = core.join_query(url, params)
        tip = ""
        if self._moved:
            tip = "（已把 URL 里的 %d 个查询参数拆到下方「请求参数」页，一行一个）" % self._moved
        self.lb_preview.config(text="实际请求：%s %s%s" % (self.cb_m.get(), full, tip))

    def _count_headers(self):
        n = len(core.parse_headers(self.t_head.get("1.0", "end")))
        self.lb_hcount.config(text="共 %d 个请求头%s" % (n, "（可滚动查看）" if n > 6 else ""))

    def _fmt_body(self):
        raw = self.t_body.get("1.0", "end").strip()
        if not raw:
            return
        try:
            pretty = json.dumps(json.loads(raw), ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showwarning("不是合法 JSON", "Body 解析失败：%s" % e, parent=self)
            return
        self.t_body.delete("1.0", "end")
        self.t_body.insert("1.0", pretty)

    def _collect(self):
        url = self.e_url.get().strip()
        params = self.tb_params.dump()
        body_text = self.t_body.get("1.0", "end").strip()
        if "?" in url:                      # 用户又把参数打回了 URL，照样拆出来
            url, extra = core.split_query(url)
            params = params + extra
        return {"id": self.e_id.get().strip() or core.next_case_id(),
                "name": self.e_name.get().strip(), "caseName": self.e_case.get().strip(),
                "url": url, "method": self.cb_m.get(),
                "enabled": self.v_en.get(),
                "headers": self.t_head.get("1.0", "end").strip(),
                "body": body_text,
                "params": [] if body_text else params,
                "memos": self.tb_fields.dump_memos() if body_text else {},
                "extract": self.t_ext.get("1.0", "end").strip(),
                "asserts": self.tb_asserts.dump(),
                "results": self.orig.get("results", [])}

    def _validate(self, nc):
        if not nc["url"]:
            messagebox.showwarning("提示", "接口URL不能为空", parent=self)
            return False
        others = [c["id"] for c in core.STATE["cases"]
                  if self.is_new or c["id"] != self.orig.get("id")]
        if nc["id"] in others:
            messagebox.showwarning("提示", "用例ID「%s」已被其他用例使用，请换一个" % nc["id"], parent=self)
            return False
        return True

    def _save(self):
        nc = self._collect()
        if not self._validate(nc):
            return
        self.on_save(nc)
        self.destroy()

    def _save_copy(self):
        nc = self._collect()
        if not self._validate(nc):
            return
        self.on_save(nc)
        self.destroy()
        self.on_copy(nc["id"])

    def _probe_case(self):
        c = self._collect()
        c["extract"] = ""
        return c

    def _test_extract(self):
        def provider():
            c = self._collect()
            c["extract"] = ""
            return c

        def apply(text):
            self.t_ext.delete("1.0", "end")
            self.t_ext.insert("1.0", text)

        ExtractTestDialog(self, provider, self.t_ext.get("1.0", "end").strip(), apply)


class ExtractTestDialog(tk.Toplevel):
    """调试提取变量：逐条校验 + 点击响应里的值自动生成路径。"""

    def __init__(self, master, case_provider, extract_text, on_apply=None):
        super().__init__(master)
        self.title("测试提取变量  ·  点击响应里的值即可生成路径")
        self.transient(master)
        center(self, 1120, fit_height(self, 860))
        self.minsize(900, 560)
        self.case_provider = case_provider
        self.on_apply = on_apply
        self.busy = False
        self.parsed = None
        self.resp_text = ""
        self.node_keys = {}

        top = ttk.Frame(self)
        top.pack(fill="x", padx=12, pady=(10, 2))
        ttk.Label(top, text="响应来源", font=("Microsoft YaHei", 9, "bold")).pack(side="left")
        self.src = tk.StringVar(value="cached")
        ttk.Radiobutton(top, text="最近一次执行结果", value="cached",
                        variable=self.src).pack(side="left", padx=6)
        ttk.Radiobutton(top, text="立即请求一次（不计入结果历史）", value="live",
                        variable=self.src).pack(side="left")
        self.btn_fetch = ttk.Button(top, text="获取响应", command=self.fetch)
        self.btn_fetch.pack(side="left", padx=10)
        self.lb_src = ttk.Label(top, text="", foreground="#6b7280")
        self.lb_src.pack(side="left")

        paned = ttk.PanedWindow(self, orient="vertical")
        paned.pack(fill="both", expand=True, padx=12, pady=4)

        # ---- 上：提取配置 + 校验结果 ----
        up = ttk.Frame(paned)
        hrow = ttk.Frame(up); hrow.pack(fill="x")
        ttk.Label(hrow, text="提取配置（每行 变量名=$.路径；可手写，也可在下方响应里点值自动添加）",
                  foreground="#374151").pack(side="left")
        ttk.Button(hrow, text="▶ 校验提取", style="Brand.TButton",
                   command=self.evaluate).pack(side="right")
        fr, self.t_ext = scroll_text(up, 4)
        fr.pack(fill="both", expand=True, pady=(2, 4))
        self.t_ext.insert("1.0", extract_text or "")

        box = ttk.Frame(up); box.pack(fill="both", expand=True)
        cols = ("line", "var", "path", "status", "value")
        self.tv = ttk.Treeview(box, columns=cols, show="headings", height=4)
        for c, t, w, a in (("line", "行", 40, "center"), ("var", "变量名", 120, "w"),
                           ("path", "JSON路径", 250, "w"), ("status", "结果", 230, "w"),
                           ("value", "取到的值", 380, "w")):
            self.tv.heading(c, text=t)
            self.tv.column(c, width=w, anchor=a)
        self.tv.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tv.yview)
        sb.pack(side="right", fill="y")
        self.tv.configure(yscrollcommand=sb.set)
        self.tv.tag_configure("ok", foreground=OKC)
        self.tv.tag_configure("err", foreground=ERRC)
        self.tv.bind("<Double-1>", self._show_value)

        # ---- 下：响应树 + 路径生成 ----
        down = ttk.Frame(paned)
        pick = ttk.Labelframe(down, text=" 点击下方响应中的任意字段 → 自动生成路径 ", padding=6)
        pick.pack(fill="x", pady=(4, 4))
        ttk.Label(pick, text="路径").grid(row=0, column=0, sticky="w")
        self.e_path = ttk.Entry(pick, font=("Consolas", 10))
        self.e_path.grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Label(pick, text="变量名").grid(row=0, column=2, sticky="w")
        self.e_var = ttk.Entry(pick, width=18)
        self.e_var.grid(row=0, column=3, sticky="w", padx=6)
        ttk.Button(pick, text="添加为变量", style="Brand.TButton",
                   command=lambda: self.add_var(False)).grid(row=0, column=4, padx=2)
        ttk.Button(pick, text="取整列 [*] 添加",
                   command=lambda: self.add_var(True)).grid(row=0, column=5, padx=2)
        ttk.Button(pick, text="复制路径", command=self.copy_path).grid(row=0, column=6, padx=2)
        self.lb_pick = ttk.Label(pick, foreground="#6b7280",
                                 text="单击选中字段，双击直接添加为变量；在数组里的字段可用「取整列」一次取所有项")
        self.lb_pick.grid(row=1, column=0, columnspan=7, sticky="w", pady=(4, 0))
        pick.columnconfigure(1, weight=1)

        self.nb = ttk.Notebook(down)
        self.nb.pack(fill="both", expand=True)
        f_tree = ttk.Frame(self.nb)
        self.nb.add(f_tree, text="  响应（树形，可点击）  ")
        self.jt = ttk.Treeview(f_tree, columns=("value", "type"), show="tree headings")
        self.jt.heading("#0", text="字段")
        self.jt.heading("value", text="值")
        self.jt.heading("type", text="类型")
        self.jt.column("#0", width=300, stretch=False)
        self.jt.column("value", width=560)
        self.jt.column("type", width=80, anchor="center", stretch=False)
        ys = ttk.Scrollbar(f_tree, orient="vertical", command=self.jt.yview)
        xs = ttk.Scrollbar(f_tree, orient="horizontal", command=self.jt.xview)
        self.jt.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.jt.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        f_tree.rowconfigure(0, weight=1)
        f_tree.columnconfigure(0, weight=1)
        self.jt.bind("<<TreeviewSelect>>", self._on_pick)
        self.jt.bind("<Double-1>", lambda e: self.add_var(False))

        f_raw = ttk.Frame(self.nb)
        self.nb.add(f_raw, text="  响应（原始文本）  ")
        self.t_raw = ScrolledText(f_raw, font=("Consolas", 10), wrap="none")
        self.t_raw.pack(fill="both", expand=True)

        paned.add(up, weight=1)
        paned.add(down, weight=3)

        bottom = ttk.Frame(self)
        bottom.pack(side="bottom", fill="x", padx=12, pady=(2, 10))
        ttk.Button(bottom, text="关闭", command=self.destroy).pack(side="right", padx=4)
        if on_apply:
            ttk.Button(bottom, text="应用并关闭", style="Brand.TButton",
                       command=self.apply).pack(side="right")
        self.grab_set()
        self.after(200, self.fetch)

    # ------------------------------------------------ 获取响应
    def fetch(self):
        if self.busy:
            return
        case = self.case_provider()
        if not case:
            messagebox.showwarning("提示", "没有可测试的接口", parent=self)
            return
        self.busy = True
        self.btn_fetch.config(text="获取中…", state="disabled")
        use_cached = (self.src.get() == "cached")

        def worker():
            try:
                resp, source, _rows = core.probe_extract(case, "", use_cached)
                err = None
            except Exception as e:
                resp, source, err = "", "", str(e)
            self.after(0, done, resp, source, err)

        def done(resp, source, err):
            self.busy = False
            self.btn_fetch.config(text="获取响应", state="normal")
            if err:
                messagebox.showerror("获取失败", err, parent=self)
                return
            self.resp_text = resp or ""
            try:
                self.parsed = json.loads(self.resp_text)
            except Exception:
                self.parsed = None
            self.lb_src.config(text="当前响应：%s" % source)
            self.t_raw.delete("1.0", "end")
            self.t_raw.insert("1.0", self.resp_text or "（没有拿到响应内容）")
            self._fill_tree()
            self.evaluate()

        threading.Thread(target=worker, daemon=True).start()

    # ------------------------------------------------ 树
    def _fill_tree(self):
        jt = self.jt
        jt.delete(*jt.get_children())
        self.node_keys = {}
        if self.parsed is None:
            jt.insert("", "end", text="响应不是 JSON，无法按字段点选（可到「原始文本」页查看）")
            self.nb.select(1)
            return
        self.nb.select(0)
        counter = [0]
        type_name = {str: "字符串", bool: "布尔", int: "数字", float: "数字", type(None): "null"}

        def add(parent, label, value, keys, depth):
            counter[0] += 1
            if counter[0] > 8000:
                return
            if isinstance(value, dict):
                iid = jt.insert(parent, "end", text=label,
                                values=("{ %d 个字段 }" % len(value), "对象"), open=depth < 2)
                self.node_keys[iid] = keys
                for k, v in value.items():
                    add(iid, str(k), v, keys + [k], depth + 1)
            elif isinstance(value, list):
                iid = jt.insert(parent, "end", text=label,
                                values=("[ %d 项 ]" % len(value), "数组"), open=depth < 2)
                self.node_keys[iid] = keys
                for i, v in enumerate(value[:300]):
                    add(iid, "[%d]" % i, v, keys + [i], depth + 1)
                if len(value) > 300:
                    jt.insert(iid, "end", text="……还有 %d 项未展示" % (len(value) - 300))
            else:
                disp = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False)
                disp = disp.replace("\n", " ")
                if len(disp) > 300:
                    disp = disp[:300] + "…"
                iid = jt.insert(parent, "end", text=label,
                                values=(disp, type_name.get(type(value), "")))
                self.node_keys[iid] = keys

        add("", "$（响应根节点）", self.parsed, [], 0)
        # 默认展开首个数组的第一项，方便直接点字段
        for iid in jt.get_children(""):
            self._open_first_items(iid, 4)

    def _open_first_items(self, iid, depth):
        if depth <= 0:
            return
        self.jt.item(iid, open=True)
        kids = self.jt.get_children(iid)
        if kids:
            self._open_first_items(kids[0], depth - 1)

    def _on_pick(self, _e=None):
        sel = self.jt.selection()
        if not sel or sel[0] not in self.node_keys:
            return
        keys = self.node_keys[sel[0]]
        self.e_path.delete(0, "end")
        self.e_path.insert(0, core.build_path(keys))
        self.e_var.delete(0, "end")
        self.e_var.insert(0, self._suggest_name(keys, False))
        has_arr = any(isinstance(k, int) for k in keys)
        self.lb_pick.config(text=("该字段位于数组中，点「取整列 [*] 添加」可一次取出所有项的这个字段"
                                  if has_arr else "单击选中字段，双击直接添加为变量"))

    def _existing_vars(self):
        return {d["var"] for d in core.parse_extract_detail(self.t_ext.get("1.0", "end")) if d["var"]}

    def _suggest_name(self, keys, whole):
        name = ""
        for k in reversed(keys):
            if isinstance(k, str) and k != "*":
                name = k
                break
        name = re.sub(r"[^0-9A-Za-z_\u4e00-\u9fa5]", "_", name) or "var"
        if name[0].isdigit():
            name = "v_" + name
        if whole:
            name += "List"
        elif keys and isinstance(keys[-1], int):
            name += str(keys[-1])
        exist = self._existing_vars()
        base, n = name, 2
        while name in exist:
            name = "%s%d" % (base, n)
            n += 1
        return name

    def add_var(self, whole):
        sel = self.jt.selection()
        if not sel or sel[0] not in self.node_keys:
            messagebox.showinfo("提示", "请先在下方响应里点选一个字段", parent=self)
            return
        keys = list(self.node_keys[sel[0]])
        if whole:
            idx = max((i for i, k in enumerate(keys) if isinstance(k, int)), default=None)
            if idx is None:
                messagebox.showinfo("提示", "这个字段不在数组里，不需要取整列，直接「添加为变量」即可", parent=self)
                return
            keys[idx] = "*"
            path = core.build_path(keys)
            var = self._suggest_name(keys, True)
        else:
            path = self.e_path.get().strip() or core.build_path(keys)
            var = self.e_var.get().strip() or self._suggest_name(keys, False)
            if var in self._existing_vars():
                var = self._suggest_name(keys, False)
        cur = self.t_ext.get("1.0", "end").rstrip("\n")
        line = "%s=%s" % (var, path)
        self.t_ext.delete("1.0", "end")
        self.t_ext.insert("1.0", (cur + "\n" + line) if cur.strip() else line)
        self.evaluate()
        self.lb_pick.config(text="已添加：%s" % line)

    def copy_path(self):
        p = self.e_path.get().strip()
        if p:
            self.clipboard_clear()
            self.clipboard_append(p)
            self.lb_pick.config(text="已复制到剪贴板：%s" % p)

    # ------------------------------------------------ 校验
    def evaluate(self):
        if not self.resp_text and self.parsed is None:
            return
        rows = core.evaluate_extract(self.parsed, self.t_ext.get("1.0", "end"))
        self.tv.delete(*self.tv.get_children())
        for r in rows:
            self.tv.insert("", "end", tags=("ok" if r.get("ok") else "err",),
                           values=(r.get("line") or "", r.get("var", ""), r.get("path", ""),
                                   r.get("status", ""), str(r.get("value", ""))[:600]))

    def _show_value(self, _e):
        sel = self.tv.selection()
        if not sel:
            return
        vals = self.tv.item(sel[0], "values")
        rows = core.evaluate_extract(self.parsed, self.t_ext.get("1.0", "end"))
        full = ""
        for r in rows:
            if str(r.get("line")) == str(vals[0]):
                full = str(r.get("value", ""))
        TextDialog(self, "变量 %s 的取值" % vals[1], [("取值", full[:PREVIEW_LIMIT], full)])

    def apply(self):
        if self.on_apply:
            self.on_apply(self.t_ext.get("1.0", "end").strip())
        self.destroy()


class CurlDialog(tk.Toplevel):
    """从 cURL 命令导入接口。"""

    def __init__(self, master, on_import):
        super().__init__(master)
        self.title("从 cURL 导入接口")
        self.transient(master)
        center(self, 980, fit_height(self, 760))
        self.minsize(760, 480)
        self.on_import = on_import
        self.cases = []

        ttk.Label(self, foreground="#374151",
                  text="把 cURL 命令粘贴到下面。支持浏览器 F12「复制为 cURL (bash)」和「复制为 cURL (cmd)」、"
                       "Postman / Apifox 导出的 cURL；一次可以粘贴多条。"
                  ).pack(anchor="w", padx=12, pady=(12, 4))
        fr, self.t = scroll_text(self, 12, wrap="word")
        fr.pack(fill="both", expand=True, padx=12)

        opt = ttk.Frame(self); opt.pack(fill="x", padx=12, pady=6)
        self.v_rel = tk.BooleanVar(value=True)
        self.v_slim = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="URL 与公共 BaseURL 相同前缀时，自动转成相对路径",
                        variable=self.v_rel).pack(side="left")
        ttk.Checkbutton(opt, text="去掉浏览器自带的无关请求头（sec-ch-*、sec-fetch-*、priority 等）",
                        variable=self.v_slim).pack(side="left", padx=16)

        bar = ttk.Frame(self); bar.pack(fill="x", padx=12)
        ttk.Button(bar, text="▶ 解析", style="Brand.TButton", command=self.parse).pack(side="left")
        ttk.Button(bar, text="从剪贴板粘贴", command=self.paste).pack(side="left", padx=6)
        ttk.Button(bar, text="清空", command=lambda: self.t.delete("1.0", "end")).pack(side="left")

        ttk.Label(self, text="解析预览", font=("Microsoft YaHei", 9, "bold")
                  ).pack(anchor="w", padx=12, pady=(10, 2))
        box = ttk.Frame(self); box.pack(fill="both", expand=True, padx=12)
        cols = ("no", "method", "url", "hcount", "body")
        self.tv = ttk.Treeview(box, columns=cols, show="headings", height=5)
        for c, t, w in (("no", "#", 36), ("method", "方法", 64), ("url", "URL", 420),
                        ("hcount", "请求头", 70), ("body", "Body 摘要", 300)):
            self.tv.heading(c, text=t)
            self.tv.column(c, width=w, anchor="w")
        self.tv.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tv.yview)
        sb.pack(side="right", fill="y")
        self.tv.configure(yscrollcommand=sb.set)
        self.lb_warn = ttk.Label(self, foreground=ERRC, wraplength=900, justify="left")
        self.lb_warn.pack(anchor="w", padx=12, pady=(4, 0))

        bottom = ttk.Frame(self)
        bottom.pack(side="bottom", fill="x", padx=12, pady=10)
        ttk.Label(bottom, foreground="#6b7280",
                  text="导入后立刻生成新的接口用例（已加入列表），并打开编辑窗口方便改名称和用例说明"
                  ).pack(side="left")
        ttk.Button(bottom, text="取消", command=self.destroy).pack(side="right", padx=4)
        self.btn_ok = ttk.Button(bottom, text="导入为新接口", style="Brand.TButton",
                                 command=self.do_import, state="disabled")
        self.btn_ok.pack(side="right")
        self.grab_set()
        self.t.focus_set()

    def paste(self):
        try:
            txt = self.clipboard_get()
        except Exception:
            txt = ""
        if txt:
            self.t.delete("1.0", "end")
            self.t.insert("1.0", txt)
            self.parse()

    def parse(self):
        text = self.t.get("1.0", "end")
        base = core.STATE["config"].get("baseUrl", "")
        try:
            self.cases, warns = core.parse_curl(text, base, self.v_rel.get(), self.v_slim.get())
        except Exception as e:
            self.cases, warns = [], ["解析出错：%s" % e]
        self.tv.delete(*self.tv.get_children())
        for i, c in enumerate(self.cases, 1):
            n = len(core.parse_headers(c["headers"]))
            body = c["body"].replace("\n", " ")
            self.tv.insert("", "end", values=(i, c["method"], c["url"], n,
                                              body[:120] + ("…" if len(body) > 120 else "")))
        dropped = sum(c.get("_dropped", 0) for c in self.cases)
        tips = list(warns)
        if dropped:
            tips.append("已去掉 %d 个浏览器自带的无关请求头" % dropped)
        if self.cases:
            tips.insert(0, "解析成功：%d 个接口" % len(self.cases))
        self.lb_warn.config(text="\n".join(tips),
                            foreground=(OKC if self.cases and not warns else ERRC))
        self.btn_ok.config(state="normal" if self.cases else "disabled")

    def do_import(self):
        if not self.cases:
            return
        for c in self.cases:
            c.pop("_dropped", None)
        cases = self.cases
        self.destroy()
        self.on_import(cases)


class StepDialog(tk.Toplevel):
    """场景中的单个步骤。"""

    def __init__(self, master, step, on_save):
        super().__init__(master)
        self.title("场景步骤")
        self.transient(master)
        center(self, 720, 380)
        self.on_save = on_save
        step = step or {"caseId": "", "extract": "", "memo": ""}

        ttk.Label(self, text="选择接口用例").pack(anchor="w", padx=12, pady=(12, 2))
        self.items = ["%s | %s | %s %s" % (c["id"], c.get("name", ""), c.get("method", ""), c.get("url", ""))
                      for c in core.STATE["cases"]]
        self.ids = [c["id"] for c in core.STATE["cases"]]
        self.cb = ttk.Combobox(self, values=self.items, state="readonly")
        self.cb.pack(fill="x", padx=12)
        if step["caseId"] in self.ids:
            self.cb.current(self.ids.index(step["caseId"]))
        elif self.items:
            self.cb.current(0)

        ttk.Label(self, text="本步骤提取变量（变量名=$.路径，每行一个；留空则沿用该用例自带的配置）"
                  ).pack(anchor="w", padx=12, pady=(10, 2))
        self.t_ext = ScrolledText(self, height=5, font=("Consolas", 10))
        self.t_ext.pack(fill="both", expand=True, padx=12)
        self.t_ext.insert("1.0", step.get("extract", "") or "")

        srow = ttk.Frame(self); srow.pack(fill="x", padx=12, pady=(4, 0))
        ttk.Button(srow, text="测试提取（看看能不能取到数据）",
                   command=self._test_extract).pack(side="left")
        ttk.Label(srow, foreground="#6b7280",
                  text="  留空时会先带出该用例自带的提取配置"
                  ).pack(side="left")

        ttk.Label(self, text="备注").pack(anchor="w", padx=12, pady=(8, 2))
        self.e_memo = ttk.Entry(self)
        self.e_memo.pack(fill="x", padx=12)
        self.e_memo.insert(0, step.get("memo", "") or "")

        bar = ttk.Frame(self); bar.pack(fill="x", padx=12, pady=10)
        ttk.Button(bar, text="取消", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(bar, text="确定", command=self._ok).pack(side="right")
        self.grab_set()

    def _test_extract(self):
        if not self.ids or self.cb.current() < 0:
            messagebox.showwarning("提示", "请先选择一个接口用例", parent=self)
            return
        cid = self.ids[self.cb.current()]

        def provider():
            c = core.find_case(cid)
            if not c:
                return None
            c = dict(c)
            c["extract"] = ""
            return c

        def apply(text):
            self.t_ext.delete("1.0", "end")
            self.t_ext.insert("1.0", text)

        cur = self.t_ext.get("1.0", "end").strip()
        if not cur:
            c = core.find_case(cid)
            cur = (c or {}).get("extract", "")
        ExtractTestDialog(self, provider, cur, apply)

    def _ok(self):
        if not self.ids:
            messagebox.showwarning("提示", "还没有接口用例，请先导入或新增接口", parent=self)
            return
        cid = self.ids[self.cb.current()] if self.cb.current() >= 0 else self.ids[0]
        self.on_save({"caseId": cid, "extract": self.t_ext.get("1.0", "end").strip(),
                      "memo": self.e_memo.get().strip()})
        self.destroy()


class SceneDialog(tk.Toplevel):
    """新建 / 编辑组合场景。"""

    def __init__(self, master, scene, on_save, preset_ids=None):
        super().__init__(master)
        self.title("编辑组合场景" if scene else "新建组合场景")
        self.transient(master)
        center(self, 860, 520)
        self.on_save = on_save
        sc = scene or {"id": core.next_scene_id(), "name": "",
                       "steps": [{"caseId": i, "extract": "", "memo": ""} for i in (preset_ids or [])],
                       "results": []}
        self.sc = sc
        self.steps = [dict(s) for s in sc.get("steps", [])]

        top = ttk.Frame(self); top.pack(fill="x", padx=12, pady=(12, 4))
        ttk.Label(top, text="场景名称", width=10).pack(side="left")
        self.e_name = ttk.Entry(top)
        self.e_name.pack(side="left", fill="x", expand=True)
        self.e_name.insert(0, sc.get("name", ""))
        ttk.Label(self, text="多接口组合（2 步及以上）必须填写场景名称，该名称会写入导出 Excel 的 Sheet2",
                  foreground="#6b7280").pack(anchor="w", padx=12)

        ttk.Label(self, text="执行步骤（自上而下顺序执行）", font=("Microsoft YaHei", 10, "bold")
                  ).pack(anchor="w", padx=12, pady=(10, 4))
        mid = ttk.Frame(self); mid.pack(fill="both", expand=True, padx=12)
        cols = ("no", "cid", "info", "ext", "memo")
        self.tv = ttk.Treeview(mid, columns=cols, show="headings", height=8)
        for c, t, w in (("no", "步骤", 50), ("cid", "用例ID", 90), ("info", "接口", 260),
                        ("ext", "提取变量", 220), ("memo", "备注", 160)):
            self.tv.heading(c, text=t)
            self.tv.column(c, width=w, anchor="w")
        self.tv.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tv.yview)
        sb.pack(side="right", fill="y")
        self.tv.configure(yscrollcommand=sb.set)
        self.tv.bind("<Double-1>", lambda e: self.edit_step())

        bb = ttk.Frame(self); bb.pack(fill="x", padx=12, pady=6)
        for txt, cmd in (("+ 添加步骤", self.add_step), ("编辑步骤", self.edit_step),
                         ("↑ 上移", lambda: self.move(-1)), ("↓ 下移", lambda: self.move(1)),
                         ("删除步骤", self.del_step)):
            ttk.Button(bb, text=txt, command=cmd).pack(side="left", padx=3)

        ttk.Label(self, text="上一步「提取变量」定义的变量，下一步可在 URL / Body / 请求头里用 ${变量名} 引用",
                  foreground="#6b7280").pack(anchor="w", padx=12)

        bar = ttk.Frame(self); bar.pack(side="bottom", fill="x", padx=12, pady=10)
        ttk.Button(bar, text="取消", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(bar, text="保存场景", style="Brand.TButton", command=self._save).pack(side="right")
        self.refresh()
        self.grab_set()

    def refresh(self):
        self.tv.delete(*self.tv.get_children())
        for i, s in enumerate(self.steps):
            c = core.find_case(s["caseId"])
            info = "%s %s" % (c.get("method", ""), c.get("url", "")) if c else "(用例不存在)"
            self.tv.insert("", "end", iid=str(i),
                           values=(i + 1, s["caseId"], info,
                                   (s.get("extract") or "").replace("\n", " ; "), s.get("memo", "")))

    def _sel(self):
        s = self.tv.selection()
        return int(s[0]) if s else None

    def add_step(self):
        StepDialog(self, None, lambda st: (self.steps.append(st), self.refresh()))

    def edit_step(self):
        i = self._sel()
        if i is None:
            messagebox.showinfo("提示", "请先选中一个步骤", parent=self); return

        def save(st):
            self.steps[i] = st
            self.refresh()
        StepDialog(self, self.steps[i], save)

    def del_step(self):
        i = self._sel()
        if i is None:
            return
        self.steps.pop(i)
        self.refresh()

    def move(self, d):
        i = self._sel()
        if i is None:
            return
        j = i + d
        if 0 <= j < len(self.steps):
            self.steps[i], self.steps[j] = self.steps[j], self.steps[i]
            self.refresh()
            self.tv.selection_set(str(j))

    def _save(self):
        name = self.e_name.get().strip()
        if not self.steps:
            messagebox.showwarning("提示", "至少需要一个步骤", parent=self); return
        if len(self.steps) > 1 and not name:
            messagebox.showwarning("提示", "多接口组合必须填写场景名称", parent=self); return
        self.on_save({"id": self.sc["id"], "name": name or self.steps[0]["caseId"],
                      "steps": self.steps, "enabled": True,
                      "results": self.sc.get("results", [])})
        self.destroy()


# ====================================================================== 主窗口
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        center(self, 1240, 820)
        self.minsize(1080, 680)
        self.configure(bg=BG)
        self.logq = queue.Queue()
        self.running = False
        self.checked_ids = set()
        self.sort_col = None
        self.sort_desc = False

        self._style()
        self._build()
        core.load_state()
        self.load_config_to_ui()
        self.refresh_cases()
        self.refresh_scenes()
        self.after(150, self._drain_log)
        self.protocol("WM_DELETE_WINDOW", self._close)

    # ------------------------------------------------ 外观
    def _style(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except Exception:
            pass
        f = ("Microsoft YaHei", 9)
        st.configure(".", font=f, background=BG)
        st.configure("TFrame", background=BG)
        st.configure("TLabel", background=BG)
        st.configure("TLabelframe", background=BG)
        st.configure("TLabelframe.Label", background=BG, foreground=BRAND,
                     font=("Microsoft YaHei", 9, "bold"))
        st.configure("TButton", padding=(10, 5))
        st.configure("Brand.TButton", foreground="#ffffff", background=BRAND, padding=(14, 6),
                     font=("Microsoft YaHei", 9, "bold"))
        st.map("Brand.TButton", background=[("active", "#3f6fb5")])
        st.configure("Danger.TButton", foreground=ERRC)
        st.configure("Treeview", rowheight=26, fieldbackground="#ffffff", background="#ffffff")
        st.configure("Treeview.Heading", font=("Microsoft YaHei", 9, "bold"))

    # ------------------------------------------------ 布局
    def _build(self):
        top = ttk.Frame(self, padding=(10, 8, 10, 4))
        top.pack(fill="x")
        ttk.Label(top, text=APP_TITLE, font=("Microsoft YaHei", 14, "bold"),
                  foreground=BRAND).pack(side="left")
        self.lb_file = ttk.Label(top, text="未导入文件", foreground="#6b7280")
        self.lb_file.pack(side="left", padx=12)
        for txt, cmd in (("重置工具", self.reset_all), ("导出结果", self.do_export),
                         ("导入Excel", self.do_import), ("生成Excel模板", self.do_template)):
            ttk.Button(top, text=txt, command=cmd).pack(side="right", padx=3)

        # ---- 公共配置 ----
        cfg = ttk.Labelframe(self, text=" 公共配置（请求头 / 信息头 / BaseURL / 编码） ",
                             padding=10)
        cfg.pack(fill="x", padx=10, pady=4)
        left = ttk.Frame(cfg); left.pack(side="left", fill="y")
        ttk.Label(left, text="BaseURL").grid(row=0, column=0, sticky="w", pady=3)
        self.e_base = ttk.Entry(left, width=46)
        self.e_base.grid(row=0, column=1, columnspan=3, sticky="w", padx=6)
        ttk.Label(left, text="内容编码").grid(row=1, column=0, sticky="w", pady=3)
        self.cb_enc = ttk.Combobox(left, values=["utf-8", "gbk", "gb2312", "utf-8-sig"],
                                   width=10, state="readonly")
        self.cb_enc.grid(row=1, column=1, sticky="w", padx=6)
        ttk.Label(left, text="超时(秒)").grid(row=1, column=2, sticky="e")
        self.e_to = ttk.Entry(left, width=8)
        self.e_to.grid(row=1, column=3, sticky="w", padx=6)
        self.v_ssl = tk.BooleanVar(value=False)
        ttk.Checkbutton(left, text="校验SSL证书（测试环境自签证书请不要勾选）",
                        variable=self.v_ssl).grid(row=2, column=0, columnspan=4, sticky="w", pady=3)
        self.v_stop = tk.BooleanVar(value=False)
        ttk.Checkbutton(left, text="组合场景里断言失败就停止后续步骤",
                        variable=self.v_stop).grid(row=3, column=0, columnspan=4, sticky="w")
        ttk.Button(left, text="保存配置", style="Brand.TButton",
                   command=self.save_config).grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))

        right = ttk.Frame(cfg); right.pack(side="left", fill="both", expand=True, padx=(24, 0))
        ttk.Label(right, text="公共请求头 / 信息头　每行一个 “名称: 值”，行首加 # 表示停用").pack(anchor="w")
        self.t_hdr = ScrolledText(right, height=5, font=("Consolas", 10))
        self.t_hdr.pack(fill="both", expand=True, pady=(3, 0))
        ttk.Label(right, text="读取顺序：先读接口下的配置，同名时覆盖这里的公共配置",
                  foreground="#6b7280").pack(anchor="w")

        # ---- 执行按钮 ----
        run = ttk.Frame(self, padding=(10, 4))
        run.pack(fill="x")
        ttk.Button(run, text="▶  全部执行", style="Brand.TButton",
                   command=self.run_all).pack(side="left")
        ttk.Button(run, text="▶  执行选中的接口 / 场景", style="Brand.TButton",
                   command=self.run_selected).pack(side="left", padx=8)
        ttk.Button(run, text="清空执行结果", style="Danger.TButton",
                   command=self.clear_all).pack(side="left", padx=8)
        self.lb_status = ttk.Label(run, text="就绪　（每个接口和场景各保留最近 5 次执行结果）",
                                   foreground="#6b7280")
        self.lb_status.pack(side="left", padx=14)

        # ---- 主体 ----
        nb = ttk.Notebook(self)
        self.nb_main = nb
        nb.pack(fill="both", expand=True, padx=10, pady=(4, 8))
        self.tab_case = ttk.Frame(nb); nb.add(self.tab_case, text="  接口用例  ")
        self.tab_scene = ttk.Frame(nb); nb.add(self.tab_scene, text="  组合场景（多接口）  ")
        self.tab_log = ttk.Frame(nb); nb.add(self.tab_log, text="  执行日志  ")
        self._build_case_tab()
        self._build_scene_tab()
        self._build_log_tab()

    def _build_case_tab(self):
        bar = ttk.Frame(self.tab_case, padding=(0, 6, 0, 2))
        bar.pack(fill="x")
        ttk.Button(bar, text="cURL 导入", style="Brand.TButton",
                   command=self.curl_import).pack(side="left", padx=(0, 6))
        for txt, cmd in (("+ 新增接口", self.add_case), ("编辑", self.edit_case),
                         ("复制", self.copy_case), ("删除", self.del_case)):
            ttk.Button(bar, text=txt, command=cmd).pack(side="left", padx=3)
        ttk.Button(bar, text="用勾选的接口组合成场景", style="Brand.TButton",
                   command=self.make_scene).pack(side="left", padx=(14, 0))
        self.lb_ccount = ttk.Label(bar, text="", foreground="#6b7280")
        self.lb_ccount.pack(side="right")

        bar2 = ttk.Frame(self.tab_case, padding=(0, 0, 0, 6))
        bar2.pack(fill="x")
        for txt, cmd in (("全选", lambda: self.sel_all(self.tv_case, True)),
                         ("取消全选", lambda: self.sel_all(self.tv_case, False)),
                         ("执行选中", self.run_selected_cases),
                         ("查看结果", self.view_case_result),
                         ("测试提取 / 点值生成路径", self.test_extract_current),
                         ("清空该接口回参", self.clear_case)):
            ttk.Button(bar2, text=txt, command=cmd).pack(side="left", padx=3)
        ttk.Label(bar2, foreground="#6b7280",
                  text="  提示：双击「接口名称」「接口用例」可直接改（回车保存，Tab 跳下一格），"
                       "双击其他列打开完整编辑，右键还有更多操作").pack(side="left")

        fbar = ttk.Frame(self.tab_case, padding=(0, 0, 0, 6))
        fbar.pack(fill="x")
        ttk.Label(fbar, text="筛选：").pack(side="left")
        ttk.Label(fbar, text="接口名称").pack(side="left")
        self.f_name = ttk.Combobox(fbar, width=22, state="readonly", values=["全部"])
        self.f_name.set("全部")
        self.f_name.pack(side="left", padx=(3, 10))
        ttk.Label(fbar, text="执行结果").pack(side="left")
        self.f_run = ttk.Combobox(fbar, width=10, state="readonly",
                                  values=["全部", "未执行", "已执行", "成功", "失败"])
        self.f_run.set("全部")
        self.f_run.pack(side="left", padx=(3, 10))
        ttk.Label(fbar, text="断言结果").pack(side="left")
        self.f_ass = ttk.Combobox(fbar, width=12, state="readonly",
                                  values=["全部", "通过", "不通过", "有警告", "未设置断言"])
        self.f_ass.set("全部")
        self.f_ass.pack(side="left", padx=(3, 10))
        ttk.Label(fbar, text="关键字").pack(side="left")
        self.f_kw = ttk.Entry(fbar, width=22)
        self.f_kw.pack(side="left", padx=3)
        ttk.Button(fbar, text="重置筛选", command=self.reset_filter).pack(side="left", padx=6)
        self.lb_filter = ttk.Label(fbar, foreground=BRAND)
        self.lb_filter.pack(side="right")
        for w in (self.f_name, self.f_run, self.f_ass):
            w.bind("<<ComboboxSelected>>", lambda e: self.refresh_cases(keep_scroll=True))
        self.f_kw.bind("<KeyRelease>", self._kw_changed)

        box = ttk.Frame(self.tab_case); box.pack(fill="both", expand=True)
        cols = ("sel", "id", "name", "case", "method", "url", "en", "res")
        self.tv_case = ttk.Treeview(box, columns=cols, show="headings")
        for c, t, w, a in (("sel", "选择", 46, "center"), ("id", "用例ID", 90, "w"),
                           ("name", "接口名称", 150, "w"), ("case", "接口用例", 180, "w"),
                           ("method", "方法", 62, "center"), ("url", "接口URL", 330, "w"),
                           ("en", "执行", 50, "center"), ("res", "最近结果", 150, "w")):
            if c == "sel":
                self.tv_case.heading(c, text=t)
            else:
                self.tv_case.heading(c, text=t, command=lambda col=c: self.sort_cases(col))
            self.tv_case.column(c, width=w, anchor=a)
        self.tv_case.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tv_case.yview)
        sb.pack(side="right", fill="y")
        self.tv_case.configure(yscrollcommand=sb.set)
        self.tv_case.tag_configure("ok", foreground=OKC)
        self.tv_case.tag_configure("err", foreground=ERRC)
        self.tv_case.bind("<Button-1>", self._toggle_click)
        self.tv_case.bind("<Double-1>", self._case_dbl)
        self.tv_case.bind("<Button-3>", self._case_menu)
        self.case_menu = tk.Menu(self, tearoff=0)
        for txt, cmd in (("修改接口名称", lambda: self.rename_case("#3")),
                         ("修改接口用例", lambda: self.rename_case("#4")),
                         (None, None),
                         ("编辑", self.edit_case), ("复制为新用例", self.copy_case),
                         ("执行这个接口", self.run_selected_cases),
                         ("测试提取 / 点值生成路径", self.test_extract_current),
                         ("查看结果", self.view_case_result), (None, None),
                         ("清空回参", self.clear_case), ("删除", self.del_case)):
            if txt is None:
                self.case_menu.add_separator()
            else:
                self.case_menu.add_command(label=txt, command=cmd)

    def _build_scene_tab(self):
        tip = ttk.Label(self.tab_scene, foreground="#475569", padding=(0, 6),
                        text="组合场景：按步骤顺序调用多个接口；前一个接口用「提取变量」取值，"
                             "后一个接口用 ${变量名} 引用。结果导出后写入 Excel 的 Sheet2（只保留最新一次）")
        tip.pack(anchor="w")
        bar = ttk.Frame(self.tab_scene); bar.pack(fill="x")
        for txt, cmd in (("全选", lambda: self.sel_all(self.tv_scene, True)),
                         ("取消全选", lambda: self.sel_all(self.tv_scene, False)),
                         ("+ 新建组合场景", self.add_scene), ("编辑", self.edit_scene),
                         ("删除", self.del_scene), ("执行选中", self.run_selected_scenes),
                         ("查看结果", self.view_scene_result), ("清空该场景结果", self.clear_scene)):
            ttk.Button(bar, text=txt, command=cmd).pack(side="left", padx=3)
        self.lb_scount = ttk.Label(bar, text="", foreground="#6b7280")
        self.lb_scount.pack(side="right")

        box = ttk.Frame(self.tab_scene); box.pack(fill="both", expand=True, pady=(6, 0))
        cols = ("sel", "name", "steps", "res")
        self.tv_scene = ttk.Treeview(box, columns=cols, show="headings")
        for c, t, w, a in (("sel", "选择", 46, "center"), ("name", "场景名称", 200, "w"),
                           ("steps", "步骤（按顺序）", 620, "w"), ("res", "最近结果", 130, "w")):
            self.tv_scene.heading(c, text=t)
            self.tv_scene.column(c, width=w, anchor=a)
        self.tv_scene.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tv_scene.yview)
        sb.pack(side="right", fill="y")
        self.tv_scene.configure(yscrollcommand=sb.set)
        self.tv_scene.tag_configure("ok", foreground=OKC)
        self.tv_scene.tag_configure("err", foreground=ERRC)
        self.tv_scene.bind("<Button-1>", self._toggle_click)
        self.tv_scene.bind("<Double-1>", self._scene_dbl)

    def _build_log_tab(self):
        bar = ttk.Frame(self.tab_log); bar.pack(fill="x", pady=6)
        ttk.Button(bar, text="清空日志",
                   command=lambda: self._set_log("")).pack(side="left")
        self.txt_log = ScrolledText(self.tab_log, font=("Consolas", 10), bg="#0f172a",
                                    fg="#cbd5e1", insertbackground="#cbd5e1")
        self.txt_log.pack(fill="both", expand=True)

    # ------------------------------------------------ 勾选
    def _toggle_click(self, event):
        tv = event.widget
        if tv.identify_region(event.x, event.y) != "cell":
            return
        if tv.identify_column(event.x) != "#1":
            return
        iid = tv.identify_row(event.y)
        if not iid:
            return
        cur = tv.set(iid, "sel")
        on = (cur != CHK_ON)
        tv.set(iid, "sel", CHK_ON if on else CHK_OFF)
        if tv is self.tv_case:
            (self.checked_ids.add if on else self.checked_ids.discard)(iid)
            self._update_filter_label()
        return "break"

    def sel_all(self, tv, flag):
        """只作用于当前可见（筛选后）的行。"""
        for iid in tv.get_children():
            tv.set(iid, "sel", CHK_ON if flag else CHK_OFF)
            if tv is self.tv_case:
                (self.checked_ids.add if flag else self.checked_ids.discard)(iid)
        if tv is self.tv_case:
            self._update_filter_label()

    def checked(self, tv):
        if tv is self.tv_case:
            # 勾选独立记录，筛选隐藏的行也算数
            alive = {c["id"] for c in core.STATE["cases"]}
            self.checked_ids &= alive
            return [c["id"] for c in core.STATE["cases"] if c["id"] in self.checked_ids]
        return [iid for iid in tv.get_children() if tv.set(iid, "sel") == CHK_ON]

    def _cur(self, tv):
        s = tv.selection()
        return s[0] if s else None

    # ------------------------------------------------ 列表刷新
    def _kw_changed(self, _e=None):
        if getattr(self, "_kw_job", None):
            try:
                self.after_cancel(self._kw_job)
            except Exception:
                pass
        self._kw_job = self.after(250, lambda: self.refresh_cases(keep_scroll=True))

    def reset_filter(self):
        self.f_name.set("全部")
        self.f_run.set("全部")
        self.f_ass.set("全部")
        self.f_kw.delete(0, "end")
        self.refresh_cases()

    def sort_cases(self, col):
        if getattr(self, "sort_col", None) == col:
            self.sort_desc = not getattr(self, "sort_desc", False)
        else:
            self.sort_col, self.sort_desc = col, False
        self.refresh_cases(keep_scroll=True)

    @staticmethod
    def _case_state(c):
        """返回 (执行状态, 断言状态)。"""
        rs = c.get("results") or []
        if not rs or rs[0].get("imported"):
            return "未执行", "未设置断言"
        r = rs[0]
        run = "成功" if r.get("ok") else "失败"
        rows = r.get("asserts") or []
        if not rows:
            ass = "未设置断言"
        else:
            fail = sum(1 for x in rows if not x.get("ok") and x.get("level") != "仅警告")
            warn = sum(1 for x in rows if not x.get("ok") and x.get("level") == "仅警告")
            ass = "不通过" if fail else ("有警告" if warn else "通过")
        return run, ass

    def _match_filter(self, c):
        name = self.f_name.get()
        if name and name != "全部" and c.get("name", "") != name:
            return False
        run, ass = self._case_state(c)
        fr = self.f_run.get()
        if fr == "未执行" and run != "未执行":
            return False
        if fr == "已执行" and run == "未执行":
            return False
        if fr in ("成功", "失败") and run != fr:
            return False
        fa = self.f_ass.get()
        if fa != "全部" and ass != fa:
            return False
        kw = self.f_kw.get().strip().lower()
        if kw:
            blob = " ".join([c.get("id", ""), c.get("name", ""), c.get("caseName", ""),
                             c.get("url", ""), c.get("method", "")]).lower()
            if kw not in blob:
                return False
        return True

    def _update_filter_label(self):
        total = len(core.STATE["cases"])
        shown = len(self.tv_case.get_children())
        n = len(self.checked_ids & {c["id"] for c in core.STATE["cases"]})
        txt = "显示 %d / 共 %d 条" % (shown, total)
        if n:
            txt += " · 已勾选 %d 条" % n
        self.lb_filter.config(text=txt)

    def refresh_cases(self, keep_scroll=False):
        tv = self.tv_case
        pos = tv.yview()[0] if keep_scroll else 0.0
        # 接口名称下拉的候选值
        names = []
        for c in core.STATE["cases"]:
            nm = c.get("name", "")
            if nm and nm not in names:
                names.append(nm)
        cur_name = self.f_name.get()
        self.f_name["values"] = ["全部"] + sorted(names)
        if cur_name not in self.f_name["values"]:
            self.f_name.set("全部")

        rows = [c for c in core.STATE["cases"] if self._match_filter(c)]
        col = getattr(self, "sort_col", None)
        if col:
            keyf = {"id": lambda c: c.get("id", ""), "name": lambda c: c.get("name", ""),
                    "case": lambda c: c.get("caseName", ""), "method": lambda c: c.get("method", ""),
                    "url": lambda c: c.get("url", ""),
                    "en": lambda c: c.get("enabled", True),
                    "res": lambda c: brief(c.get("results"))}.get(col)
            if keyf:
                rows = sorted(rows, key=lambda c: str(keyf(c)),
                              reverse=getattr(self, "sort_desc", False))

        tv.delete(*tv.get_children())
        for c in rows:
            rs = c.get("results", [])
            tag = ""
            if rs and not rs[0].get("imported"):
                tag = "ok" if (rs[0].get("ok") and rs[0].get("assertOk") is not False) else "err"
            tv.insert("", "end", iid=c["id"], tags=(tag,),
                      values=(CHK_ON if c["id"] in self.checked_ids else CHK_OFF,
                              c["id"], c.get("name", ""), c.get("caseName", ""),
                              c.get("method", ""), c.get("url", ""),
                              "Y" if c.get("enabled", True) else "N", brief(rs)))
        if keep_scroll:
            tv.yview_moveto(pos)
        self.lb_ccount.config(text="共 %d 个接口用例" % len(core.STATE["cases"]))
        self._update_filter_label()
        self.lb_file.config(text=("当前文件：" + core.STATE["fileName"])
                            if core.STATE.get("fileName") else "未导入文件")

    def refresh_scenes(self):
        keep = {iid: self.tv_scene.set(iid, "sel") for iid in self.tv_scene.get_children()}
        self.tv_scene.delete(*self.tv_scene.get_children())
        for s in core.STATE["scenarios"]:
            rs = s.get("results", [])
            tag = "ok" if (rs and rs[0].get("ok")) else ("err" if rs else "")
            steps = "  →  ".join("%d.%s" % (i + 1, st["caseId"])
                                 for i, st in enumerate(s.get("steps", [])))
            self.tv_scene.insert("", "end", iid=s["id"], tags=(tag,),
                                 values=(keep.get(s["id"], CHK_OFF), s.get("name", ""),
                                         steps, brief(rs)))
        self.lb_scount.config(text="共 %d 个组合场景" % len(core.STATE["scenarios"]))

    # ------------------------------------------------ 配置
    def load_config_to_ui(self):
        c = core.STATE["config"]
        self.e_base.delete(0, "end"); self.e_base.insert(0, c.get("baseUrl", ""))
        self.cb_enc.set(c.get("encoding", "utf-8"))
        self.e_to.delete(0, "end"); self.e_to.insert(0, str(c.get("timeout", 30)))
        self.v_ssl.set(bool(c.get("verifySsl")))
        self.v_stop.set(bool(c.get("stopOnAssertFail")))
        lines = []
        for h in c.get("headers", []):
            prefix = "" if h.get("enabled", True) else "# "
            lines.append("%s%s: %s" % (prefix, h.get("key", ""), h.get("value", "")))
        self.t_hdr.delete("1.0", "end")
        self.t_hdr.insert("1.0", "\n".join(lines))

    def collect_config(self):
        headers = []
        for line in self.t_hdr.get("1.0", "end").splitlines():
            raw = line.strip()
            if not raw:
                continue
            enabled = True
            if raw.startswith("#"):
                enabled = False
                raw = raw.lstrip("#").strip()
            if ":" not in raw:
                continue
            k, v = raw.split(":", 1)
            if k.strip():
                headers.append({"key": k.strip(), "value": v.strip(), "enabled": enabled})
        try:
            to = int(float(self.e_to.get().strip() or 30))
        except Exception:
            to = 30
        return {"baseUrl": self.e_base.get().strip(), "encoding": self.cb_enc.get() or "utf-8",
                "timeout": to, "verifySsl": self.v_ssl.get(),
                "stopOnAssertFail": self.v_stop.get(), "headers": headers}

    def save_config(self, silent=False):
        core.STATE["config"].update(self.collect_config())
        core.save_state()
        if not silent:
            self.status("公共配置已保存")

    # ------------------------------------------------ 状态/日志
    def status(self, text):
        self.lb_status.config(text=text)

    def _set_log(self, text):
        self.txt_log.delete("1.0", "end")
        if text:
            self.txt_log.insert("end", text)

    def log(self, line):
        self.logq.put(line)

    def _drain_log(self):
        try:
            while True:
                line = self.logq.get_nowait()
                self.txt_log.insert("end", line + "\n")
                self.txt_log.see("end")
        except queue.Empty:
            pass
        self.after(150, self._drain_log)

    # ------------------------------------------------ 执行
    def _run(self, case_ids=None, scene_ids=None, run_all=False, label=""):
        if self.running:
            messagebox.showinfo("提示", "正在执行中，请稍候")
            return
        self.save_config(silent=True)
        self.running = True
        self.status("⏳ 正在执行：%s" % label)
        self.log("===== %s  %s =====" % (label, core.now_str()))

        def worker():
            try:
                msg = core.execute(case_ids, scene_ids, run_all, self.log)
            except Exception as e:
                msg = "执行出错：%s" % e
                self.log(traceback.format_exc())
            self.after(0, done, msg)

        def done(msg):
            self.running = False
            self.refresh_cases(keep_scroll=True)
            self.refresh_scenes()
            self.status(msg)
            self.log(msg)

        threading.Thread(target=worker, daemon=True).start()

    def run_all(self):
        if not core.STATE["cases"] and not core.STATE["scenarios"]:
            messagebox.showinfo("提示", "还没有接口用例。请先「导入Excel」或「+ 新增接口」")
            return
        self._run(run_all=True, label="全部执行")

    def run_selected(self):
        cs, ss = self.checked(self.tv_case), self.checked(self.tv_scene)
        if not cs and not ss:
            messagebox.showinfo("提示", "请先在列表左侧「选择」列勾选要执行的接口或场景")
            return
        self._run(cs, ss, label="执行选中（接口%d个 / 场景%d个）" % (len(cs), len(ss)))

    def run_selected_cases(self):
        cs = self.checked(self.tv_case) or ([self._cur(self.tv_case)] if self._cur(self.tv_case) else [])
        if not cs:
            messagebox.showinfo("提示", "请先勾选或选中接口"); return
        self._run(cs, [], label="执行接口 %d 个" % len(cs))

    def run_selected_scenes(self):
        ss = self.checked(self.tv_scene) or ([self._cur(self.tv_scene)] if self._cur(self.tv_scene) else [])
        if not ss:
            messagebox.showinfo("提示", "请先勾选或选中场景"); return
        self._run([], ss, label="执行场景 %d 个" % len(ss))

    # ------------------------------------------------ 右键 / 复制 / cURL
    def _case_menu(self, event):
        iid = self.tv_case.identify_row(event.y)
        if not iid:
            return
        self.tv_case.selection_set(iid)
        try:
            self.case_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.case_menu.grab_release()

    def copy_case(self, cid=None):
        if cid:
            ids = [cid]
        else:
            cur = self._cur(self.tv_case)
            ids = [cur] if cur else self.checked(self.tv_case)
        if not ids:
            messagebox.showinfo("提示", "请先选中要复制的接口（单击某一行）")
            return
        made = []
        for i in ids:
            nc = core.duplicate_case(i)
            if nc:
                made.append(nc)
        self.refresh_cases()
        if made:
            last = made[-1]["id"]
            self.tv_case.selection_set(last)
            self.tv_case.see(last)
            self.status("已复制 %d 个：%s" % (len(made), "，".join(
                "%s（%s）" % (m["id"], m["caseName"]) for m in made)))

    def curl_import(self):
        def on_import(cases):
            # 先直接加入列表（不依赖用户点保存），再打开编辑窗口让他改名称
            ids = []
            for c in cases:
                c["id"] = core.next_case_id()
                c.setdefault("asserts", [])
                core.STATE["cases"].append(c)
                ids.append(c["id"])
            core.save_state()
            self.refresh_cases()
            if ids:
                self.tv_case.selection_set(ids[-1])
                self.tv_case.see(ids[-1])
            self.status("已通过 cURL 导入 %d 个接口：%s" % (len(ids), "、".join(ids)))

            if len(cases) == 1:
                cid = ids[0]
                idx = [i for i, c in enumerate(core.STATE["cases"]) if c["id"] == cid][0]

                def save(nc):
                    core.STATE["cases"][idx] = nc
                    core.save_state()
                    self.refresh_cases()
                    self.tv_case.selection_set(nc["id"])
                    self.status("已保存接口 %s" % nc["id"])
                CaseDialog(self, core.STATE["cases"][idx], save, is_new=False,
                           on_copy=lambda new_id: self.copy_case(new_id))
            else:
                messagebox.showinfo("导入完成",
                                    "已导入 %d 个接口：%s\n\n双击列表里的「接口名称」「接口用例」即可直接改"
                                    % (len(ids), "、".join(ids)))
        CurlDialog(self, on_import)

    def test_extract_current(self):
        cid = self._cur(self.tv_case)
        if not cid:
            messagebox.showinfo("提示", "请先单击选中一个接口")
            return

        def provider():
            c = core.find_case(cid)
            if not c:
                return None
            c = dict(c)
            c["extract"] = ""
            return c

        def apply(text):
            c = core.find_case(cid)
            if c is not None:
                c["extract"] = text
                core.save_state()
                self.status("已更新 %s 的提取变量" % cid)

        ExtractTestDialog(self, provider, (core.find_case(cid) or {}).get("extract", ""), apply)

    # ------------------------------------------------ 接口增删改
    def add_case(self):
        def save(nc):
            core.STATE["cases"].append(nc)
            core.save_state()
            self.refresh_cases()
        CaseDialog(self, None, save)

    def _case_dbl(self, event):
        col = self.tv_case.identify_column(event.x)
        if col == "#1":
            return
        iid = self.tv_case.identify_row(event.y)
        if iid and col in ("#3", "#4"):
            self._edit_cell(iid, col)
            return "break"
        self.edit_case()

    def _scene_dbl(self, event):
        col = self.tv_scene.identify_column(event.x)
        if col == "#1":
            return
        iid = self.tv_scene.identify_row(event.y)
        if iid and col == "#2":
            self._edit_scene_cell(iid)
            return "break"
        self.edit_scene()

    # ------------------------------------------------ 列表内就地编辑
    def _inline_entry(self, tv, iid, col, on_commit, on_tab=None):
        return inline_edit(tv, iid, col, on_commit, on_tab)

    def _edit_cell(self, iid, col):
        case = core.find_case(iid)
        if not case:
            return
        field = "name" if col == "#3" else "caseName"
        label = "接口名称" if field == "name" else "接口用例"

        def commit(text):
            if text == (case.get(field) or ""):
                return
            case[field] = text
            core.save_state()
            self.tv_case.set(iid, col, text)
            self.status("已修改 %s 的%s：%s" % (iid, label, text or "（空）"))

        self.nb_main.select(self.tab_case)
        nxt = (lambda: self._edit_cell(iid, "#4")) if col == "#3" else None
        ent = self._inline_entry(self.tv_case, iid, col, commit, nxt)
        if ent is None:          # 行不可见等异常情况，退回完整编辑窗口
            self.tv_case.selection_set(iid)
            self.edit_case()
        return ent

    def _edit_scene_cell(self, iid):
        sc = core.find_scene(iid)
        if not sc:
            return

        def commit(text):
            if not text or text == sc.get("name"):
                return
            sc["name"] = text
            core.save_state()
            self.tv_scene.set(iid, "#2", text)
            self.status("已修改场景名称：%s" % text)

        self.nb_main.select(self.tab_scene)
        ent = self._inline_entry(self.tv_scene, iid, "#2", commit)
        if ent is None:
            self.tv_scene.selection_set(iid)
            self.edit_scene()
        return ent

    def rename_case(self, col):
        cid = self._cur(self.tv_case)
        if not cid:
            messagebox.showinfo("提示", "请先单击选中一个接口")
            return
        self._edit_cell(cid, col)

    def edit_case(self):
        cid = self._cur(self.tv_case)
        if not cid:
            messagebox.showinfo("提示", "请先选中一行"); return
        idx = [i for i, c in enumerate(core.STATE["cases"]) if c["id"] == cid][0]

        def save(nc):
            core.STATE["cases"][idx] = nc
            core.save_state()
            self.refresh_cases()
        CaseDialog(self, core.STATE["cases"][idx], save, is_new=False,
                   on_copy=lambda new_id: self.copy_case(new_id))

    def del_case(self):
        ids = self.checked(self.tv_case) or ([self._cur(self.tv_case)] if self._cur(self.tv_case) else [])
        if not ids:
            return
        if not messagebox.askyesno("确认", "确定删除选中的 %d 个接口用例？" % len(ids)):
            return
        core.STATE["cases"] = [c for c in core.STATE["cases"] if c["id"] not in ids]
        core.save_state()
        self.refresh_cases()

    # ------------------------------------------------ 场景增删改
    def make_scene(self):
        ids = self.checked(self.tv_case)
        if len(ids) < 2:
            messagebox.showinfo("提示", "请先在「选择」列勾选 2 个及以上接口，再点这里组合")
            return
        order = [c["id"] for c in core.STATE["cases"] if c["id"] in ids]
        self.add_scene(order)

    def add_scene(self, preset=None):
        def save(ns):
            core.STATE["scenarios"].append(ns)
            core.save_state()
            self.refresh_scenes()
        SceneDialog(self, None, save, preset_ids=preset)

    def edit_scene(self):
        sid = self._cur(self.tv_scene)
        if not sid:
            messagebox.showinfo("提示", "请先选中一个场景"); return
        idx = [i for i, s in enumerate(core.STATE["scenarios"]) if s["id"] == sid][0]

        def save(ns):
            core.STATE["scenarios"][idx] = ns
            core.save_state()
            self.refresh_scenes()
        SceneDialog(self, core.STATE["scenarios"][idx], save)

    def del_scene(self):
        ids = self.checked(self.tv_scene) or ([self._cur(self.tv_scene)] if self._cur(self.tv_scene) else [])
        if not ids:
            return
        if not messagebox.askyesno("确认", "确定删除选中的 %d 个场景？" % len(ids)):
            return
        core.STATE["scenarios"] = [s for s in core.STATE["scenarios"] if s["id"] not in ids]
        core.save_state()
        self.refresh_scenes()

    # ------------------------------------------------ 结果
    def view_case_result(self):
        cid = self._cur(self.tv_case)
        if not cid:
            messagebox.showinfo("提示", "请先选中一行"); return
        c = core.find_case(cid)
        rs = c.get("results", [])
        if not rs:
            messagebox.showinfo("提示", "该接口还没有执行结果"); return
        pages = [("第%d次%s" % (i + 1, "(最新)" if i == 0 else ""),
                  fmt_result(r, full=False), fmt_result(r, full=True))
                 for i, r in enumerate(rs)]
        TextDialog(self, "接口 %s 的执行结果（最多保留5次）" % cid, pages)

    def view_scene_result(self):
        sid = self._cur(self.tv_scene)
        if not sid:
            messagebox.showinfo("提示", "请先选中一个场景"); return
        s = core.find_scene(sid)
        rs = s.get("results", [])
        if not rs:
            messagebox.showinfo("提示", "该场景还没有执行结果"); return
        pages = []
        for i, r in enumerate(rs):
            def build(full):
                out = ["场景执行时间：%s　总体：%s" % (r.get("time", ""),
                                                    "成功" if r.get("ok") else "失败"), ""]
                for k, st in enumerate(r.get("steps", []), 1):
                    out.append("──── 步骤 %d ： %s %s ────"
                               % (k, st.get("caseId", ""), st.get("caseName", "")))
                    out.append(fmt_result(st, full=full))
                    out.append("")
                return "\n".join(out)
            pages.append(("第%d次%s" % (i + 1, "(最新)" if i == 0 else ""),
                          build(False), build(True)))
        TextDialog(self, "场景「%s」的执行结果" % s.get("name", ""), pages)

    # ------------------------------------------------ 清空
    def clear_all(self):
        cnt = max([0] + [len(c.get("results", [])) for c in core.STATE["cases"]]
                  + [len(s.get("results", [])) for s in core.STATE["scenarios"]])
        if not cnt:
            messagebox.showinfo("提示", "当前没有可清空的执行结果"); return

        def ok(idx):
            core.clear_results("all", None, idx)
            self.refresh_cases(); self.refresh_scenes()
            self.status("已清空执行结果")
        ClearDialog(self, "全部接口 + 全部场景的执行结果", cnt, ok)

    def clear_case(self):
        cid = self._cur(self.tv_case)
        if not cid:
            messagebox.showinfo("提示", "请先选中一行"); return
        c = core.find_case(cid)
        cnt = len(c.get("results", []))
        if not cnt:
            messagebox.showinfo("提示", "该接口没有可清空的回参"); return

        def ok(idx):
            core.clear_results("case", cid, idx)
            self.refresh_cases()
            self.status("已清空接口 %s 的回参" % cid)
        ClearDialog(self, "接口 %s 的回参" % cid, cnt, ok)

    def clear_scene(self):
        sid = self._cur(self.tv_scene)
        if not sid:
            messagebox.showinfo("提示", "请先选中一个场景"); return
        s = core.find_scene(sid)
        cnt = len(s.get("results", []))
        if not cnt:
            messagebox.showinfo("提示", "该场景没有可清空的结果"); return

        def ok(idx):
            core.clear_results("scenario", sid, idx)
            self.refresh_scenes()
            self.status("已清空场景「%s」的结果" % s.get("name", ""))
        ClearDialog(self, "场景「%s」的结果" % s.get("name", ""), cnt, ok)

    # ------------------------------------------------ Excel
    def do_template(self):
        path = filedialog.asksaveasfilename(
            title="保存Excel模板", defaultextension=".xlsx",
            initialfile="接口测试模板.xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")])
        if not path:
            return
        try:
            core.build_template(path)
            self.status("模板已保存：%s" % path)
            if messagebox.askyesno("完成", "模板已保存到：\n%s\n\n是否现在打开它？" % path):
                self._open_path(path)
        except Exception as e:
            messagebox.showerror("出错", "生成模板失败：%s" % e)

    def do_import(self):
        path = filedialog.askopenfilename(title="选择接口测试Excel",
                                          filetypes=[("Excel 工作簿", "*.xlsx *.xlsm")])
        if not path:
            return
        mode = "replace"
        if core.STATE["cases"]:
            r = messagebox.askyesnocancel(
                "导入方式", "当前已有 %d 条用例。\n\n是(Y)＝覆盖替换\n否(N)＝追加到现有列表\n取消＝不导入"
                % len(core.STATE["cases"]))
            if r is None:
                return
            mode = "replace" if r else "append"
        try:
            msg = core.import_excel(path, mode)
        except Exception as e:
            messagebox.showerror("导入失败", "解析该文件失败：\n%s" % e)
            return
        self.load_config_to_ui()
        self.refresh_cases()
        self.refresh_scenes()
        self.status(msg)
        messagebox.showinfo("导入完成", msg)

    def do_export(self):
        if not core.STATE["cases"]:
            messagebox.showinfo("提示", "还没有可导出的内容"); return
        self.save_config(silent=True)
        core.flush_state()
        import time as _t
        path = filedialog.asksaveasfilename(
            title="导出执行结果", defaultextension=".xlsx",
            initialfile="接口测试结果_%s.xlsx" % _t.strftime("%Y%m%d_%H%M%S"),
            filetypes=[("Excel 工作簿", "*.xlsx")])
        if not path:
            return
        try:
            core.export_excel(path)
        except Exception as e:
            messagebox.showerror("导出失败", str(e)); return
        self.status("已导出：%s" % path)
        if messagebox.askyesno("导出完成", "已导出到：\n%s\n\n是否现在打开？" % path):
            self._open_path(path)

    def _open_path(self, path):
        try:
            if os.name == "nt":
                os.startfile(path)
            else:
                import subprocess
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            messagebox.showinfo("提示", "无法自动打开，请手动到该路径查看：\n%s" % path)

    def reset_all(self):
        if not messagebox.askyesno("确认", "将清空当前所有用例、场景和执行结果，确定？"):
            return
        core.STATE.clear()
        core.STATE.update(core.default_state())
        core.save_state()
        self.load_config_to_ui()
        self.refresh_cases()
        self.refresh_scenes()
        self.status("已重置")

    def _close(self):
        try:
            self.save_config(silent=True)
            core.flush_state()
        except Exception:
            pass
        self.destroy()


def main():
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    app = App()
    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        err = traceback.format_exc()
        try:
            with open(os.path.join(core.DATA_DIR, "error.log"), "w", encoding="utf-8") as f:
                f.write(err)
        except Exception:
            pass
        try:
            root = tk.Tk(); root.withdraw()
            messagebox.showerror("启动失败", err[-1500:])
        except Exception:
            print(err)
