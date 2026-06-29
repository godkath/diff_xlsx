#!/usr/bin/env python3
"""
diff_xlsx_gui.py - 鼠标操作的 xls/xlsx 差异对比工具
- 鼠标选两个 xls/xlsx 文件
- 鼠标点选 sheet
- 鼠标点"开始对比"
- 完成后可在 Text 区查看报告，或点"保存 Excel 差异表"生成 .xlsx

主键: 自动找 ZJHM 证件号码列 (可配置)

用法: python3 diff_xlsx_gui.py
"""
import os
import sys
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

# 复用 diff_xlsx 核心函数
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from diff_xlsx import (  # noqa: E402
    load_any, list_sheets, build_mapping, cell, find_col,
    KEY_CANDIDATES, generate_diff_xlsx
)


def do_diff(file1: str, sheet1: str, file2: str, sheet2: str,
            label1: str, label2: str, q: queue.Queue):
    """实际对比逻辑 (后台线程), 同时把数据返回以便生成 Excel"""
    try:
        q.put(("status", f"📂 读取 {os.path.basename(file1)} (sheet={sheet1})"))
        h1, r1, k1 = load_any(file1, sheet1)
        q.put(("status", f"  ✅ 读取完成: {len(r1)} 条, 主键={h1[k1]}"))

        q.put(("status", f"📂 读取 {os.path.basename(file2)} (sheet={sheet2})"))
        h2, r2, k2 = load_any(file2, sheet2)
        q.put(("status", f"  ✅ 读取完成: {len(r2)} 条, 主键={h2[k2]}"))

        # 1. 人员差异
        s1, s2 = set(r1), set(r2)
        common = s1 & s2

        # 2. 字段结构差异
        mapping = build_mapping(h1, h2)
        f1_only = [h1[i] for i in sorted(mapping) if mapping[i] is None and h1[i]]
        f2_only = [h for h in h2
                   if h and h.upper() not in {(c or "").upper() for c in h1 if c}]

        # 3. 全字段 cell 比对
        diff_count = {}
        diff_samples = {}

        for f1i, f2i in mapping.items():
            if f2i is None:
                continue
            f1c = h1[f1i]
            cnt = 0
            samples = []
            for key in common:
                v1 = cell(r1[key][f1i])
                v2 = cell(r2[key][f2i])
                if v1 != v2:
                    cnt += 1
                    if len(samples) < 3:
                        samples.append((key, v1, v2))
            if cnt > 0:
                diff_count[f1c] = cnt
                diff_samples[f1c] = samples

        # 4. 生成报告
        out = []
        out.append("=" * 70)
        out.append(f"=== {label1}  vs  {label2} ===")
        out.append("=" * 70)
        out.append(f"{label1}: {file1} (sheet={sheet1})")
        out.append(f"{label2}: {file2} (sheet={sheet2})\n")
        out.append(f"{label1} 列数={len(h1)} (主键 {h1[k1]} 在第 {k1} 列), 记录数={len(r1)}")
        out.append(f"{label2} 列数={len(h2)} (主键 {h2[k2]} 在第 {k2} 列), 记录数={len(r2)}\n")

        out.append("--- 人员差异（按主键集合）---")
        out.append(f"  共有: {len(common)}")
        out.append(f"  仅 {label1} 有: {len(s1 - s2)}")
        out.append(f"  仅 {label2} 有: {len(s2 - s1)}")

        n1 = find_col(h1, ["XM", "姓名", "NAME", "name"])
        n2 = find_col(h2, ["XM", "姓名", "NAME", "name"])

        if s1 - s2:
            out.append(f"\n  [仅 {label1} 有] (n={len(s1 - s2)}):")
            for k in sorted(s1 - s2)[:20]:
                name = cell(r1[k][n1]) if n1 is not None else ""
                out.append(f"    {k}  {name}")
            if len(s1 - s2) > 20:
                out.append(f"    ... 共 {len(s1 - s2)} 条")

        if s2 - s1:
            out.append(f"\n  [仅 {label2} 有] (n={len(s2 - s1)}):")
            for k in sorted(s2 - s1)[:20]:
                name = cell(r2[k][n2]) if n2 is not None else ""
                out.append(f"    {k}  {name}")
            if len(s2 - s1) > 20:
                out.append(f"    ... 共 {len(s2 - s1)} 条")

        out.append(f"\n--- 字段结构差异 ---")
        out.append(f"  {label1} 独有列 ({len(f1_only)}): {f1_only}")
        out.append(f"  {label2} 独有列 ({len(f2_only)}): {f2_only}")

        if not common:
            out.append("\n无共有记录，跳过字段比对")
        else:
            out.append(f"\n--- 共有 {len(common)} 条记录的字段差异（按差异条数排序）---\n")
            if not diff_count:
                out.append("  ✅ 所有可比字段值完全一致")
            for k, v in sorted(diff_count.items(), key=lambda x: -x[1]):
                out.append(f"  {k}: {v} 条")
                for s in diff_samples[k]:
                    out.append(f"    例: 主键={s[0]}  {label1}='{s[1]}'  {label2}='{s[2]}'")
                out.append("")

        out.append("=" * 70)
        out.append("【总结】")
        out.append(f"  人员: {label1}={len(r1)}  {label2}={len(r2)}  差 {abs(len(r1) - len(r2))} 人")
        out.append(f"  字段: {label1}={len(h1)} 列  {label2}={len(h2)} 列  {label1}独有 {len(f1_only)} {label2}独有 {len(f2_only)}")
        out.append(f"  数据: {len(diff_count)} 个字段有差异（总计 {sum(diff_count.values())} cell）")
        out.append("=" * 70)

        report = "\n".join(out)
        # 缓存数据供生成 Excel
        ctx = {
            "file1": file1, "sheet1": sheet1, "file2": file2, "sheet2": sheet2,
            "label1": label1, "label2": label2,
            "h1": h1, "r1": r1, "k1": k1, "h2": h2, "r2": r2, "k2": k2,
        }
        q.put(("done", report, ctx))

    except Exception as e:
        import traceback
        q.put(("error", f"{e}\n\n{traceback.format_exc()}"))


class DiffApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("xls/xlsx 差异对比工具 (主键: ZJHM 证件号码)")
        root.geometry("900x700")

        self.q: queue.Queue = queue.Queue()
        self.running = False
        # 缓存最后一次对比的数据, 用于生成 Excel
        self.last_ctx = None  # type: ignore

        self._build_ui()
        self._poll_queue()

    def _build_ui(self):
        # === 文件 1 区 ===
        f1 = ttk.LabelFrame(self.root, text="📁 文件 1 (基准)", padding=10)
        f1.pack(fill="x", padx=10, pady=5)

        ttk.Label(f1, text="路径:").grid(row=0, column=0, sticky="w")
        self.path1_var = tk.StringVar()
        ttk.Entry(f1, textvariable=self.path1_var, width=70).grid(row=0, column=1, padx=5)
        ttk.Button(f1, text="浏览…", command=self._pick_file1).grid(row=0, column=2)
        ttk.Label(f1, text="标签:").grid(row=1, column=0, sticky="w", pady=5)
        self.label1_var = tk.StringVar(value="F1")
        ttk.Entry(f1, textvariable=self.label1_var, width=20).grid(row=1, column=1, sticky="w", padx=5)
        ttk.Label(f1, text="Sheet:").grid(row=2, column=0, sticky="w")
        self.sheet1_var = tk.StringVar()
        self.sheet1_combo = ttk.Combobox(f1, textvariable=self.sheet1_var, width=67, state="readonly")
        self.sheet1_combo.grid(row=2, column=1, padx=5, pady=5)

        # === 文件 2 区 ===
        f2 = ttk.LabelFrame(self.root, text="📁 文件 2 (对照)", padding=10)
        f2.pack(fill="x", padx=10, pady=5)

        ttk.Label(f2, text="路径:").grid(row=0, column=0, sticky="w")
        self.path2_var = tk.StringVar()
        ttk.Entry(f2, textvariable=self.path2_var, width=70).grid(row=0, column=1, padx=5)
        ttk.Button(f2, text="浏览…", command=self._pick_file2).grid(row=0, column=2)
        ttk.Label(f2, text="标签:").grid(row=1, column=0, sticky="w", pady=5)
        self.label2_var = tk.StringVar(value="F2")
        ttk.Entry(f2, textvariable=self.label2_var, width=20).grid(row=1, column=1, sticky="w", padx=5)
        ttk.Label(f2, text="Sheet:").grid(row=2, column=0, sticky="w")
        self.sheet2_var = tk.StringVar()
        self.sheet2_combo = ttk.Combobox(f2, textvariable=self.sheet2_var, width=67, state="readonly")
        self.sheet2_combo.grid(row=2, column=1, padx=5, pady=5)

        # === 控制区 ===
        ctrl = ttk.Frame(self.root, padding=10)
        ctrl.pack(fill="x", padx=10)
        self.run_btn = ttk.Button(ctrl, text="▶ 开始对比", command=self._run)
        self.run_btn.pack(side="left")
        ttk.Button(ctrl, text="📊 保存 Excel 差异表", command=self._save_xlsx).pack(side="left", padx=5)
        ttk.Button(ctrl, text="📝 保存报告 (txt)", command=self._save_txt).pack(side="left", padx=5)
        ttk.Button(ctrl, text="🗑 清空", command=self._clear).pack(side="left", padx=5)

        # === 状态 + 进度 ===
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(ctrl, textvariable=self.status_var, foreground="blue").pack(side="left", padx=20)
        self.progress = ttk.Progressbar(ctrl, mode="indeterminate", length=200)
        self.progress.pack(side="right")

        # === 报告区 ===
        rep = ttk.LabelFrame(self.root, text="📊 对比报告", padding=5)
        rep.pack(fill="both", expand=True, padx=10, pady=5)
        self.report_text = tk.Text(rep, wrap="none", font=("Menlo", 11))
        self.report_text.pack(fill="both", expand=True)
        sb_y = ttk.Scrollbar(self.report_text, orient="vertical", command=self.report_text.yview)
        sb_y.pack(side="right", fill="y")
        self.report_text.config(yscrollcommand=sb_y.set)

    def _pick_file1(self):
        p = filedialog.askopenfilename(
            title="选择文件 1",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("全部", "*.*")],
        )
        if p:
            self.path1_var.set(p)
            self._refresh_sheets(p, self.sheet1_combo, self.sheet1_var)

    def _pick_file2(self):
        p = filedialog.askopenfilename(
            title="选择文件 2",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("全部", "*.*")],
        )
        if p:
            self.path2_var.set(p)
            self._refresh_sheets(p, self.sheet2_combo, self.sheet2_var)

    def _refresh_sheets(self, path: str, combo: ttk.Combobox, var: tk.StringVar):
        names = list_sheets(path)
        combo["values"] = names
        if names and not str(names[0]).startswith("⚠️"):
            var.set(names[0])

    def _clear(self):
        self.path1_var.set("")
        self.path2_var.set("")
        self.sheet1_combo["values"] = []
        self.sheet2_combo["values"] = []
        self.sheet1_var.set("")
        self.sheet2_var.set("")
        self.report_text.delete("1.0", "end")
        self.status_var.set("就绪")
        self.last_ctx = None

    def _run(self):
        if self.running:
            messagebox.showinfo("提示", "正在对比中，请稍候…")
            return
        f1 = self.path1_var.get().strip()
        f2 = self.path2_var.get().strip()
        s1 = self.sheet1_var.get().strip()
        s2 = self.sheet2_var.get().strip()
        l1 = self.label1_var.get().strip() or "F1"
        l2 = self.label2_var.get().strip() or "F2"

        if not f1 or not os.path.isfile(f1):
            messagebox.showerror("错误", f"文件 1 不存在:\n{f1}")
            return
        if not f2 or not os.path.isfile(f2):
            messagebox.showerror("错误", f"文件 2 不存在:\n{f2}")
            return
        if not s1:
            messagebox.showerror("错误", "请选择文件 1 的 sheet")
            return
        if not s2:
            messagebox.showerror("错误", "请选择文件 2 的 sheet")
            return

        self.report_text.delete("1.0", "end")
        self.status_var.set("🔄 对比中…")
        self.last_ctx = None
        self.running = True
        self.run_btn.config(state="disabled")
        self.progress.start(10)

        threading.Thread(
            target=do_diff,
            args=(f1, s1, f2, s2, l1, l2, self.q),
            daemon=True,
        ).start()

    def _save_xlsx(self):
        if not self.last_ctx:
            messagebox.showinfo("提示", "请先跑一次对比")
            return
        default_name = f"diff_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        p = filedialog.asksaveasfilename(
            title="保存 Excel 差异表",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel", "*.xlsx"), ("全部", "*.*")],
        )
        if not p:
            return
        try:
            ctx = self.last_ctx
            generate_diff_xlsx(
                p, ctx["file1"], ctx["sheet1"], ctx["file2"], ctx["sheet2"],
                ctx["label1"], ctx["label2"],
                ctx["h1"], ctx["r1"], ctx["k1"], ctx["h2"], ctx["r2"], ctx["k2"],
            )
            messagebox.showinfo("已保存", f"Excel 差异表已保存到:\n{p}\n\n包含 4 个 sheet:\n  · 人员差异\n  · 字段结构差异\n  · 字段值差异 (主键/字段/F1/F2)\n  · 总结")
        except Exception as e:
            import traceback
            messagebox.showerror("错误", f"保存失败:\n{e}\n\n{traceback.format_exc()}")

    def _save_txt(self):
        content = self.report_text.get("1.0", "end").strip()
        if not content:
            messagebox.showinfo("提示", "报告为空")
            return
        default_name = f"diff_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        p = filedialog.asksaveasfilename(
            title="保存报告",
            defaultextension=".txt",
            initialfile=default_name,
            filetypes=[("文本", "*.txt"), ("Markdown", "*.md"), ("全部", "*.*")],
        )
        if p:
            with open(p, "w", encoding="utf-8") as f:
                f.write(content)
            messagebox.showinfo("已保存", f"报告已保存到:\n{p}")

    def _poll_queue(self):
        try:
            while True:
                item = self.q.get_nowait()
                kind = item[0]
                if kind == "status":
                    self.status_var.set(item[1])
                elif kind == "done":
                    _, report, ctx = item
                    self.report_text.delete("1.0", "end")
                    self.report_text.insert("1.0", report)
                    self.status_var.set(f"✅ 完成 - {len(report)} 字符 (可保存 Excel)")
                    self.last_ctx = ctx
                    self.running = False
                    self.run_btn.config(state="normal")
                    self.progress.stop()
                elif kind == "error":
                    self.report_text.delete("1.0", "end")
                    self.report_text.insert("1.0", f"❌ 错误:\n{item[1]}")
                    self.status_var.set("❌ 出错")
                    self.running = False
                    self.run_btn.config(state="normal")
                    self.progress.stop()
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)


def main():
    root = tk.Tk()
    try:
        root.tk.call("tk", "scaling", 1.2)
    except Exception:
        pass
    app = DiffApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()