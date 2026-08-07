# -*- coding: utf-8 -*-
"""
湍流参数计算工具
================

基于《湍流参数计算流程.md》中的公式，计算 CFD 湍流模型入口/初始参数。

支持的湍流模型：
  * k-ε              入口需给定 k、ε（或 I 与混合长度）
  * k-ω / SST        入口需给定 k、ω
  * Spalart-Allmaras 入口需给定 nuTilda

核心公式：
  I      = 0.16·Re^(-1/8)                     (步骤0，可选)
  k      = 1.5·(|U|·I)^2                      (步骤1)
  L      = 0.07·Lc                            (步骤2)
  ε      = Cmu^0.75·k^1.5/L                   (步骤3)
  ω      = √k/(Cmu^0.25·L) = ε/(Cmu·k)       (步骤4)
  nuTilda ≈ nut = Cmu·k²/ε，或 3ν/5ν/10ν，或 ν·√(0.01·Re)  (步骤5)
  nut    = Cmu·k²/ε (k-ε) | k/ω (k-ω) | nuTilda·fv1 (SA)    (步骤6)
  nutRatio = nut/ν                            (步骤7)
"""

import math
import tkinter as tk
from tkinter import ttk, messagebox

APP_TITLE = "湍流参数计算工具"
SUBTITLE = "支持 k-ε / k-ω / Spalart-Allmaras"

# ---- 配色方案（ttk 主题）----
C_BG = "#f0f4f8"          # 窗口背景
C_PANEL = "#ffffff"       # 面板背景
C_PRIMARY = "#2563eb"     # 主色（蓝）
C_PRIMARY_H = "#1d4ed8"   # 主色悬停
C_PRIMARY_P = "#1e40af"   # 主色按下
TEXT = "#1f2937"          # 正文
MUTED = "#64748b"         # 次要文字
BORDER = "#cbd5e1"        # 边框
ENTRY_BG = "#f8fafc"      # 输入框背景
BTN_BG = "#e5e7eb"        # 普通按钮背景
BTN_HOVER = "#d1d5db"     # 按钮悬停
BTN_PRESSED = "#9ca3af"   # 按钮按下
ACCENT = "#0f766e"        # 结果强调色（青）


def fmt(x):
    """数值格式化：过大/过小用科学计数法，否则保留 5 位有效数字。"""
    if x is None:
        return "—"
    if x == 0:
        return "0"
    a = abs(x)
    if a >= 1e5 or a < 1e-4:
        return f"{x:.4e}"
    return f"{x:.5g}"


class TurbulenceCalcApp(tk.Tk):
    """湍流参数计算 GUI。"""

    MODEL_NOTES = {
        "keps":   "k-ε 模型：入口给定 k、ε（或 I + ε）",
        "komega": "k-ω / SST 模型：入口给定 k、ω",
        "sa":     "Spalart-Allmaras 模型：入口给定 nuTilda",
    }

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("940x780")
        self.minsize(880, 720)
        self.configure(bg=C_BG)

        # ---- 输入变量 ----
        self.var_U = tk.StringVar(value="10")
        self.var_Lc = tk.StringVar(value="0.5")
        self.var_nu = tk.StringVar(value="1e-5")
        self.var_I = tk.StringVar(value="0.05")
        self.var_I_mode = tk.StringVar(value="given")      # given / re
        self.var_Re = tk.StringVar(value="")
        self.var_Cmu = tk.StringVar(value="0.09")
        self.var_model = tk.StringVar(value="keps")        # keps / komega / sa
        self.var_Cv1 = tk.StringVar(value="7.1")
        self.var_nuT_method = tk.StringVar(value="exact")  # exact/3nu/5nu/10nu/re/custom
        self.var_nuTilda_custom = tk.StringVar(value="3e-5")

        self.results = {}

        self._apply_style()
        self._build_ui()
        self._toggle_states()

    # ------------------------------------------------------------- 样式
    def _apply_style(self):
        """应用统一的自定义 ttk 主题样式。"""
        style = ttk.Style(self)
        style.theme_use("clam")

        # 基础默认
        style.configure(".", font=("Microsoft YaHei", 10),
                        background=C_BG, foreground=TEXT)
        style.configure("TFrame", background=C_BG)
        style.configure("Panel.TFrame", background=C_PANEL)

        # 标签
        style.configure("TLabel", background=C_PANEL, foreground=TEXT)
        style.configure("Title.TLabel", background=C_BG, foreground=TEXT,
                        font=("Microsoft YaHei", 17, "bold"))
        style.configure("Subtitle.TLabel", background=C_BG, foreground=MUTED,
                        font=("Microsoft YaHei", 9))
        style.configure("Value.TLabel", background=C_PANEL, foreground=ACCENT,
                        font=("Microsoft YaHei", 10, "bold"))

        # 面板（LabelFrame）
        style.configure("TLabelframe", background=C_PANEL, bordercolor=BORDER,
                        relief="solid", borderwidth=1, padding=12)
        style.configure("TLabelframe.Label", background=C_PANEL, foreground=TEXT,
                        font=("Microsoft YaHei", 10, "bold"))

        # 输入框（聚焦时主色边框高亮）
        style.configure("TEntry", fieldbackground=ENTRY_BG, foreground=TEXT,
                        bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
                        insertcolor=TEXT, padding=5)
        style.map("TEntry",
                  bordercolor=[("focus", C_PRIMARY)],
                  lightcolor=[("focus", C_PRIMARY)],
                  darkcolor=[("focus", C_PRIMARY)])

        # 按钮
        style.configure("TButton", background=BTN_BG, foreground=TEXT,
                        bordercolor=BORDER, padding=(16, 7),
                        font=("Microsoft YaHei", 10))
        style.map("TButton",
                  background=[("active", BTN_HOVER), ("pressed", BTN_PRESSED)],
                  bordercolor=[("active", BORDER)])
        style.configure("Accent.TButton", background=C_PRIMARY, foreground="white",
                        bordercolor=C_PRIMARY, padding=(24, 7),
                        font=("Microsoft YaHei", 10, "bold"))
        style.map("Accent.TButton",
                  background=[("active", C_PRIMARY_H), ("pressed", C_PRIMARY_P)],
                  bordercolor=[("active", C_PRIMARY_H)])

        # 单选按钮（选中时主色圆点）
        style.configure("TRadiobutton", background=C_PANEL, foreground=TEXT,
                        padding=(2, 2))
        style.map("TRadiobutton",
                  background=[("active", C_PANEL)],
                  indicatorcolor=[("selected", C_PRIMARY)])

    # ---------------------------------------------------------------- UI
    def _build_ui(self):
        title = ttk.Label(self, text=APP_TITLE, style="Title.TLabel")
        title.pack(pady=(12, 0))
        sub = ttk.Label(self, text=SUBTITLE, style="Subtitle.TLabel")
        sub.pack(pady=(0, 6))

        body = ttk.Frame(self, padding=(12, 4))
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(2, weight=1)

        self._build_input_panel(body).grid(row=0, column=0, sticky="nsew",
                                           padx=(0, 6), pady=(0, 6))
        self._build_result_panel(body).grid(row=0, column=1, sticky="nsew",
                                            padx=(6, 0), pady=(0, 6))
        self._build_button_bar(body).grid(row=1, column=0, columnspan=2,
                                          sticky="ew", pady=(0, 6))
        self._build_verify_panel(body).grid(row=2, column=0, columnspan=2,
                                            sticky="nsew")

    def _build_input_panel(self, parent):
        inp = ttk.LabelFrame(parent, text="输入参数", padding=12)
        inp.columnconfigure(1, weight=1)
        inp.columnconfigure(2, minsize=72)

        r = 0
        self._row_entry(inp, r, "平均流速 |U| (m/s)", self.var_U); r += 1
        self._row_entry(inp, r, "特征长度 Lc (m)", self.var_Lc); r += 1
        self._row_entry(inp, r, "运动粘度 ν (m²/s)", self.var_nu); r += 1

        # 湍流强度 I：直接给定 或 由 Re 估算
        ttk.Label(inp, text="湍流强度 I").grid(row=r, column=0, sticky="w", pady=4)
        fm_i = ttk.Frame(inp, style="Panel.TFrame")
        fm_i.grid(row=r, column=1, sticky="ew", pady=4)
        ttk.Radiobutton(fm_i, text="直接给定", value="given", variable=self.var_I_mode,
                        command=self._toggle_states).pack(side="left")
        ttk.Radiobutton(fm_i, text="由 Re 估算", value="re", variable=self.var_I_mode,
                        command=self._toggle_states).pack(side="left", padx=(10, 0))
        self.ent_I = ttk.Entry(inp, textvariable=self.var_I, width=10)
        self.ent_I.grid(row=r, column=2, sticky="e", pady=4)
        r += 1

        # 雷诺数 Re（可选，留空自动计算）
        self._row_entry(inp, r, "雷诺数 Re（留空自动算）", self.var_Re); r += 1
        self._row_entry(inp, r, "模型系数 Cmu", self.var_Cmu); r += 1

        # 湍流模型
        ttk.Label(inp, text="湍流模型").grid(row=r, column=0, sticky="w", pady=4)
        fm_m = ttk.Frame(inp, style="Panel.TFrame")
        fm_m.grid(row=r, column=1, columnspan=2, sticky="w", pady=4)
        for txt, val in (("k-ε", "keps"), ("k-ω", "komega"),
                         ("Spalart-Allmaras", "sa")):
            ttk.Radiobutton(fm_m, text=txt, value=val, variable=self.var_model,
                            command=self._toggle_states).pack(side="left", padx=(0, 12))
        r += 1

        # SA 模型系数 Cv1
        self.ent_Cv1 = self._row_entry(inp, r, "SA 模型系数 Cv1", self.var_Cv1)
        r += 1

        # nuTilda 估算方法
        ttk.Label(inp, text="nuTilda 估算").grid(row=r, column=0, sticky="w", pady=4)
        fm_n = ttk.Frame(inp, style="Panel.TFrame")
        fm_n.grid(row=r, column=1, columnspan=2, sticky="w", pady=4)
        opts = [("精确 (≈nut)", "exact"), ("3ν", "3nu"), ("5ν", "5nu"),
                ("10ν", "10nu"), ("由 Re", "re"), ("自定义", "custom")]
        for i, (txt, val) in enumerate(opts):
            ttk.Radiobutton(fm_n, text=txt, value=val, variable=self.var_nuT_method,
                            command=self._toggle_states).grid(row=i // 3, column=i % 3,
                                                              sticky="w", padx=(0, 16), pady=2)
        r += 1

        # 自定义 nuTilda
        ttk.Label(inp, text="自定义 nuTilda (m²/s)").grid(row=r, column=0,
                                                          sticky="w", pady=4)
        self.ent_nuTilda = ttk.Entry(inp, textvariable=self.var_nuTilda_custom)
        self.ent_nuTilda.grid(row=r, column=1, columnspan=2, sticky="ew", pady=4)
        r += 1
        return inp

    def _row_entry(self, parent, row, label, var, unit=""):
        """通用的「标签 + 输入框(+单位)」行。"""
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=3)
        ent = ttk.Entry(parent, textvariable=var)
        ent.grid(row=row, column=1, sticky="ew", pady=3)
        if unit:
            ttk.Label(parent, text=unit).grid(row=row, column=2, sticky="e",
                                              pady=3, padx=(4, 0))
        return ent

    def _build_result_panel(self, parent):
        res = ttk.LabelFrame(parent, text="计算结果", padding=12)
        res.columnconfigure(1, weight=1)
        rows = [
            ("雷诺数 Re", "re"),
            ("湍流强度 I", "i"),
            ("湍动能 k (m²/s²)", "k"),
            ("混合长度 L (m)", "l"),
            ("耗散率 ε (m²/s³)", "eps"),
            ("耗散频率 ω (1/s)", "omega"),
            ("nuTilda (m²/s)", "nuTilda"),
            ("阻尼函数 fv1", "fv1"),
            ("湍流粘度 nut (m²/s)", "nut"),
            ("湍流粘度比 nutRatio", "nutRatio"),
            ("当前模型入口参数", "note"),
        ]
        for row, (label, key) in enumerate(rows):
            ttk.Label(res, text=label).grid(row=row, column=0, sticky="w", pady=3)
            var = tk.StringVar(value="—")
            self.results[key] = var
            ttk.Label(res, textvariable=var, style="Value.TLabel").grid(row=row, column=1,
                                                                        sticky="e", pady=3)
        return res

    def _build_button_bar(self, parent):
        bar = ttk.Frame(parent)
        ttk.Button(bar, text="计 算", style="Accent.TButton",
                   command=self.compute).pack(side="left")
        ttk.Button(bar, text="加载示例", command=self.load_example).pack(side="left", padx=(10, 0))
        ttk.Button(bar, text="清空", command=self.clear).pack(side="left", padx=(10, 0))
        return bar

    def _build_verify_panel(self, parent):
        ver = ttk.LabelFrame(parent, text="计算过程与交叉验证", padding=8)
        ver.rowconfigure(0, weight=1)
        ver.columnconfigure(0, weight=1)
        self.verify = tk.Text(ver, height=9, wrap="word", font=("Consolas", 9),
                              bg="#ffffff", fg=TEXT, insertbackground=TEXT,
                              relief="solid", borderwidth=1, padx=6, pady=6,
                              highlightthickness=1, highlightbackground=BORDER,
                              highlightcolor=C_PRIMARY, state="disabled")
        self.verify.grid(row=0, column=0, sticky="nsew")
        return ver

    def _toggle_states(self):
        """根据选项联动启用/禁用相关输入框。"""
        self.ent_I.config(state="normal" if self.var_I_mode.get() == "given" else "disabled")
        self.ent_Cv1.config(state="normal" if self.var_model.get() == "sa" else "disabled")
        self.ent_nuTilda.config(state="normal" if self.var_nuT_method.get() == "custom" else "disabled")
        if "note" in self.results:
            self.results["note"].set(self.MODEL_NOTES[self.var_model.get()])

    # ------------------------------------------------------------- 计算
    def compute(self):
        # 基础输入
        try:
            U = float(self.var_U.get().strip())
            Lc = float(self.var_Lc.get().strip())
            nu = float(self.var_nu.get().strip())
            Cmu = float(self.var_Cmu.get().strip())
            Cv1 = float(self.var_Cv1.get().strip())
        except ValueError:
            messagebox.showerror("输入错误", "请检查数值输入：|U|、Lc、ν、Cmu、Cv1 必须为数字。")
            return

        if U <= 0 or Lc <= 0 or nu <= 0:
            messagebox.showerror("输入错误", "|U|、Lc、ν 必须大于 0。")
            return
        if not (0 < Cmu <= 1):
            messagebox.showerror("输入错误", "Cmu 应为 (0, 1] 之间的值，通常取 0.09。")
            return

        # 雷诺数（可输入，留空则由 |U|·Lc/ν 自动计算）
        re_txt = self.var_Re.get().strip()
        Re_auto = U * Lc / nu
        try:
            Re = float(re_txt) if re_txt else Re_auto
        except ValueError:
            messagebox.showerror("输入错误", "雷诺数 Re 必须为数字（留空则自动计算）。")
            return
        if Re <= 0:
            messagebox.showerror("输入错误", "雷诺数 Re 必须大于 0。")
            return

        # 湍流强度 I
        use_re_I = self.var_I_mode.get() == "re"
        if use_re_I:
            I = 0.16 * Re ** (-1 / 8)
        else:
            try:
                I = float(self.var_I.get().strip())
            except ValueError:
                messagebox.showerror("输入错误", "湍流强度 I 必须为数字（小数形式，如 0.05）。")
                return
            if not (0 < I < 1):
                messagebox.showerror("输入错误", "湍流强度 I 应在 (0, 1) 之间（如 0.05 表示 5%）。")
                return

        # ---- 计算步骤 ----
        k = 1.5 * (U * I) ** 2                       # 步骤1
        L = 0.07 * Lc                                # 步骤2
        eps = Cmu ** 0.75 * k ** 1.5 / L             # 步骤3
        omega1 = math.sqrt(k) / (Cmu ** 0.25 * L)    # 步骤4 形式1
        omega2 = eps / (Cmu * k)                     # 步骤4 形式2（交叉验证）

        # ---- 步骤5：nuTilda ----
        method = self.var_nuT_method.get()
        if method == "exact":
            nuTilda = Cmu * k ** 2 / eps
            nuT_note = "精确：nuTilda ≈ nut = Cmu·k²/ε"
        elif method == "3nu":
            nuTilda = 3 * nu
            nuT_note = "工程经验：3ν（内部流动，最常用）"
        elif method == "5nu":
            nuTilda = 5 * nu
            nuT_note = "工程经验：5ν（外部流动）"
        elif method == "10nu":
            nuTilda = 10 * nu
            nuT_note = "工程经验：10ν（高湍流）"
        elif method == "re":
            nuTilda = nu * math.sqrt(0.01 * Re)
            nuT_note = "基于雷诺数：ν·√(0.01·Re)"
        else:  # custom
            try:
                nuTilda = float(self.var_nuTilda_custom.get().strip())
            except ValueError:
                messagebox.showerror("输入错误", "自定义 nuTilda 必须为数字。")
                return
            if nuTilda < 0:
                messagebox.showerror("输入错误", "nuTilda 不能为负。")
                return
            nuT_note = "自定义输入"

        # ---- 步骤6：湍流粘度 nut（按模型）----
        model = self.var_model.get()
        chi = nuTilda / nu
        fv1 = chi ** 3 / (chi ** 3 + Cv1 ** 3)
        nut_keps = Cmu * k ** 2 / eps
        nut_kw = k / omega1
        if model == "keps":
            nut = nut_keps
        elif model == "komega":
            nut = nut_kw
        else:  # sa
            nut = nuTilda * fv1
        nutRatio = nut / nu                          # 步骤7

        # ---- 结果展示 ----
        self.results["re"].set(fmt(Re) + ("" if re_txt else "（自动）"))
        self.results["i"].set(f"{I:.5g}  ({I * 100:.4g}%)")
        self.results["k"].set(fmt(k) + " m²/s²")
        self.results["l"].set(fmt(L) + " m")
        self.results["eps"].set(fmt(eps) + " m²/s³")
        self.results["omega"].set(fmt(omega1) + " 1/s")
        self.results["nuTilda"].set(fmt(nuTilda) + " m²/s")
        self.results["fv1"].set(fmt(fv1))
        self.results["nut"].set(fmt(nut) + " m²/s")
        self.results["nutRatio"].set(fmt(nutRatio))
        self.results["note"].set(self.MODEL_NOTES[model])

        # ---- 计算过程与交叉验证 ----
        lines = [
            f"① 雷诺数：Re = |U|·Lc/ν = {U:g}×{Lc:g}/{nu:g} = {fmt(Re_auto)}"
            + (f"（使用自定义 Re = {fmt(Re)}）" if re_txt else "（自动）"),
            f"② 湍流强度：I = {I:.5g}"
            + ("（由 0.16·Re^(-1/8) 估算）" if use_re_I else "（直接给定）"),
            f"③ 湍动能：k = 1.5×(|U|·I)² = 1.5×({U:g}×{I:.4g})² = {fmt(k)} m²/s²",
            f"④ 混合长度：L = 0.07·Lc = 0.07×{Lc:g} = {fmt(L)} m",
            f"⑤ 耗散率：ε = Cmu^0.75·k^1.5/L = {fmt(eps)} m²/s³",
        ]
        rel_om = abs(omega1 - omega2) / omega1 * 100 if omega1 else 0
        lines.append(
            f"⑥ 耗散频率：ω = √k/(Cmu^0.25·L) = {fmt(omega1)} 1/s；"
            f"交叉验证 ω' = ε/(Cmu·k) = {fmt(omega2)} 1/s，偏差 {rel_om:.4g}%")
        lines.append(f"⑦ nuTilda（{nuT_note}）= {fmt(nuTilda)} m²/s；χ = nuTilda/ν = {chi:.5g}")
        if model == "sa":
            lines.append(
                f"⑧ fv1 = χ³/(χ³+Cv1³) = {fmt(fv1)}；nut(SA) = nuTilda·fv1 = {fmt(nut)} m²/s")
            if chi < 10:
                lines.append("   提示：χ 较小时 fv1 明显小于 1，nut 会远小于 nuTilda，属正常现象。")
        else:
            lines.append(
                f"⑧ 交叉验证：nut(k-ε) = Cmu·k²/ε = {fmt(nut_keps)} m²/s；"
                f"nut(k-ω) = k/ω = {fmt(nut_kw)} m²/s")
            lines.append(f"   当前模型 nut = {fmt(nut)} m²/s（{self.MODEL_NOTES[model]}）")
        lines.append(
            f"⑨ 湍流粘度比：nutRatio = nut/ν = {fmt(nutRatio)}（入口建议通常 1~10，此处为初始估计值）")
        self._set_verify(lines)

    def _set_verify(self, lines):
        self.verify.config(state="normal")
        self.verify.delete("1.0", tk.END)
        self.verify.insert("1.0", "\n".join(lines))
        self.verify.config(state="disabled")

    # ------------------------------------------------------------- 辅助
    def load_example(self):
        """载入《湍流参数计算流程.md》第 5 节的数值算例。"""
        self.var_U.set("10")
        self.var_Lc.set("0.5")
        self.var_nu.set("1e-5")
        self.var_I.set("0.05")
        self.var_I_mode.set("given")
        self.var_Re.set("")
        self.var_Cmu.set("0.09")
        self.var_model.set("keps")
        self.var_Cv1.set("7.1")
        self.var_nuT_method.set("exact")
        self._toggle_states()
        self.compute()

    def clear(self):
        for v in (self.var_U, self.var_Lc, self.var_nu, self.var_I, self.var_Re,
                  self.var_Cmu, self.var_Cv1, self.var_nuTilda_custom):
            v.set("")
        for key, var in self.results.items():
            var.set("—")
        self.verify.config(state="normal")
        self.verify.delete("1.0", tk.END)
        self.verify.config(state="disabled")


if __name__ == "__main__":
    app = TurbulenceCalcApp()
    app.mainloop()
