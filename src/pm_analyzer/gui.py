"""Arabic desktop interface for selecting source files and exporting the report."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from pm_analyzer.config import DUE_SOON_PERCENT, PM_INTERVAL_KM
from pm_analyzer.engine import AnalysisResult, analyze, export_report


class AnalyzerApp(ttk.Frame):
    """Small, guided desktop workflow for non-technical users."""

    def __init__(self, master: tk.Tk) -> None:
        super().__init__(master, padding=20)
        self.root = master
        self.gps_paths: list[Path] = []
        self.maintenance_path: Path | None = None
        self.materials_path: Path | None = None
        self.result: AnalysisResult | None = None
        self.status = tk.StringVar(value="اختر الملفات بالترتيب ثم اضغط إنشاء التقرير")
        self.interval = tk.IntVar(value=PM_INTERVAL_KM)
        self.threshold = tk.IntVar(value=DUE_SOON_PERCENT)
        self._build()

    def _build(self) -> None:
        self.root.title("PM Analyzer - تحليل الصيانة الوقائية")
        self.root.geometry("760x560")
        self.root.minsize(700, 520)
        ttk.Label(self, text="تحليل الصيانة الوقائية", font=("Arial", 20, "bold")).pack(pady=(0, 15))
        ttk.Label(self, text="1) ملفات GPS (من ملف واحد إلى 7 ملفات)", font=("Arial", 12, "bold")).pack(anchor="e")
        ttk.Button(self, text="استدعاء ملفات GPS", command=self._select_gps).pack(fill="x", pady=5)
        self.gps_label = ttk.Label(self, text="لم يتم اختيار ملفات", anchor="e")
        self.gps_label.pack(fill="x")
        ttk.Separator(self).pack(fill="x", pady=12)
        ttk.Label(self, text="2) ملف أوامر الصيانة (الملف الأساسي)", font=("Arial", 12, "bold")).pack(anchor="e")
        ttk.Button(self, text="اختيار ملف أوامر الصيانة", command=self._select_maintenance).pack(fill="x", pady=5)
        self.maintenance_label = ttk.Label(self, text="لم يتم اختيار ملف", anchor="e")
        self.maintenance_label.pack(fill="x")
        ttk.Label(self, text="3) ملف صرف المواد", font=("Arial", 12, "bold")).pack(anchor="e", pady=(12, 0))
        ttk.Button(self, text="اختيار ملف صرف المواد", command=self._select_materials).pack(fill="x", pady=5)
        self.materials_label = ttk.Label(self, text="لم يتم اختيار ملف", anchor="e")
        self.materials_label.pack(fill="x")
        options = ttk.Frame(self)
        options.pack(fill="x", pady=15)
        ttk.Label(options, text="فترة الصيانة (كم)").grid(row=0, column=3, padx=5)
        ttk.Entry(options, textvariable=self.interval, width=12).grid(row=0, column=2)
        ttk.Label(options, text="نسبة الصيانة القريبة %").grid(row=0, column=1, padx=5)
        ttk.Entry(options, textvariable=self.threshold, width=8).grid(row=0, column=0)
        ttk.Button(self, text="إنشاء وتصدير التقرير", command=self._run).pack(fill="x", ipady=8)
        ttk.Label(self, textvariable=self.status, anchor="center", foreground="#1F4E78").pack(fill="x", pady=12)
        self.pack(fill="both", expand=True)

    def _select_gps(self) -> None:
        names = filedialog.askopenfilenames(title="اختر من 1 إلى 7 ملفات GPS", filetypes=[("Excel", "*.xlsx")])
        if not names:
            return
        if len(names) > 7:
            messagebox.showerror("عدد الملفات", "الحد الأقصى 7 ملفات GPS")
            return
        self.gps_paths = [Path(name) for name in names]
        self.gps_label.configure(text=f"تم اختيار {len(names)} ملف: " + "، ".join(Path(name).name for name in names))

    def _select_maintenance(self) -> None:
        name = filedialog.askopenfilename(title="اختر ملف أوامر الصيانة", filetypes=[("Excel", "*.xlsx")])
        if name:
            self.maintenance_path = Path(name)
            self.maintenance_label.configure(text=self.maintenance_path.name)

    def _select_materials(self) -> None:
        name = filedialog.askopenfilename(title="اختر ملف صرف المواد", filetypes=[("Excel", "*.xlsx")])
        if name:
            self.materials_path = Path(name)
            self.materials_label.configure(text=self.materials_path.name)

    def _run(self) -> None:
        if not self.gps_paths or self.maintenance_path is None or self.materials_path is None:
            messagebox.showwarning("ملفات ناقصة", "اختر ملفات GPS وملف الأوامر وملف صرف المواد أولاً")
            return
        output = filedialog.asksaveasfilename(
            title="حفظ تقرير الصيانة الوقائية",
            defaultextension=".xlsx",
            initialfile="تقرير_تحليل_الصيانة_الوقائية.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not output:
            return
        self.status.set("جاري قراءة الملفات والحساب...")
        self.update_idletasks()
        try:
            self.result = analyze(
                self.maintenance_path,
                self.materials_path,
                self.gps_paths,
                interval_km=self.interval.get(),
                due_soon_percent=self.threshold.get(),
            )
            export_report(self.result, Path(output))
        except (OSError, ValueError) as error:
            self.status.set("تعذر إنشاء التقرير")
            messagebox.showerror("خطأ", str(error))
            return
        self.status.set(f"تم إنشاء التقرير: {output}")
        messagebox.showinfo("تم", "تم إنشاء تقرير الصيانة الوقائية بنجاح")


def main() -> None:
    root = tk.Tk()
    AnalyzerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
