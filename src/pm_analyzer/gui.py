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
        self.idle_equivalent = tk.DoubleVar(value=30)
        self._build()

    def _build(self) -> None:
        self.root.title("PM Analyzer - تحليل الصيانة الوقائية")
        self.root.geometry("900x700")
        self.root.minsize(820, 650)
        self.root.configure(background="#F3F6FA")
        style = ttk.Style(self.root)
        style.configure("TFrame", background="#F3F6FA")
        style.configure("Card.TFrame", background="#FFFFFF", relief="solid", borderwidth=1)
        style.configure("Title.TLabel", background="#1F4E78", foreground="#FFFFFF", font=("Arial", 22, "bold"), padding=18)
        style.configure("Section.TLabel", background="#FFFFFF", foreground="#1F4E78", font=("Arial", 12, "bold"))
        style.configure("Card.TLabel", background="#FFFFFF", foreground="#44546A", font=("Arial", 10))
        style.configure("Primary.TButton", font=("Arial", 12, "bold"), padding=10)
        ttk.Label(self, text="PM Analyzer  |  تحليل الصيانة الوقائية", style="Title.TLabel", anchor="center").pack(fill="x", pady=(0, 16))
        files_card = ttk.Frame(self, style="Card.TFrame", padding=18)
        files_card.pack(fill="x")
        ttk.Label(files_card, text="مصادر البيانات", style="Section.TLabel").pack(anchor="e", pady=(0, 10))
        ttk.Label(files_card, text="① ملفات GPS — اختر من ملف واحد إلى 7 ملفات", style="Card.TLabel").pack(anchor="e")
        ttk.Button(files_card, text="استدعاء ملفات GPS", command=self._select_gps).pack(fill="x", pady=5)
        self.gps_label = ttk.Label(
            files_card, text="لم يتم اختيار ملفات", anchor="e", style="Card.TLabel"
        )
        self.gps_label.pack(fill="x", pady=(0, 8))
        ttk.Label(files_card, text="② ملف أوامر الصيانة — المصدر الأساسي للمعدات", style="Card.TLabel").pack(anchor="e")
        ttk.Button(files_card, text="اختيار ملف أوامر الصيانة", command=self._select_maintenance).pack(fill="x", pady=5)
        self.maintenance_label = ttk.Label(files_card, text="لم يتم اختيار ملف", anchor="e", style="Card.TLabel")
        self.maintenance_label.pack(fill="x", pady=(0, 8))
        ttk.Label(files_card, text="③ ملف صرف المواد", style="Card.TLabel").pack(anchor="e")
        ttk.Button(files_card, text="اختيار ملف صرف المواد", command=self._select_materials).pack(fill="x", pady=5)
        self.materials_label = ttk.Label(files_card, text="لم يتم اختيار ملف", anchor="e", style="Card.TLabel")
        self.materials_label.pack(fill="x")

        policy_card = ttk.Frame(self, style="Card.TFrame", padding=18)
        policy_card.pack(fill="x", pady=14)
        ttk.Label(policy_card, text="إعدادات دورة الصيانة", style="Section.TLabel").grid(row=0, column=0, columnspan=6, sticky="e", pady=(0, 12))
        ttk.Label(policy_card, text="كل ساعة تشغيل ساكن تعادل", style="Card.TLabel").grid(row=1, column=5, padx=5, sticky="e")
        ttk.Entry(policy_card, textvariable=self.idle_equivalent, width=10, justify="center").grid(row=1, column=4)
        ttk.Label(policy_card, text="كم", style="Card.TLabel").grid(row=1, column=3, padx=(3, 25))
        ttk.Label(policy_card, text="الصيانة الوقائية كل", style="Card.TLabel").grid(row=1, column=2, padx=5, sticky="e")
        ttk.Entry(policy_card, textvariable=self.interval, width=12, justify="center").grid(row=1, column=1)
        ttk.Label(policy_card, text="كم", style="Card.TLabel").grid(row=1, column=0, padx=3)
        ttk.Label(policy_card, text="تنبيه الصيانة القريبة عند", style="Card.TLabel").grid(row=2, column=5, padx=5, pady=(12, 0), sticky="e")
        ttk.Entry(policy_card, textvariable=self.threshold, width=10, justify="center").grid(row=2, column=4, pady=(12, 0))
        ttk.Label(policy_card, text="% من الدورة", style="Card.TLabel").grid(row=2, column=3, pady=(12, 0), sticky="w")
        policy_card.columnconfigure(5, weight=1)
        ttk.Button(self, text="إنشاء وتصدير تقرير Excel", command=self._run, style="Primary.TButton").pack(fill="x", ipady=7)
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
        try:
            interval = self.interval.get()
            threshold = self.threshold.get()
            idle_equivalent = self.idle_equivalent.get()
        except tk.TclError:
            messagebox.showerror("إعدادات غير صالحة", "أدخل أرقاماً صحيحة في إعدادات دورة الصيانة")
            return
        if interval <= 0 or idle_equivalent < 0 or not 0 <= threshold <= 100:
            messagebox.showerror("إعدادات غير صالحة", "راجع فترة الصيانة ومعامل الساكن ونسبة التنبيه")
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
                interval_km=interval,
                due_soon_percent=threshold,
                idle_hour_equivalent_km=idle_equivalent,
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
