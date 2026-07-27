"""Arabic desktop interface for selecting source files and exporting the report."""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from pm_analyzer.config import UserPreferences
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
        try:
            preferences = UserPreferences.load()
        except (OSError, ValueError, KeyError):
            preferences = UserPreferences()
        self.interval = tk.StringVar(value=str(preferences.interval_km or ""))
        self.threshold = tk.StringVar(value=str(preferences.due_soon_percent or ""))
        self.idle_equivalent = tk.StringVar(value=str(preferences.idle_equivalent_km or ""))
        self.policy_summary = tk.StringVar()
        self.result_summary = tk.StringVar(value="لم يتم تنفيذ تحليل بعد")
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
        style.configure("Success.TLabel", background="#E2F0D9", foreground="#375623", font=("Arial", 11, "bold"), padding=12)
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

        policy_card = ttk.Frame(self, style="Card.TFrame", padding=14)
        policy_card.pack(fill="x", pady=12)
        ttk.Label(policy_card, text="⚙ سياسة الصيانة", style="Section.TLabel").pack(side="right")
        ttk.Button(policy_card, text="فتح الإعدادات", command=self._open_settings).pack(side="left")
        ttk.Label(policy_card, textvariable=self.policy_summary, style="Card.TLabel").pack(side="right", padx=20)
        self._update_policy_summary()
        self.run_button = ttk.Button(self, text="▶ إنشاء وتصدير تقرير Excel", command=self._run, style="Primary.TButton")
        self.run_button.pack(fill="x", ipady=7)
        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", pady=(10, 3))
        ttk.Label(self, textvariable=self.status, anchor="center", foreground="#1F4E78").pack(fill="x", pady=12)
        ttk.Label(self, textvariable=self.result_summary, anchor="center", style="Success.TLabel").pack(fill="x")
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

    def _update_policy_summary(self) -> None:
        if self.interval.get() and self.idle_equivalent.get() and self.threshold.get():
            self.policy_summary.set(
                f"الدورة: {self.interval.get()} كم  |  الساكن: {self.idle_equivalent.get()} كم/ساعة  |  قريب: {self.threshold.get()}%"
            )
        else:
            self.policy_summary.set("الإعدادات مطلوبة قبل أول تشغيل")

    def _open_settings(self) -> None:
        window = tk.Toplevel(self.root)
        window.title("إعدادات سياسة الصيانة")
        window.geometry("520x340")
        window.resizable(False, False)
        window.transient(self.root)
        window.grab_set()
        card = ttk.Frame(window, padding=25)
        card.pack(fill="both", expand=True)
        ttk.Label(card, text="⚙ إعدادات سياسة الصيانة", font=("Arial", 17, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 22))
        fields = (
            ("فترة الصيانة الوقائية (KM)", self.interval),
            ("الكيلومترات المكافئة لساعة Idle", self.idle_equivalent),
            ("نسبة Due Soon (%)", self.threshold),
        )
        for row, (label, variable) in enumerate(fields, 1):
            ttk.Label(card, text=label).grid(row=row, column=1, sticky="e", padx=12, pady=8)
            ttk.Entry(card, textvariable=variable, justify="center", width=18).grid(row=row, column=0, pady=8)
        ttk.Label(card, text="هذه القيم تخص سياسة شركتك ولا يفرض البرنامج قيماً افتراضية.", foreground="#666666").grid(row=4, column=0, columnspan=2, pady=12)
        ttk.Button(card, text="حفظ الإعدادات", command=lambda: self._save_settings(window), style="Primary.TButton").grid(row=5, column=0, columnspan=2, sticky="ew")
        card.columnconfigure(0, weight=1)
        card.columnconfigure(1, weight=1)

    def _preferences(self) -> UserPreferences:
        try:
            preferences = UserPreferences(
                interval_km=int(self.interval.get()),
                idle_equivalent_km=float(self.idle_equivalent.get()),
                due_soon_percent=int(self.threshold.get()),
            )
        except ValueError as error:
            raise ValueError("أدخل أرقاماً صحيحة في إعدادات سياسة الصيانة") from error
        return preferences.validated()

    def _save_settings(self, window: tk.Toplevel) -> None:
        try:
            self._preferences().save()
        except (OSError, ValueError) as error:
            messagebox.showerror("إعدادات غير صالحة", str(error), parent=window)
            return
        self._update_policy_summary()
        window.destroy()
        self.status.set("تم حفظ إعدادات سياسة الصيانة")

    def _run(self) -> None:
        if not self.gps_paths or self.maintenance_path is None or self.materials_path is None:
            messagebox.showwarning("ملفات ناقصة", "اختر ملفات GPS وملف الأوامر وملف صرف المواد أولاً")
            return
        try:
            preferences = self._preferences()
        except ValueError as error:
            messagebox.showerror("الإعدادات مطلوبة", str(error))
            self._open_settings()
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
        self.status.set("جاري قراءة الملفات وتحليل أوامر الصيانة والمواد...")
        self.result_summary.set("التحليل قيد التنفيذ")
        self.progress.start(12)
        self.run_button.state(["disabled"])
        threading.Thread(
            target=self._run_worker,
            args=(Path(output), preferences),
            daemon=True,
        ).start()

    def _run_worker(self, output: Path, preferences: UserPreferences) -> None:
        assert self.maintenance_path is not None
        assert self.materials_path is not None
        assert preferences.interval_km is not None
        assert preferences.due_soon_percent is not None
        assert preferences.idle_equivalent_km is not None
        try:
            result = analyze(
                self.maintenance_path,
                self.materials_path,
                self.gps_paths,
                interval_km=preferences.interval_km,
                due_soon_percent=preferences.due_soon_percent,
                idle_hour_equivalent_km=preferences.idle_equivalent_km,
            )
            export_report(result, output)
        except (OSError, ValueError) as error:
            self.root.after(0, self._analysis_failed, str(error))
            return
        self.root.after(0, self._analysis_complete, result, output)

    def _analysis_failed(self, error: str) -> None:
        self.progress.stop()
        self.run_button.state(["!disabled"])
        self.status.set("تعذر إنشاء التقرير")
        self.result_summary.set("لم يكتمل التحليل")
        messagebox.showerror("تعذر إنشاء التقرير", error)

    def _analysis_complete(self, result: AnalysisResult, output: Path) -> None:
        self.result = result
        self.progress.stop()
        self.run_button.state(["!disabled"])
        counts: dict[str, int] = {}
        for row in result.analysis:
            status = str(row["status"])
            counts[status] = counts.get(status, 0) + 1
        pm_orders = sum(
            row["classification"] == "PM" for row in result.order_classifications
        )
        self.result_summary.set(
            f"المعدات: {len(result.analysis)}  |  أوامر PM: {pm_orders}  |  مستحق: {counts.get('Due', 0)}  |  قريب: {counts.get('Due Soon', 0)}  |  بدون GPS: {counts.get('No GPS Data', 0)}"
        )
        self.status.set(f"تم إنشاء التقرير بنجاح: {output}")
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
