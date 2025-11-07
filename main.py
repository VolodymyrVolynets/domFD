import os
import random
import traceback
from datetime import date, datetime, timedelta
from typing import Dict

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from EmployeeManager import Employee, EmployeeManager
from PDFManipulator import PDFManipulator
from SettingsManager import SettingsManager


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Employee File Manager")
        self.geometry("650x750")
        self.resizable(False, False)

        self.settings = SettingsManager()
        self.entry_vars: Dict[str, tk.StringVar] = {}

        self.selected_excel: str | None = None
        self.employee: Employee | None = None
        self.employee_folder: str | None = None
        self.pdf_manipulator: PDFManipulator | None = None

        self.upload_status: Dict[str, ttk.Label] = {}
        self.uploaded_files: Dict[str, str] = {}
        self.entries: Dict[str, ttk.Entry] = {}
        self.upload_buttons: list[ttk.Button] = []

        self._build_ui()
        self._set_ui_state(enabled=False)

    # ---------- UI BUILD ----------
    def _build_ui(self):
        frame_excel = ttk.LabelFrame(self, text="Excel File")
        frame_excel.pack(padx=10, pady=10, fill="x")

        ttk.Button(frame_excel, text="Select Excel File", command=self.select_excel).pack(pady=5)
        self.excel_label = ttk.Label(frame_excel, text="No file selected", foreground="gray")
        self.excel_label.pack(pady=2)

        frame_settings = ttk.LabelFrame(self, text="Settings")
        frame_settings.pack(padx=10, pady=10, fill="x")

        for field in ["franchise_name", "shop_name", "store_manager_name", "date"]:
            row = ttk.Frame(frame_settings)
            row.pack(fill="x", pady=3)
            ttk.Label(row, text=field.replace("_", " ").title() + ":", width=20, anchor="w").pack(side="left")
            var = tk.StringVar(value=self.settings.get(field))
            entry = ttk.Entry(row, textvariable=var)
            entry.pack(side="left", expand=True, fill="x")
            var.trace_add("write", lambda *_ , f=field, v=var: self.on_setting_change(f, v.get()))
            self.entries[field] = entry
            self.entry_vars[field] = var

        frame_uploads = ttk.LabelFrame(self, text="Document Uploads")
        frame_uploads.pack(padx=10, pady=10, fill="x")

        uploads = [
            "Tax", "NCT", "Insurance", "License",
            "Passport", "IRP", "Penalty Points",
            "GDPR", "SPF", "OBU"
        ]

        for name in uploads:
            self.create_upload_button(frame_uploads, name)

        ttk.Button(self, text="Generate PDF", command=self.generate_pdf).pack(pady=15)

    # ---------- EVENT HANDLERS ----------
    def select_excel(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")]
        )
        if not file_path:
            return

        employee = self.validate_excel(file_path)
        if employee:
            self.selected_excel = file_path
            self.excel_label.config(text=os.path.basename(file_path), foreground="green")
            self._set_ui_state(enabled=True)
        else:
            self.excel_label.config(text="Invalid Excel file", foreground="red")
            self._set_ui_state(enabled=False)
            messagebox.showwarning("Invalid Excel", "Please select a valid Excel file with all fields filled.")

    def validate_excel(self, file_path) -> Employee | None:
        try:
            manager = EmployeeManager(file_path, "Sheet2")
            employee = manager.load_employee()
            print(f"[VALID] Employee loaded: {employee}")
            self.employee = employee
            self.employee_folder = self._ensure_employee_folder(employee, file_path)
            self.pdf_manipulator = PDFManipulator(self.employee_folder)
            self.uploaded_files.clear()
            for status in self.upload_status.values():
                status.config(text="⏺", foreground="gray")
            return employee
        except Exception as exc:  # noqa: BLE001
            traceback.print_exc()
            messagebox.showerror("Excel Validation Failed", str(exc))
            self.employee = None
            self.employee_folder = None
            self.pdf_manipulator = None
            return None

    def on_setting_change(self, key, value):
        print(f"[SETTING] Changed {key}: {value}")
        self.settings.set(key, value)

    def create_upload_button(self, parent, label_text):
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=3, padx=5)

        label = ttk.Label(frame, text=label_text, width=20, anchor="w")
        label.pack(side="left")

        btn = ttk.Button(frame, text=f"Upload {label_text}", command=lambda l=label_text: self.upload_file(l))
        btn.pack(side="left", padx=5)

        status_label = ttk.Label(frame, text="⏺", foreground="gray")
        status_label.pack(side="left")

        self.upload_status[label_text] = status_label
        self.upload_buttons.append(btn)

    def upload_file(self, doc_type):
        if not self.employee or not self.pdf_manipulator:
            messagebox.showwarning("No Employee", "Please load a valid employee first.")
            return

        file_path = filedialog.askopenfilename(
            title=f"Upload {doc_type} Document",
            filetypes=[("PDF or Image", "*.pdf *.jpg *.jpeg *.png")]
        )
        if not file_path:
            return

        try:
            output_name = self._build_document_name(doc_type)
        except ValueError as exc:
            messagebox.showerror("Upload Error", str(exc))
            return

        try:
            stored_path = self.pdf_manipulator.convert_to_pdf(file_path, output_name)
        except Exception as exc:  # noqa: BLE001
            traceback.print_exc()
            messagebox.showerror("Conversion Failed", str(exc))
            return

        print(f"[UPLOAD] {doc_type}: {stored_path}")
        self.upload_status[doc_type].config(text="✔", foreground="green")
        self.uploaded_files[doc_type] = stored_path

    def generate_pdf(self):
        print("[ACTION] Generate PDF clicked")
        if not self.employee or not self.pdf_manipulator:
            messagebox.showwarning("No Employee", "Please load a valid employee first.")
            return

        try:
            ui_date = self._get_ui_date()
        except ValueError as exc:
            messagebox.showerror("Invalid Date", str(exc))
            return

        template_path = os.path.join(os.path.dirname(__file__), "templates", "template.pdf")
        if not os.path.exists(template_path):
            messagebox.showerror(
                "Template Missing",
                f"Template PDF not found at {template_path}.",
            )
            return

        employee = self.employee
        data = {
            "store_manager_name": self.entry_vars["store_manager_name"].get(),
            "employee_name": f"{employee.first_name} {employee.last_name}",
            "date": ui_date.strftime("%d/%m/%Y"),
            "penalty_points": str(employee.penalty_points),
            "time": self._random_time(),
            "franchise_name": self.entry_vars["franchise_name"].get(),
            "shop_name": self.entry_vars["shop_name"].get(),
            "employee_age": str(employee.calculate_age()),
            "car_make": employee.car_make,
            "car_model": employee.car_model,
            "car_reg": employee.car_reg,
            "next_date": (ui_date + timedelta(days=1)).strftime("%d/%m/%Y"),
            "last_name": employee.last_name,
            "first_name": employee.first_name,
            "title": employee.get_title(),
            "address": employee.address,
            "date_of_birth": self._format_date(employee.date_of_birth),
            "phone_number": employee.phone_number,
        }

        output_path = self.pdf_manipulator.fill_pdf_form(
            template_path,
            data,
            output_name="to print.pdf",
        )
        messagebox.showinfo("PDF Generated", f"Generated file saved to {output_path}")

    # ---------- UI CONTROL ----------
    def _set_ui_state(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for entry in self.entries.values():
            entry.config(state=state)
        for btn in self.upload_buttons:
            btn.config(state=state)
        print(f"[UI] {'Enabled' if enabled else 'Disabled'} all controls.")

    # ---------- HELPERS ----------
    def _ensure_employee_folder(self, employee: Employee, file_path: str) -> str:
        base_dir = os.path.dirname(file_path) or os.getcwd()
        folder_name = f"{employee.first_name} {employee.last_name}".strip()
        target = os.path.join(base_dir, folder_name)
        os.makedirs(target, exist_ok=True)
        return target

    def _get_ui_date(self) -> date:
        date_str = self.entry_vars["date"].get()
        try:
            return datetime.strptime(date_str, "%d/%m/%Y").date()
        except ValueError as exc:
            raise ValueError("Date must be in DD/MM/YYYY format.") from exc

    def _random_time(self) -> str:
        hour = random.randint(12, 16)
        minute = 0 if hour == 16 else random.randint(0, 59)
        return f"{hour:02d}:{minute:02d}"

    def _format_date(self, value: date | None) -> str:
        if not value:
            return ""
        if isinstance(value, datetime):
            value = value.date()
        return value.strftime("%d/%m/%Y") if isinstance(value, date) else str(value)

    def _build_document_name(self, doc_type: str) -> str:
        if not self.employee:
            raise ValueError("No employee data available.")

        employee = self.employee
        ui_date = self._get_ui_date()

        def fmt(value):
            if isinstance(value, datetime):
                value = value.date()
            if isinstance(value, date):
                return value.strftime("%d-%m-%Y")
            return str(value) if value else "Unknown"

        last = employee.last_name
        car_reg = employee.car_reg

        if doc_type == "Tax":
            return f"{last} Tax {fmt(employee.tax_expiry)} {car_reg}"
        if doc_type == "NCT":
            return f"{last} NCT {fmt(employee.nct_expiry)} {car_reg}"
        if doc_type == "Insurance":
            return f"{last} Insurance {fmt(employee.insurance_expiry)} {car_reg}"
        if doc_type == "License":
            return f"{last} License {fmt(employee.license_expiry)}"
        if doc_type == "Passport":
            return f"{last} Passport {fmt(employee.passport_expiry)}"
        if doc_type == "IRP":
            expiry = fmt(employee.irp_expiry)
            irp_type = employee.irp_type or "Unknown"
            return f"{last} Irp {expiry} {irp_type}".strip()
        if doc_type == "Penalty Points":
            pp_expiry = employee.calculate_pp_expiry(ui_date)
            return f"{last} PP {fmt(pp_expiry)}"
        if doc_type == "GDPR":
            return f"{last} GPDR"
        if doc_type == "SPF":
            return f"{last} SPF"
        if doc_type == "OBU":
            return f"{last} OBU {car_reg}"

        raise ValueError(f"Unknown document type: {doc_type}")


# ---------- MAIN ----------
if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
