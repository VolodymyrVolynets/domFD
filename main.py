import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from SettingsManager import SettingsManager
from EmployeeManager import EmployeeManager
import traceback


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Employee File Manager")
        self.geometry("650x750")
        self.resizable(False, False)

        self.settings = SettingsManager()
        self.selected_excel = None
        self.upload_status = {}
        self.entries = {}
        self.upload_buttons = []

        self._build_ui()
        self._set_ui_state(enabled=False)

    # ---------- UI BUILD ----------
    def _build_ui(self):
        # Excel selector
        frame_excel = ttk.LabelFrame(self, text="Excel File")
        frame_excel.pack(padx=10, pady=10, fill="x")

        ttk.Button(frame_excel, text="Select Excel File", command=self.select_excel).pack(pady=5)
        self.excel_label = ttk.Label(frame_excel, text="No file selected", foreground="gray")
        self.excel_label.pack(pady=2)

        # Settings
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

        # Upload buttons
        frame_uploads = ttk.LabelFrame(self, text="Document Uploads")
        frame_uploads.pack(padx=10, pady=10, fill="x")

        # Todo: on every upload must convert uploaded files to a singe pdf document and place it to a generated folder with specific filename. Tax = "{last_name} Tax {tax_expiry} {car_reg}" NCT = "{last_name} NCT {nct_expiry} {car_reg}" Insurance="{last_name} Insurance {insurance_expiry} {car_reg}" License="{last_name} License {license_expiry}", passport="{last_name} Passport {passport_expiry}", irp="{last_name} Irp {irp_expiry} {irp_type}" pp={last_name} PP {pp_expiry}, GDPR="{last_name} GPDR" spf= "{last_name} SPF", obu= "{last_name} OBU {car_reg}"
        uploads = [
            "Tax", "NCT", "Insurance", "License",
            "Passport", "IRP", "Penalty Points",
            "GDPR", "SPF", "OBU"
        ]

        for name in uploads:
            self.create_upload_button(frame_uploads, name)

        # Generate PDF
        ttk.Button(self, text="Generate PDF", command=self.generate_pdf).pack(pady=15)

    # ---------- EVENT HANDLERS ----------
    def select_excel(self):
        #Todo: once excel file uploaded and valid(all fields are filled), must create a folder "{first_name} {last_name}"
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        if self.validate_excel(file_path):
            self.selected_excel = file_path
            self.excel_label.config(text=os.path.basename(file_path), foreground="green")
            self._set_ui_state(enabled=True)
        else:
            self.excel_label.config(text="Invalid Excel file", foreground="red")
            self._set_ui_state(enabled=False)
            messagebox.showwarning("Invalid Excel", "Please select a valid Excel file with all fields filled.")

    def validate_excel(self, file_path) -> bool:
        """Use EmployeeManager to validate and load employee."""
        try:
            manager = EmployeeManager(file_path, "Sheet2")
            employee = manager.load_employee()
            print(f"[VALID] Employee loaded: {employee}")
            return True
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Excel Validation Failed", str(e))
            return False

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
        """Handle document upload."""
        file_path = filedialog.askopenfilename(
            title=f"Upload {doc_type} Document",
            filetypes=[("PDF or Image", "*.pdf *.jpg *.jpeg *.png")]
        )
        if file_path:
            print(f"[UPLOAD] {doc_type}: {file_path}")
            self.upload_status[doc_type].config(text="✔", foreground="green")

    def generate_pdf(self):
        """Placeholder for PDF generation."""
        print("[ACTION] Generate PDF clicked")

        # TODO: once button is clicked should fill pdf file(templates/template.pdf) with data: store_manager_name(settings), employee_name(employee object first name + last name), date(editable date from ui), penalty_points(employee), time(random time between 12:00 to 16:00), franchise_name(settings), shop_name(settings), employee_age(employee object age in years, but just a number), car_make, car_model, car_reg, next_date(date from ui+ 1 day), last_name(employee object), first_name(employee object), title(employee object), address(employee object), date_of_birth, phone_number. generated file must be placed in folder "{first_name} {last_name}" and have name 'to print.pdf'


    # ---------- UI CONTROL ----------
    def _set_ui_state(self, enabled: bool):
        """Enable or disable settings and upload buttons."""
        state = "normal" if enabled else "disabled"
        for entry in self.entries.values():
            entry.config(state=state)
        for btn in self.upload_buttons:
            btn.config(state=state)
        print(f"[UI] {'Enabled' if enabled else 'Disabled'} all controls.")


# ---------- MAIN ----------
if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
