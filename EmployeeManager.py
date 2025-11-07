"""Employee domain model and Excel-backed loader."""

from __future__ import annotations

import calendar
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any, Dict, Iterable

import pandas as pd


@dataclass
class Employee:
    """Represents a single employee with relevant data and behaviour."""

    first_name: str
    last_name: str
    phone_number: str
    address: str
    gender: str
    date_of_birth: Any
    car_reg: str
    penalty_points: Any
    car_make: str
    car_model: str
    tax_expiry: Any
    nct_expiry: Any
    insurance_expiry: Any
    passport_expiry: Any
    license_expiry: Any
    irp_type: str | None = None
    irp_expiry: Any | None = None

    def __post_init__(self) -> None:
        self.date_of_birth = self._parse_date(self.date_of_birth)
        self.penalty_points = (
            int(self.penalty_points) if str(self.penalty_points).isdigit() else 0
        )
        self.tax_expiry = self._parse_date(self.tax_expiry)
        self.nct_expiry = self._parse_date(self.nct_expiry)
        self.insurance_expiry = self._parse_date(self.insurance_expiry)
        self.passport_expiry = self._parse_date(self.passport_expiry)
        self.license_expiry = self._parse_date(self.license_expiry)
        self.irp_type = self.irp_type or None
        self.irp_expiry = self._parse_date(self.irp_expiry)

    # ---------------------------------------------------------------------
    @staticmethod
    def _parse_date(value: Any) -> date | None:
        """Parse a cell value into a :class:`datetime.date` when possible."""
        if isinstance(value, date):
            return value
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, str):
            value = value.strip()
            if not value:
                return None
            for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
                try:
                    return datetime.strptime(value, fmt).date()
                except ValueError:
                    continue
        return None

    # ------------------------------------------------------------------
    def calculate_age(self) -> int:
        if not self.date_of_birth:
            return 0
        today = date.today()
        return today.year - self.date_of_birth.year - (
            (today.month, today.day)
            < (self.date_of_birth.month, self.date_of_birth.day)
        )

    def get_title(self) -> str:
        gender = str(self.gender).strip().lower()
        if gender in {"male", "m"}:
            return "Mr"
        if gender in {"female", "f"}:
            return "Ms"
        return "Mx"

    def calculate_pp_expiry(self, reference_date: date) -> date:
        """Calculate the penalty-points expiry using UI ``reference_date``."""
        months_to_add = 3 if self.penalty_points == 0 else 1
        return self._add_months(reference_date, months_to_add)

    @staticmethod
    def _add_months(source_date: date, months: int) -> date:
        if not isinstance(source_date, date):
            raise ValueError("reference_date must be a date instance")

        month_index = source_date.month - 1 + months
        year = source_date.year + month_index // 12
        month = month_index % 12 + 1
        day = min(source_date.day, calendar.monthrange(year, month)[1])
        return date(year, month, day)

    # ------------------------------------------------------------------
    def __repr__(self) -> str:  # pragma: no cover - debug helper
        attrs = {
            "first_name": self.first_name,
            "last_name": self.last_name,
            "phone_number": self.phone_number,
            "address": self.address,
            "gender": self.gender,
            "date_of_birth": self.date_of_birth,
            "car_reg": self.car_reg,
            "penalty_points": self.penalty_points,
            "car_make": self.car_make,
            "car_model": self.car_model,
            "tax_expiry": self.tax_expiry,
            "nct_expiry": self.nct_expiry,
            "insurance_expiry": self.insurance_expiry,
            "passport_expiry": self.passport_expiry,
            "license_expiry": self.license_expiry,
            "irp_type": self.irp_type,
            "irp_expiry": self.irp_expiry,
        }
        formatted = ", ".join(f"{k}={v!r}" for k, v in attrs.items())
        return f"Employee({formatted})"


class EmployeeManager:
    """Reads employee data from an Excel sheet into :class:`Employee`."""

    FIELD_ALIASES: Dict[str, Iterable[str]] = {
        "first_name": ("first_name", "first name"),
        "last_name": ("last_name", "last name"),
        "phone_number": ("phone_number", "phone number", "mobile"),
        "address": ("address",),
        "gender": ("male/female", "gender", "sex"),
        "date_of_birth": ("date_of_birth", "date of birth", "dob"),
        "car_reg": ("car_reg", "car reg", "registration"),
        "penalty_points": (
            "number_of_penalty_points",
            "number of penalty points",
            "penalty points",
            "penalty_points",
        ),
        "car_make": ("car_make", "car make"),
        "car_model": ("car_model", "car model"),
        "tax_expiry": ("tax_expiry", "tax expiry"),
        "nct_expiry": ("nct_expiry", "nct expiry"),
        "insurance_expiry": ("insurance_expiry", "insurance expiry"),
        "passport_expiry": ("passport_expiry", "passport expiry"),
        "license_expiry": ("license_expiry", "license expiry", "licence expiry"),
    }

    OPTIONAL_FIELD_ALIASES: Dict[str, Iterable[str]] = {
        "irp_type": ("irp_type", "irp type"),
        "irp_expiry": ("irp_expiry", "irp expiry"),
    }

    def __init__(self, excel_path: str, sheet_name: str):
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.employee: Employee | None = None

    # ------------------------------------------------------------------
    def load_employee(self) -> Employee:
        df = pd.read_excel(
            self.excel_path,
            sheet_name=self.sheet_name,
            header=None,
            usecols="A:B",
        )
        df = df.dropna(subset=[0])

        raw_data = {str(row[0]).strip().lower(): row[1] for _, row in df.iterrows()}
        normalised = self._normalise_fields(raw_data)

        self._validate_required(normalised)

        self.employee = Employee(**normalised)
        print(f"[✔] Employee loaded from sheet '{self.sheet_name}': {self.employee}")
        return self.employee

    # ------------------------------------------------------------------
    def _normalise_fields(self, raw: Dict[str, Any]) -> Dict[str, Any]:
        data: Dict[str, Any] = {}

        for key, aliases in self.FIELD_ALIASES.items():
            found, value = self._first_present(raw, aliases)
            if not found:
                alias_display = ", ".join(aliases)
                raise ValueError(
                    f"Missing required field '{alias_display}' in sheet '{self.sheet_name}'."
                )
            data[key] = value

        for key, aliases in self.OPTIONAL_FIELD_ALIASES.items():
            found, value = self._first_present(raw, aliases)
            if found:
                data[key] = value
            else:
                data.setdefault(key, None)

        return data

    def _first_present(self, raw: Dict[str, Any], aliases: Iterable[str]) -> tuple[bool, Any]:
        for alias in aliases:
            alias_key = alias.lower()
            if alias_key in raw:
                value = self._clean_value(raw[alias_key])
                return True, value
        return False, None

    @staticmethod
    def _clean_value(value: Any) -> Any:
        if isinstance(value, str):
            value = value.strip()
            return value or None
        if pd.isna(value):
            return None
        return value

    def _validate_required(self, data: Dict[str, Any]) -> None:
        for field in self.FIELD_ALIASES:
            value = data.get(field)
            if self._is_missing(value):
                raise ValueError(
                    f"Field '{field}' is empty in sheet '{self.sheet_name}'."
                )
        print("[VALIDATION] All required fields are filled.")

    @staticmethod
    def _is_missing(value: Any) -> bool:
        if value is None:
            return True
        if isinstance(value, str):
            return not value.strip()
        return pd.isna(value)
