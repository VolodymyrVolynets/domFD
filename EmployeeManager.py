import pandas as pd
from datetime import datetime, date


class Employee:
    """Represents a single employee with relevant data and behavior."""
    #Todo: add tax_expiry, nct_expiry, insurance_expiry, passport_expiry, license_expiry, optional irp_expiry and create function pp_expiry(if 0 return date from ui+ 3 month if more than 0 ui date + 1 month), irp_type(string)
    def __init__(self, first_name, last_name, phone_number, address,
                 gender, date_of_birth, car_reg, penalty_points,
                 car_make, car_model):
        self.first_name = first_name
        self.last_name = last_name
        self.phone_number = phone_number
        self.address = address
        self.gender = gender
        self.date_of_birth = self._parse_date(date_of_birth)
        self.car_reg = car_reg
        self.penalty_points = int(penalty_points) if str(penalty_points).isdigit() else 0
        self.car_make = car_make
        self.car_model = car_model

    def _parse_date(self, value):
        """Parse date in dd/mm/yyyy format, or detect Excel datetime."""
        if isinstance(value, (datetime, date)):
            return value.date() if isinstance(value, datetime) else value
        if isinstance(value, str):
            value = value.strip()
            try:
                return datetime.strptime(value, "%d/%m/%Y").date()
            except ValueError:
                pass
        return None

    def calculate_age(self) -> int:
        """Return employee's age in years."""
        if not self.date_of_birth:
            return 0
        today = date.today()
        return today.year - self.date_of_birth.year - (
            (today.month, today.day) < (self.date_of_birth.month, self.date_of_birth.day)
        )

    def get_title(self) -> str:
        """Return title based on gender."""
        gender = str(self.gender).strip().lower()
        if gender in ("male", "m"):
            return "Mr"
        elif gender in ("female", "f"):
            return "Ms"
        else:
            return "Mx"

    def __repr__(self):
        return (f"{self.__class__.__name__}("
                f"first_name='{self.first_name}', "
                f"last_name='{self.last_name}', "
                f"phone_number='{self.phone_number}', "
                f"address='{self.address}', "
                f"gender='{self.gender}', "
                f"date_of_birth='{self.date_of_birth}', "
                f"car_reg='{self.car_reg}', "
                f"penalty_points={self.penalty_points}, "
                f"car_make='{self.car_make}', "
                f"car_model='{self.car_model}')")


class EmployeeManager:
    #Todo: add more properties from employee class and use all of them to check validity(do not check irp type and irp expiry)
    """Reads employee data from an Excel sheet where:
       - Column A has field names
       - Column B has corresponding values
    """

    def __init__(self, excel_path: str, sheet_name: str):
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.employee: Employee | None = None

    def load_employee(self) -> Employee:
        """Read Excel sheet and create an Employee object."""
        df = pd.read_excel(self.excel_path, sheet_name=self.sheet_name, header=None, usecols="A:B")
        df = df.dropna(subset=[0])  # ignore empty rows

        data_dict = {str(row[0]).strip().lower(): row[1] for _, row in df.iterrows()}

        required_fields = [
            "first_name", "last_name", "phone_number", "address",
            "male/female", "date_of_birth", "car_reg",
            "number_of_penalty_points", "car_make", "car_model"
        ]

        missing = [f for f in required_fields if f not in data_dict]
        if missing:
            raise ValueError(f"Missing fields in sheet '{self.sheet_name}': {', '.join(missing)}")

        self.validate_fields_not_empty(data_dict, required_fields)

        self.employee = Employee(
            first_name=data_dict["first_name"],
            last_name=data_dict["last_name"],
            phone_number=data_dict["phone_number"],
            address=data_dict["address"],
            gender=data_dict["male/female"],
            date_of_birth=data_dict["date_of_birth"],
            car_reg=data_dict["car_reg"],
            penalty_points=data_dict["number_of_penalty_points"],
            car_make=data_dict["car_make"],
            car_model=data_dict["car_model"]
        )

        print(f"[✔] Employee loaded from sheet '{self.sheet_name}': {self.employee}")
        return self.employee

    def validate_fields_not_empty(self, data_dict, required_fields):
        """Check all required fields are filled (not NaN or empty)."""
        for field in required_fields:
            val = data_dict.get(field)
            if pd.isna(val) or str(val).strip() == "":
                raise ValueError(f"Field '{field}' is empty in sheet '{self.sheet_name}'.")
        print("[VALIDATION] All required fields are filled.")

