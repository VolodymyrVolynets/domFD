import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from tkinter import ttk

class SettingsManager:
    def __init__(self, filepath="settings.json"):
        self.filepath = filepath
        self.default_settings = {
            "franchise_name": "",
            "shop_name": "",
            "store_manager_name": "",
            "date": datetime.today().strftime("%d/%m/%Y")
        }
        self.settings = {}
        self._load_or_create()

    def _load_or_create(self):
        if not os.path.exists(self.filepath):
            self.settings = self.default_settings.copy()
            self._save()
        else:
            try:
                with open(self.filepath, "r", encoding="utf-8") as f:
                    self.settings = json.load(f)
            except Exception:
                self.settings = self.default_settings.copy()
        # Ensure all required keys exist
        for k, v in self.default_settings.items():
            self.settings.setdefault(k, v)

    def _save(self):
        with open(self.filepath, "w", encoding="utf-8") as f:
            json.dump(self.settings, f, indent=4)

    def set(self, key, value):
        self.settings[key] = value
        self._save()

    def get(self, key):
        return self.settings.get(key, "")
