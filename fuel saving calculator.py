"""Fuel Saving Calculator and AOA Optimizer.

This desktop application combines voyage fuel saving calculations with the
hydrodynamic AOA optimisation graph. 
"""

import os
import tkinter as tk
from tkinter import messagebox
from typing import Callable, Dict, Iterable, Optional

import numpy as np
import pandas as pd
from scipy.interpolate import interp1d, make_interp_spline

# Matplotlib UI bridge integration setup
import matplotlib

matplotlib.use('TkAgg')
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from mpl_toolkits.mplot3d import Axes3D


class CyberMarineOptimizer:

    def resolve_local_file(self, preferred_name, alternatives=None):
        """Return the first existing file from cwd/script folder, else preferred_name."""
        alternatives = alternatives or []
        candidates = [preferred_name] + [name for name in alternatives if name != preferred_name]
        base_dirs = []
        try:
            base_dirs.append(os.getcwd())
        except Exception:
            pass
        try:
            base_dirs.append(os.path.dirname(os.path.abspath(__file__)))
        except Exception:
            pass

        for base_dir in base_dirs:
            for name in candidates:
                candidate = os.path.join(base_dir, name)
                if os.path.exists(candidate):
                    return candidate

        for name in candidates:
            if os.path.exists(name):
                return name

        return preferred_name

    def __init__(self, root):
        self.root = root
        self.root.title("Fuel Saving Calculator & AOA Optimizer")
        self.root.geometry("1450x920")
        self.root.minsize(1180, 760)
        self.root.option_add("*Font", "{Segoe UI} 10")
        self.nav_buttons = {}

        # State tracking — whether any voyage case has been run
        self.case_has_run = False

        # Target Assets Registry
        self.sail_file = self.resolve_local_file("sail properties.xlsx", ["sail properties(1).xlsx"])
        self.prop_file = self.resolve_local_file("propeller open water curve.xlsx", ["propeller open water curve(1).xlsx"])
        self.resistance_file = self.resolve_local_file("ship total resistance.xlsx")
        self.voyage_file = self.resolve_local_file("voyage details.xlsx")
        self.init_default_spreadsheets()

        self.theme_name = "light"
        self.theme_var = tk.StringVar(value=self.theme_name)
        self.apply_theme_palette(self.theme_name)
        self.root.configure(bg=self.bg_dark)

        # Interactive 3D View Boundary Dimension Trackers
        self.orig_xlim = None
        self.orig_ylim = None
        self.orig_zlim = None
        self.ax_3d = None
        self.canvas_3d = None

        # Pre-declare all frame containers to avoid structural initialization errors
        self.frame_fuel = tk.Frame(self.root, bg=self.bg_dark)
        self.frame_inputs = tk.Frame(self.root, bg=self.bg_dark)
        self.frame_sail = tk.Frame(self.root, bg=self.bg_dark)
        self.frame_prop = tk.Frame(self.root, bg=self.bg_dark)
        self.frame_resistance = tk.Frame(self.root, bg=self.bg_dark)
        self.frame_opt = tk.Frame(self.root, bg=self.bg_dark)

        self.fuel_aoa_plot_box = tk.Frame(self.frame_fuel, bg=self.bg_dark)
        self.sail_plot_box = tk.Frame(self.frame_sail, bg=self.bg_dark)
        self.prop_plot_box = tk.Frame(self.frame_prop, bg=self.bg_dark)
        self.resistance_plot_box = tk.Frame(self.frame_resistance, bg=self.bg_dark)
        self.opt_plot_box = tk.Frame(self.frame_opt, bg=self.bg_dark)

        # Assemble Subwindow Navigation & Sub-Panels
        self.build_navigation_bar()
        self.build_inputs_panel()
        self.setup_sail_ui_layout()
        self.setup_propeller_ui_layout()
        self.setup_resistance_ui_layout()
        self.setup_optimization_ui_layout()
        self.setup_fuel_saving_ui_layout()

        # Show fuel calculator as the first/default landing tab
        self.show_fuel_tab()

    def apply_theme_palette(self, theme_name):
        """Load the active colour palette.

        The rest of the application reads colours from these attributes, so the
        look and feel can be changed centrally without rewriting every widget.
        """
        self.theme_name = theme_name
        palettes = {
            "dark": {
                "bg_dark": "#0b1120",
                "bg_card": "#111827",
                "input_bg": "#0f172a",
                "table_alt": "#162033",
                "fg_green": "#22c55e",
                "fg_blue": "#38bdf8",
                "fg_white": "#e5e7eb",
                "text_muted": "#94a3b8",
                "hover_bg": "#1f2937",
                "border": "#263244",
                "grid": "#334155",
                "danger": "#ef4444",
                "warning": "#f59e0b",
                "plot_alt": "#f97316",
                "button_text": "#f8fafc",
                "disabled_bg": "#1a1f2e",
                "disabled_fg": "#4a5568",
            },
            "light": {
                "bg_dark": "#f4f7fb",
                "bg_card": "#ffffff",
                "input_bg": "#eef2f7",
                "table_alt": "#f8fafc",
                "fg_green": "#047857",
                "fg_blue": "#0369a1",
                "fg_white": "#0f172a",
                "text_muted": "#64748b",
                "hover_bg": "#e2e8f0",
                "border": "#cbd5e1",
                "grid": "#cbd5e1",
                "danger": "#b91c1c",
                "warning": "#b45309",
                "plot_alt": "#c2410c",
                "button_text": "#ffffff",
                "disabled_bg": "#e2e8f0",
                "disabled_fg": "#94a3b8",
            },
        }

        self.palette = palettes.get(theme_name, palettes["dark"])
        self.bg_dark = self.palette["bg_dark"]
        self.bg_card = self.palette["bg_card"]
        self.input_bg = self.palette["input_bg"]
        self.table_alt = self.palette["table_alt"]
        self.fg_green = self.palette["fg_green"]
        self.fg_blue = self.palette["fg_blue"]
        self.fg_white = self.palette["fg_white"]
        self.text_muted = self.palette["text_muted"]
        self.hover_bg = self.palette["hover_bg"]
        self.border_color = self.palette["border"]
        self.grid_color = self.palette["grid"]
        self.danger_color = self.palette["danger"]
        self.warning_color = self.palette["warning"]
        self.plot_alt = self.palette["plot_alt"]
        self.button_text = self.palette["button_text"]
        self.disabled_bg = self.palette["disabled_bg"]
        self.disabled_fg = self.palette["disabled_fg"]

    # -------------------------------------------------------------------------
    # UI helper methods
    # -------------------------------------------------------------------------
    # These helpers reduce repeated Tkinter configuration code and make it easy
    # to adjust spacing, colours, and typography across the full application.

    def clear_children(self, widget):
        """Remove all child widgets from a container."""
        for child in widget.winfo_children():
            child.destroy()

    def make_section(self, parent, title, *, padx=12, pady=10):
        """Create a consistent card/section container."""
        section = tk.LabelFrame(
            parent,
            text=title,
            bg=self.bg_card,
            fg=self.fg_green,
            font=("Segoe UI", 10, "bold"),
            bd=1,
            relief="solid",
            padx=padx,
            pady=pady,
            highlightthickness=1,
            highlightbackground=self.border_color,
        )
        return section

    def make_label(self, parent, text, *, role="body", bg=None, **kwargs):
        """Create a themed label.

        role controls default typography: title, subtitle, muted, metric, or body.
        """
        role_styles = {
            "title": {"fg": self.fg_white, "font": ("Segoe UI", 16, "bold")},
            "subtitle": {"fg": self.text_muted, "font": ("Segoe UI", 9)},
            "muted": {"fg": self.text_muted, "font": ("Segoe UI", 9)},
            "metric": {"fg": self.fg_white, "font": ("Consolas", 11, "bold")},
            "body": {"fg": self.fg_white, "font": ("Segoe UI", 10)},
        }
        options = role_styles.get(role, role_styles["body"]).copy()
        options.update(kwargs)
        return tk.Label(parent, text=text, bg=bg or self.bg_card, **options)

    def make_button(self, parent, text, command=None, *, variant="secondary", **kwargs):
        """Create a consistent application button.

        Keyword options passed by the caller intentionally override the defaults.
        This avoids duplicate Tkinter keyword errors when a caller customises
        spacing, font, border, or active colours.
        """
        variants = {
            "primary": {"bg": self.fg_green, "fg": "#04130b", "activebackground": self.fg_blue},
            "secondary": {"bg": self.input_bg, "fg": self.fg_blue, "activebackground": self.hover_bg},
            "ghost": {"bg": self.bg_card, "fg": self.text_muted, "activebackground": self.hover_bg},
            "danger": {"bg": self.input_bg, "fg": self.danger_color, "activebackground": self.hover_bg},
        }

        # Start with safe defaults, layer on the variant, then apply caller
        # overrides once. This prevents "multiple values for keyword argument"
        # errors for options such as padx/pady.
        options = {
            "relief": "flat",
            "bd": 0,
            "font": ("Segoe UI", 9, "bold"),
            "cursor": "hand2",
            "padx": 12,
            "pady": 7,
            "activeforeground": self.button_text,
        }
        options.update(variants.get(variant, variants["secondary"]))
        options.update(kwargs)

        return tk.Button(parent, text=text, command=command, **options)

    def make_table_cell(self, parent, text, *, header=False, alt=False, width=13, fg=None):
        """Create a styled table cell used by the voyage case table."""
        bg = self.bg_card if header else (self.table_alt if alt else self.input_bg)
        return tk.Label(
            parent,
            text=text,
            bg=bg,
            fg=fg or (self.fg_green if header else self.fg_white),
            font=("Segoe UI", 9, "bold") if header else ("Consolas", 9),
            bd=1,
            relief="solid",
            highlightthickness=0,
            padx=8,
            pady=7,
            width=width,
        )

    def format_number(self, value, decimals=3, suffix=""):
        """Format numbers for the UI while keeping missing values readable."""
        try:
            if value is None or pd.isna(value):
                return "--"
            return f"{float(value):.{decimals}f}{suffix}"
        except Exception:
            return "--"

    def set_active_tab(self, tab_name):
        """Highlight the active navigation tab."""
        for name, button in getattr(self, "nav_buttons", {}).items():
            is_active = name == tab_name
            # Check if this is a frozen tab
            if name in ("Inputs", "AOA") and not self.case_has_run:
                if is_active:
                    button.config(bg=self.disabled_bg, fg=self.disabled_fg)
                else:
                    button.config(bg=self.disabled_bg, fg=self.disabled_fg)
            else:
                button.config(
                    bg=self.fg_blue if is_active else self.bg_card,
                    fg=self.bg_dark if is_active else self.text_muted,
                )

    def load_data(self, file_path):
        if os.path.exists(file_path):
            return pd.read_excel(file_path)
        csv_variant = file_path + " - Sheet1.csv"
        if os.path.exists(csv_variant):
            return pd.read_csv(csv_variant)
        pure_csv = file_path.replace('.xlsx', '.csv')
        if os.path.exists(pure_csv):
            return pd.read_csv(pure_csv)
        raise FileNotFoundError(f"Missing Required Resource Matrix File: {file_path}")

    def init_default_spreadsheets(self):
        try:
            if not os.path.exists(self.sail_file) and not os.path.exists(self.sail_file + " - Sheet1.csv"):
                pd.DataFrame({
                    "Angle of Attack (∘)": [0, 5, 10, 15, 20, 25, 30, 45, 60, 90, 120, 150, 180],
                    "Lift Coefficient (Cl​)": [0, 0.22, 0.72, 1.06, 1.15, 1.25, 1.35, 1.44, 1.48, 1.3, 0.92, 0.52, 0],
                    "Drag Coefficient (Cd​)": [0.02, 0.03, 0.05, 0.12, 0.2, 0.28, 0.38, 0.55, 0.7, 1.01, 1.09, 1.05,
                                               0.9]
                }).to_excel(self.sail_file, index=False)

            if not os.path.exists(self.prop_file) and not os.path.exists(self.prop_file + " - Sheet1.csv"):
                pd.DataFrame({
                    "Advance Coefficient (J)": [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.88, 0.92],
                    "Thrust Coefficient (KT​)": [0.435, 0.395, 0.355, 0.312, 0.268, 0.222, 0.174, 0.123, 0.07, 0.024,
                                                 0],
                    "Torque Coefficient (10KQ​)": [0.65, 0.595, 0.54, 0.482, 0.422, 0.362, 0.3, 0.235, 0.165, 0.105,
                                                   0.075],
                    "Open Water Efficiency (Eta_o)": [0, 0.106, 0.209, 0.31, 0.404, 0.488, 0.554, 0.584, 0.54, 0.322, 0]
                }).to_excel(self.prop_file, index=False)

            if not os.path.exists(self.resistance_file) and not os.path.exists(self.resistance_file + " - Sheet1.csv"):
                pd.DataFrame({
                    "Speed (V in knots)": [0, 2, 4, 6, 8, 10, 12, 14, 16, 18],
                    "Total Resistance without WAP (RT in kN)": [0, 15, 45, 80, 125, 135, 240, 300, 520, 760],
                    "Sail Resistance (kN)": [0] * 10,
                    "Rudder resistance": [0] * 10,
                    "Total resistance with WAP (kN)": [0, 15, 45, 80, 125, 135, 240, 300, 520, 760]
                }).to_excel(self.resistance_file, index=False)
        except Exception:
            pass

    def build_navigation_bar(self):
        """Create the top navigation bar.

        Navigation metadata lives in one list, which makes it easy to add,
        remove, or rename tabs later.
        """
        self.nav_frame = tk.Frame(self.root, bg=self.bg_dark, bd=0)
        self.nav_frame.pack(side='top', fill='x', padx=18, pady=(14, 8))
        title_box = tk.Frame(self.nav_frame, bg=self.bg_dark)
        title_box.pack(side='left', padx=(0, 18))
        self.make_label(title_box, "Fuel Saving & AOA Optimizer", role="title", bg=self.bg_dark).pack(anchor='w')
        self.make_label(title_box, "Voyage case analysis, fuel savings, and propeller operating point visualisation",
                        role="muted", bg=self.bg_dark).pack(anchor='w')

        tabs = [
            ("Fuel saving", "FUEL SAVING CALCULATOR", self.show_fuel_tab),
            ("Inputs", "KINEMATICS", self.show_inputs_tab),
            ("Sail", "SAIL CHARACTERISTICS", self.show_sail_tab),
            ("Propeller", "PROPELLER OPEN WATER", self.show_propeller_tab),
            ("Resistance", "RESISTANCE CURVE", self.show_resistance_tab),
            ("AOA", "AOA OPTIMISER", self.show_optimization_tab),
        ]

        tab_strip = tk.Frame(self.nav_frame, bg=self.bg_dark)
        tab_strip.pack(side='left', fill='x', expand=True)

        self.nav_buttons = {}
        for tab_name, label, command in tabs:
            btn = self.make_button(tab_strip, label, command=command, variant="ghost", padx=12, pady=8)
            btn.pack(side='left', padx=3)
            # Frozen tabs start greyed out
            if tab_name in ("Inputs", "AOA"):
                btn.config(bg=self.disabled_bg, fg=self.disabled_fg)
            btn.bind("<Enter>", lambda e, b=btn, n=tab_name: self._nav_enter(b, n))
            btn.bind("<Leave>", lambda e, n=tab_name, b=btn: self.set_active_tab(getattr(self, "active_tab", "Fuel saving")))
            self.nav_buttons[tab_name] = btn

        self.settings_btn = self.make_button(self.nav_frame, "Settings", command=self.open_settings_window,
                                             variant="secondary", padx=12, pady=8)
        self.settings_btn.pack(side='right', padx=(8, 0))

    def _nav_enter(self, btn, tab_name):
        """Hover effect for nav buttons, respecting frozen state."""
        if tab_name in ("Inputs", "AOA") and not self.case_has_run:
            btn.config(bg=self.disabled_bg, fg=self.disabled_fg)
        else:
            btn.config(bg=self.hover_bg, fg=self.fg_white)

    def open_settings_window(self):
        if hasattr(self, "settings_window") and self.settings_window.winfo_exists():
            self.settings_window.lift()
            return

        self.settings_window = tk.Toplevel(self.root)
        self.settings_window.title("Settings")
        self.settings_window.geometry("280x180")
        self.settings_window.resizable(False, False)
        self.settings_window.configure(bg=self.bg_card)
        self.settings_window.transient(self.root)

        tk.Label(self.settings_window, text="Settings", bg=self.bg_card, fg=self.fg_white,
                 font=('Segoe UI', 13, 'bold')).pack(anchor='w', padx=18, pady=(16, 6))
        tk.Label(self.settings_window, text="Color theme", bg=self.bg_card, fg=self.fg_blue,
                 font=('Segoe UI', 10, 'bold')).pack(anchor='w', padx=18, pady=(8, 4))

        options = tk.Frame(self.settings_window, bg=self.bg_card)
        options.pack(fill='x', padx=18, pady=4)
        for label, value in [("Dark", "dark"), ("Light", "light")]:
            rb = tk.Radiobutton(options, text=label, value=value, variable=self.theme_var,
                                command=lambda v=value: self.set_theme(v),
                                bg=self.bg_card, fg=self.fg_white, selectcolor=self.input_bg,
                                activebackground=self.bg_card, activeforeground=self.fg_green,
                                font=('Segoe UI', 10), anchor='w')
            rb.pack(anchor='w', pady=3)

        tk.Button(self.settings_window, text="Close", command=self.settings_window.destroy,
                  bg=self.bg_dark, fg=self.fg_blue, activebackground=self.hover_bg,
                  activeforeground=self.fg_white, font=('Segoe UI', 9, 'bold'),
                  relief='solid', bd=1, padx=12, pady=4).pack(anchor='e', padx=18, pady=(10, 0))

    def set_theme(self, theme_name):
        self.apply_theme_palette(theme_name)
        self.theme_var.set(theme_name)
        self.refresh_theme()

    def sync_data(self):
        try:
            # Force reload check on both database files
            self.load_data(self.sail_file)
            self.load_data(self.prop_file)
            self.load_resistance_dataframe()
            if hasattr(self, "sail_grid_entries"):
                self.load_sail_grid_from_excel(silent=True)
            if hasattr(self, "prop_grid_entries"):
                self.load_propeller_grid_from_excel(silent=True)
            if hasattr(self, "resistance_grid_entries"):
                self.load_resistance_grid_from_excel(silent=True)
            messagebox.showinfo("Data Sync Status", "Success! Matrix values re-synchronized from Excel sheets.")
        except Exception as e:
            messagebox.showerror("Sync Failed", f"Could not sync spreadsheets:\n{e}")

    def refresh_theme(self):
        self.root.configure(bg=self.bg_dark)
        self.apply_theme_to_widget(self.root)
        if hasattr(self, "settings_window") and self.settings_window.winfo_exists():
            self.settings_window.configure(bg=self.bg_card)
            self.apply_theme_to_widget(self.settings_window)

        for plot_box in [self.fuel_aoa_plot_box, self.sail_plot_box, self.prop_plot_box, self.resistance_plot_box, self.opt_plot_box]:
            try:
                plot_box.configure(bg=self.bg_dark)
            except Exception:
                pass

        if hasattr(self, "fuel_table_frame"):
            self.refresh_fuel_window()
        if hasattr(self, "sail_plot_box"):
            self.refresh_sail_plot()
        if hasattr(self, "prop_plot_box"):
            self.refresh_propeller_plot()
        if hasattr(self, "resistance_plot_box"):
            self.refresh_resistance_plot()
        if self.frame_opt.winfo_ismapped():
            self.compute_3d_convergence()

    def apply_theme_to_widget(self, widget):
        try:
            if isinstance(widget, tk.LabelFrame):
                widget.configure(bg=self.bg_card, fg=self.fg_green)
            elif isinstance(widget, tk.Frame):
                parent = widget.master
                parent_bg = None
                try:
                    parent_bg = parent.cget("bg")
                except Exception:
                    pass
                widget.configure(bg=self.bg_card if parent_bg == self.bg_card else self.bg_dark)
            elif isinstance(widget, tk.Text):
                widget.configure(bg=self.input_bg, fg=self.fg_green, insertbackground=self.fg_green)
            elif isinstance(widget, tk.Canvas):
                widget.configure(bg=self.bg_card)
            elif isinstance(widget, tk.Entry):
                widget.configure(bg=self.input_bg, fg=self.fg_green, insertbackground=self.fg_green)
                try:
                    widget.configure(disabledbackground=self.disabled_bg, disabledforeground=self.disabled_fg)
                except Exception:
                    pass
            elif isinstance(widget, tk.Radiobutton):
                widget.configure(bg=self.bg_card, fg=self.fg_white, selectcolor=self.input_bg,
                                 activebackground=self.bg_card, activeforeground=self.fg_green)
            elif isinstance(widget, tk.Button):
                widget.configure(bg=self.bg_card, fg=self.fg_blue, activebackground=self.hover_bg,
                                 activeforeground=self.fg_white)
            elif isinstance(widget, tk.Label):
                parent = widget.master
                parent_bg = None
                try:
                    parent_bg = parent.cget("bg")
                except Exception:
                    pass
                widget.configure(bg=self.bg_card if parent_bg == self.bg_card else self.bg_dark)
                if widget.cget("fg").lower() not in ("red", "#ff0000"):
                    widget.configure(fg=self.fg_blue)
        except Exception:
            pass

        for child in widget.winfo_children():
            self.apply_theme_to_widget(child)

    def hide_all_frames(self):
        """Hide all main content frames before showing the selected tab."""
        self.frame_fuel.pack_forget()
        self.frame_inputs.pack_forget()
        self.frame_sail.pack_forget()
        self.frame_prop.pack_forget()
        self.frame_resistance.pack_forget()
        self.frame_opt.pack_forget()

    def show_fuel_tab(self):
        self.active_tab = "Fuel saving"
        self.set_active_tab(self.active_tab)
        self.hide_all_frames()
        self.frame_fuel.pack(fill='both', expand=True, padx=18, pady=(0, 16))

    def show_inputs_tab(self):
        if not self.case_has_run:
            messagebox.showinfo("Kinematics Locked",
                                "Run a voyage case first to unlock the Kinematics tab.")
            return
        self.active_tab = "Inputs"
        self.set_active_tab(self.active_tab)
        self.hide_all_frames()
        self.frame_inputs.pack(fill='both', expand=True, padx=18, pady=(0, 16))

    def show_sail_tab(self):
        self.active_tab = "Sail"
        self.set_active_tab(self.active_tab)
        self.hide_all_frames()
        self.frame_sail.pack(fill='both', expand=True, padx=18, pady=(0, 16))
        self.build_sail_panel()

    def show_propeller_tab(self):
        self.active_tab = "Propeller"
        self.set_active_tab(self.active_tab)
        self.hide_all_frames()
        self.frame_prop.pack(fill='both', expand=True, padx=18, pady=(0, 16))
        self.build_propeller_panel()

    def show_resistance_tab(self):
        self.active_tab = "Resistance"
        self.set_active_tab(self.active_tab)
        self.hide_all_frames()
        self.frame_resistance.pack(fill='both', expand=True, padx=18, pady=(0, 16))
        self.refresh_resistance_plot()

    def show_optimization_tab(self):
        if not self.case_has_run:
            messagebox.showinfo("AOA Optimiser Locked",
                                "Run a voyage case first to unlock the AOA Optimiser tab.")
            return
        self.active_tab = "AOA"
        self.set_active_tab(self.active_tab)
        self.hide_all_frames()
        self.frame_opt.pack(fill='both', expand=True, padx=18, pady=(0, 16))
        self.compute_3d_convergence()

    def _unlock_frozen_tabs(self):
        """Called after a case is run to unfreeze Kinematics and AOA tabs."""
        self.case_has_run = True
        for tab_name in ("Inputs", "AOA"):
            btn = self.nav_buttons.get(tab_name)
            if btn:
                btn.config(bg=self.bg_card, fg=self.text_muted)
        self.set_active_tab(getattr(self, "active_tab", "Fuel saving"))

    # -------------------------------------------------------------------------
    # Kinematics / Inputs panel
    # -------------------------------------------------------------------------
    def build_inputs_panel(self):
        card = tk.LabelFrame(self.frame_inputs, text=" INPUTS ", bg=self.bg_card,
                             fg=self.fg_green,
                             font=('Segoe UI', 11, 'bold'), bd=2, padx=15, pady=15)
        card.pack(side='left', fill='both', expand=True, padx=15, pady=15)

        self.kinematics_active_case_label = tk.Label(card, text="Active Case: None", bg=self.bg_card, fg=self.fg_blue, font=('Segoe UI', 12, 'bold'))
        self.kinematics_active_case_label.grid(row=0, column=0, columnspan=2, pady=(0, 10))

        input_configs = [
            ("True Wind Speed, vw (m/s)", ""),
            ("True Wind Angle, theta_w (deg)", ""),
            ("Desired Ship Speed, vs (knots)", ""),
            ("Density of Water, rho_w (kg/m3)", "1025.0"),
            ("Density of Air, rho_a (kg/m3)", "1.025"),
            ("Sail Area, As (m2)", "1000.0"),
            ("Propeller Diameter, D (m)", "4.0"),
            ("Wake Fraction, w", "0.312"),
            ("Thrust Deduction Fraction, t", "0.1"),
            ("Relative Rotative Efficiency, eta_R", "1.0")
        ]

        # Track which fields are frozen (first 3)
        self.frozen_input_keys = [
            "True Wind Speed, vw (m/s)",
            "True Wind Angle, theta_w (deg)",
            "Desired Ship Speed, vs (knots)",
        ]

        self.entries = {}
        for idx, (label, val) in enumerate(input_configs, start=1):
            lbl = tk.Label(card, text=label + ":", bg=self.bg_card, fg=self.fg_white, font=('Segoe UI', 10, 'bold'))
            lbl.grid(row=idx, column=0, sticky='w', pady=6, padx=5)

            is_frozen = label in self.frozen_input_keys
            ent = tk.Entry(card, bg=self.disabled_bg if is_frozen else self.input_bg,
                           fg=self.disabled_fg if is_frozen else self.fg_green,
                           insertbackground=self.fg_green,
                           font=('Consolas', 10, 'bold'), bd=1, relief='solid', width=15,
                           disabledbackground=self.disabled_bg,
                           disabledforeground=self.disabled_fg,
                           state='disabled' if is_frozen else 'normal')
            if val:
                ent.insert(0, val)
            ent.grid(row=idx, column=1, sticky='w', pady=6, padx=10)
            self.entries[label] = ent

            # Bind click on frozen entries to show message
            if is_frozen:
                ent.bind("<Button-1>", lambda e: messagebox.showinfo(
                    "Input Locked",
                    "This value is set by the voyage case.\nRun a voyage case from the Fuel Saving Calculator."))

        btn = tk.Button(card, text="Recalculate kinematics", command=self.calculate_kinematics,
                        bg=self.bg_dark, fg=self.fg_green, font=('Segoe UI', 10, 'bold'),
                        bd=1, relief='solid', activebackground=self.fg_green, activeforeground=self.bg_dark, padx=15,
                        pady=6)
        btn.grid(row=len(input_configs) + 1, column=0, columnspan=2, pady=20)

        self.monitor_card = tk.LabelFrame(self.frame_inputs, text=" CALCULATIONS ", bg=self.bg_card,
                                          fg=self.fg_blue,
                                          font=('Segoe UI', 11, 'bold'), bd=2, padx=15, pady=15)
        self.monitor_card.pack(side='right', fill='both', expand=True, padx=15, pady=15)

        self.labels_hud = {}
        hud_lines = ["Velocity in x-dir (vx)", "Velocity in y-dir (vy)", "Apparent Wind Speed (va)",
                     "Apparent Wind Angle (theta)", "Velocity of Advance (vadv)", "Hull Efficiency (eta_h)",
                     "Interpolated Total Resistance (RT)"]
        for line in hud_lines:
            lbl = tk.Label(self.monitor_card, text=f"{line} : --", bg=self.bg_card, fg=self.fg_blue,
                           font=('Consolas', 11, 'bold'))
            lbl.pack(anchor='w', pady=10)
            self.labels_hud[line] = lbl

    def set_case_inputs(self, case):
        """Set the first 3 frozen input fields from a voyage case."""
        frozen_map = {
            "True Wind Speed, vw (m/s)": case['wind_speed'],
            "True Wind Angle, theta_w (deg)": case['wind_angle'],
            "Desired Ship Speed, vs (knots)": case['ship_speed'],
        }
        for key, val in frozen_map.items():
            ent = self.entries[key]
            ent.config(state='normal')
            ent.delete(0, tk.END)
            ent.insert(0, f"{val:g}")
            ent.config(state='disabled')

    def calculate_kinematics(self):
        """Calculate apparent wind, advance velocity, hull efficiency, and RT."""
        try:
            vw = float(self.entries["True Wind Speed, vw (m/s)"].get())
            tw_deg = float(self.entries["True Wind Angle, theta_w (deg)"].get())
            vs_knots = float(self.entries["Desired Ship Speed, vs (knots)"].get())
            w = float(self.entries["Wake Fraction, w"].get())
            t = float(self.entries["Thrust Deduction Fraction, t"].get())

            vs_ms = vs_knots * 0.5144
            tw_rad = np.radians(tw_deg)

            vx = -vs_ms + vw * np.cos(tw_rad)
            vy = vw * np.sin(tw_rad)
            va = np.sqrt(vx ** 2 + vy ** 2)
            theta_deg = np.degrees(np.arctan2(vy, vx))
            vadv = vs_ms * (1 - w)

            eta_h = (1 - t) / (1 - w)

            self.labels_hud["Velocity in x-dir (vx)"].config(text=f"Velocity in x-dir (vx) : {vx:.4f} m/s",
                                                             fg=self.fg_white)
            self.labels_hud["Velocity in y-dir (vy)"].config(text=f"Velocity in y-dir (vy) : {vy:.4f} m/s",
                                                             fg=self.fg_white)
            self.labels_hud["Apparent Wind Speed (va)"].config(text=f"Apparent Wind Speed (va) : {va:.4f} m/s",
                                                               fg=self.fg_green)
            self.labels_hud["Apparent Wind Angle (theta)"].config(
                text=f"Apparent Wind Angle (theta) : {theta_deg:.2f}°", fg=self.fg_green)
            self.labels_hud["Velocity of Advance (vadv)"].config(text=f"Velocity of Advance (vadv) : {vadv:.4f} m/s",
                                                                 fg=self.fg_blue)
            self.labels_hud["Hull Efficiency (eta_h)"].config(text=f"Hull Efficiency (eta_h) : {eta_h:.4f} ",
                                                              fg=self.fg_blue)
            RT = self.get_total_resistance(vs_knots)
            self.labels_hud["Interpolated Total Resistance (RT)"].config(
                text=f"Interpolated Total Resistance (RT) : {RT:.2f} N ({RT / 1000:.3f} kN)",
                fg=self.fg_green
            )

            return vx, vy, va, theta_deg, vadv, vs_ms, eta_h
        except Exception as e:
            messagebox.showerror("HUD Calculations Exception", f"Format verification checkpoint failed:\n{e}")
            return None

    # =========================================================================
    # GRID-BASED EDITABLE TABLE BUILDER  (replaces old Text-based table)
    # =========================================================================
    def build_grid_table_panel(self, parent, title, help_text, columns, grid_attr,
                               apply_cmd, reload_cmd, save_cmd):
        """Build a grid-based editable table with Entry widgets and borders.

        Args:
            columns: list of column header strings
            grid_attr: attribute name to store list-of-row-entries on self
        Returns:
            The editor LabelFrame widget.
        """
        editor = tk.LabelFrame(parent, text=title, bg=self.bg_card, fg=self.fg_green,
                               font=('Segoe UI', 10, 'bold'), bd=2, padx=10, pady=10)
        editor.grid(row=0, column=0, sticky='nsew', padx=(0, 10), pady=0)
        editor.configure(width=480, height=560)
        editor.grid_propagate(False)

        tk.Label(editor, text=help_text, bg=self.bg_card, fg=self.fg_white,
                 font=('Segoe UI', 9), wraplength=420, justify='left').pack(anchor='w', pady=(0, 6))

        # Scrollable table area
        table_outer = tk.Frame(editor, bg=self.bg_card)
        table_outer.pack(fill='both', expand=True)

        canvas = tk.Canvas(table_outer, bg=self.bg_card, highlightthickness=0)
        y_scroll = tk.Scrollbar(table_outer, orient='vertical', command=canvas.yview)
        x_scroll = tk.Scrollbar(editor, orient='horizontal', command=canvas.xview)
        canvas.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        canvas.pack(side='left', fill='both', expand=True)
        y_scroll.pack(side='right', fill='y')
        x_scroll.pack(fill='x')

        inner = tk.Frame(canvas, bg=self.bg_card)
        win_id = canvas.create_window((0, 0), window=inner, anchor='nw')
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win_id, width=max(e.width, inner.winfo_reqwidth())))

        # Store references
        setattr(self, f"_{grid_attr}_inner", inner)
        setattr(self, f"_{grid_attr}_columns", columns)
        setattr(self, f"_{grid_attr}_canvas", canvas)
        setattr(self, grid_attr, [])  # list of lists of Entry widgets

        # Column headers
        for col_idx, col_name in enumerate(columns):
            hdr = tk.Label(inner, text=col_name, bg=self.bg_card, fg=self.fg_green,
                           font=('Segoe UI', 9, 'bold'), bd=1, relief='solid',
                           padx=6, pady=5, width=14)
            hdr.grid(row=0, column=col_idx, sticky='nsew', padx=0, pady=0)
            inner.grid_columnconfigure(col_idx, weight=1)

        # Store status label attr name
        status_attr = f"_{grid_attr}_status"
        status = tk.Label(editor, text="", bg=self.bg_card, fg=self.fg_blue,
                          font=('Consolas', 9, 'bold'), justify='left', wraplength=420)
        status.pack(anchor='w', pady=(4, 0))
        setattr(self, status_attr, status)

        # Controls
        controls = tk.Frame(editor, bg=self.bg_card)
        controls.pack(fill='x', pady=8)
        for label, cmd in [
            ("APPLY TABLE", apply_cmd),
            ("RELOAD EXCEL", reload_cmd),
            ("SAVE TO EXCEL", save_cmd),
        ]:
            tk.Button(controls, text=label, command=cmd, bg=self.bg_dark, fg=self.fg_blue,
                      activebackground=self.fg_blue, activeforeground=self.bg_dark,
                      font=('Segoe UI', 9, 'bold'), bd=1, relief='solid', padx=8, pady=4).pack(side='left', padx=3)

        # Add / Delete row buttons
        row_controls = tk.Frame(editor, bg=self.bg_card)
        row_controls.pack(fill='x', pady=(0, 4))
        tk.Button(row_controls, text="+ Add Row",
                  command=lambda: self._grid_add_row(grid_attr),
                  bg=self.bg_dark, fg=self.fg_green,
                  activebackground=self.fg_green, activeforeground=self.bg_dark,
                  font=('Segoe UI', 9, 'bold'), bd=1, relief='solid', padx=8, pady=3).pack(side='left', padx=3)
        tk.Button(row_controls, text="− Delete Last Row",
                  command=lambda: self._grid_delete_last_row(grid_attr),
                  bg=self.bg_dark, fg=self.danger_color,
                  activebackground=self.danger_color, activeforeground=self.bg_dark,
                  font=('Segoe UI', 9, 'bold'), bd=1, relief='solid', padx=8, pady=3).pack(side='left', padx=3)

        return editor

    def _grid_add_row(self, grid_attr, values=None):
        """Add a new row of Entry widgets to a grid table."""
        inner = getattr(self, f"_{grid_attr}_inner")
        columns = getattr(self, f"_{grid_attr}_columns")
        rows = getattr(self, grid_attr)
        row_idx = len(rows) + 1  # +1 for header

        row_entries = []
        alt = (len(rows) % 2 == 1)
        bg = self.table_alt if alt else self.input_bg

        for col_idx in range(len(columns)):
            ent = tk.Entry(inner, bg=bg, fg=self.fg_white,
                           insertbackground=self.fg_green,
                           font=('Consolas', 9), bd=1, relief='solid',
                           width=14, justify='center')
            ent.grid(row=row_idx, column=col_idx, sticky='nsew', padx=0, pady=0)
            if values and col_idx < len(values):
                ent.insert(0, str(values[col_idx]))
            row_entries.append(ent)

        rows.append(row_entries)

    def _grid_delete_last_row(self, grid_attr):
        """Remove the last row from a grid table."""
        rows = getattr(self, grid_attr)
        if not rows:
            return
        last_row = rows.pop()
        for ent in last_row:
            ent.destroy()

    def _grid_populate(self, grid_attr, df):
        """Clear and repopulate a grid table from a DataFrame."""
        rows = getattr(self, grid_attr)
        # Remove existing data rows
        for row in rows:
            for ent in row:
                ent.destroy()
        rows.clear()

        # Add rows from DataFrame
        for _, row_data in df.iterrows():
            values = []
            for val in row_data:
                if pd.isna(val):
                    values.append("")
                elif isinstance(val, (int, float, np.integer, np.floating)):
                    values.append(f"{val:g}")
                else:
                    values.append(str(val))
            self._grid_add_row(grid_attr, values)

    def _grid_to_dataframe(self, grid_attr, columns):
        """Read a grid table's Entry widgets into a DataFrame."""
        rows = getattr(self, grid_attr)
        data = []
        for row_entries in rows:
            row_vals = []
            for ent in row_entries:
                row_vals.append(ent.get().strip())
            # Skip fully empty rows
            if all(v == "" for v in row_vals):
                continue
            try:
                numeric_row = [float(v) for v in row_vals]
                data.append(numeric_row)
            except ValueError:
                continue

        if not data:
            raise ValueError("Enter at least one numeric data row.")
        return pd.DataFrame(data, columns=columns)

    # =========================================================================
    # SAIL CHARACTERISTICS
    # =========================================================================
    def setup_sail_ui_layout(self):
        top = tk.Frame(self.frame_sail, bg=self.bg_dark)
        top.pack(side='top', fill='x', padx=15, pady=5)
        tk.Label(top, text="DATA LAYER INTERFACE: sail properties matrix", bg=self.bg_dark, fg=self.fg_blue,
                 font=('Segoe UI', 9, 'italic')).pack(side='left')

        body = tk.Frame(self.frame_sail, bg=self.bg_dark)
        body.pack(fill='both', expand=True, padx=15, pady=10)
        body.grid_columnconfigure(0, minsize=480, weight=0)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        self.build_grid_table_panel(
            body,
            " EDITABLE SAIL TABLE ",
            "Edit the sail data below. Columns: AOA (deg), CL, CD.",
            ["AOA (deg)", "CL", "CD"],
            "sail_grid_entries",
            self.refresh_sail_plot,
            lambda: self.load_sail_grid_from_excel(silent=False),
            self.save_sail_table_to_excel
        )
        self.sail_plot_box = tk.Frame(body, bg=self.bg_dark)
        self.sail_plot_box.grid(row=0, column=1, sticky='nsew')
        self.load_sail_grid_from_excel(silent=True)

    def load_sail_dataframe(self):
        try:
            df = self.load_data(self.sail_file)
            if df.shape[1] < 3:
                raise ValueError("sail properties.xlsx must contain AOA, CL and CD columns.")
            sail_df = pd.DataFrame({
                "AOA (deg)": pd.to_numeric(df.iloc[:, 0], errors='coerce'),
                "Lift Coefficient (CL)": pd.to_numeric(df.iloc[:, 1], errors='coerce'),
                "Drag Coefficient (CD)": pd.to_numeric(df.iloc[:, 2], errors='coerce')
            }).dropna()
            if sail_df.empty:
                raise ValueError("No numeric sail rows found.")
            return sail_df.sort_values("AOA (deg)")
        except Exception:
            return pd.DataFrame({
                "AOA (deg)": [0, 5, 10, 15, 20, 25, 30, 45, 60, 90, 120, 150, 180],
                "Lift Coefficient (CL)": [0, 0.22, 0.72, 1.06, 1.15, 1.25, 1.35, 1.44, 1.48, 1.3, 0.92, 0.52, 0],
                "Drag Coefficient (CD)": [0.02, 0.03, 0.05, 0.12, 0.2, 0.28, 0.38, 0.55, 0.7, 1.01, 1.09, 1.05, 0.9]
            })

    def load_sail_grid_from_excel(self, silent=False):
        try:
            df = self.load_sail_dataframe()
            self._grid_populate("sail_grid_entries", df)
            self.refresh_sail_plot()
            if not silent:
                messagebox.showinfo("Sail Table", "Sail table reloaded from Excel.")
        except Exception as e:
            if hasattr(self, "_sail_grid_entries_status"):
                self._sail_grid_entries_status.config(text=f"Sail load failed: {e}", fg='red')
            if not silent:
                messagebox.showerror("Sail Load Failed", str(e))

    def parse_sail_table(self):
        if not hasattr(self, "sail_grid_entries") or not self.sail_grid_entries:
            return self.load_sail_dataframe()
        df = self._grid_to_dataframe(
            "sail_grid_entries",
            ["AOA (deg)", "Lift Coefficient (CL)", "Drag Coefficient (CD)"]
        )
        return df.groupby("AOA (deg)", as_index=False).mean().sort_values("AOA (deg)")

    def save_sail_table_to_excel(self):
        try:
            df = self.parse_sail_table()
            out_df = pd.DataFrame({
                "Angle of Attack (deg)": df["AOA (deg)"],
                "Lift Coefficient (Cl)": df["Lift Coefficient (CL)"],
                "Drag Coefficient (Cd)": df["Drag Coefficient (CD)"]
            })
            out_df.to_excel(self.sail_file, index=False)
            messagebox.showinfo("Sail Table Saved", f"Saved {len(df)} rows to {self.sail_file}.")
        except Exception as e:
            messagebox.showerror("Sail Save Failed", str(e))

    def build_sail_panel(self):
        self.refresh_sail_plot()

    def refresh_sail_plot(self):
        for w in self.sail_plot_box.winfo_children(): w.destroy()
        try:
            df = self.parse_sail_table()
            aoa = df["AOA (deg)"].to_numpy(dtype=float)
            cl = df["Lift Coefficient (CL)"].to_numpy(dtype=float)
            cd = df["Drag Coefficient (CD)"].to_numpy(dtype=float)

            if np.max(aoa) < 180.0:
                aoa = np.append(aoa, 180.0);
                cl = np.append(cl, 0.0);
                cd = np.append(cd, cd[-1])

            aoa_smooth = np.linspace(aoa.min(), aoa.max(), 300)
            spline_order = min(3, len(aoa) - 1)
            spline_cl = make_interp_spline(aoa, cl, k=spline_order)
            spline_cd = make_interp_spline(aoa, cd, k=spline_order)
            cl_smooth = spline_cl(aoa_smooth)
            cd_smooth = spline_cd(aoa_smooth)

            fig = Figure(figsize=(7.8, 4.6), facecolor=self.bg_card)
            ax = fig.add_subplot(111, facecolor=self.bg_dark)

            ax.plot(aoa_smooth, cl_smooth, color=self.fg_blue, linewidth=2.5, label='$C_L$ Curve')
            ax.plot(aoa_smooth, cd_smooth, color=self.fg_green, linewidth=2.5, label='$C_D$ Curve')
            ax.scatter(aoa, cl, color=self.bg_dark, edgecolors=self.fg_blue, s=40, zorder=5)
            ax.scatter(aoa, cd, color=self.bg_dark, edgecolors=self.fg_green, s=40, zorder=5)

            ax.set_xlabel("Angle of Attack (AOA°)", color=self.fg_white, fontweight='bold')
            ax.set_ylabel("Properties Dimensionless Scale", color=self.fg_white)
            ax.set_title("Sail Characteristic Curve", color=self.fg_green,
                         fontweight='bold', pad=12)
            ax.grid(True, color=self.grid_color, linestyle=':')
            ax.tick_params(colors=self.fg_white)
            ax.legend(facecolor=self.bg_card, edgecolor=self.fg_green, labelcolor=self.fg_white)
            fig.tight_layout()

            canvas = FigureCanvasTkAgg(fig, master=self.sail_plot_box)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)
            if hasattr(self, "_sail_grid_entries_status"):
                self._sail_grid_entries_status.config(
                    text=f"Rows: {len(df)} | AOA range: {df.iloc[:, 0].min():g} to {df.iloc[:, 0].max():g} deg",
                    fg=self.fg_blue)
        except Exception as e:
            tk.Label(self.sail_plot_box, text=f"Pipeline error connecting to asset: {e}", fg='red',
                     bg=self.bg_dark).pack(pady=20)
            if hasattr(self, "_sail_grid_entries_status"):
                self._sail_grid_entries_status.config(text=f"Sail curve error: {e}", fg='red')

    # =========================================================================
    # PROPELLER OPEN WATER
    # =========================================================================
    def setup_propeller_ui_layout(self):
        top = tk.Frame(self.frame_prop, bg=self.bg_dark)
        top.pack(side='top', fill='x', padx=15, pady=5)
        tk.Label(top, text="DATA LAYER INTERFACE: propeller performance benchmarks", bg=self.bg_dark, fg=self.fg_blue,
                 font=('Segoe UI', 9, 'italic')).pack(side='left')

        body = tk.Frame(self.frame_prop, bg=self.bg_dark)
        body.pack(fill='both', expand=True, padx=15, pady=10)
        body.grid_columnconfigure(0, minsize=480, weight=0)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        self.build_grid_table_panel(
            body,
            " EDITABLE PROPELLER TABLE ",
            "Edit propeller data below. Columns: J, KT, 10KQ, Eta_o.",
            ["J", "KT", "10KQ", "Eta_o"],
            "prop_grid_entries",
            self.refresh_propeller_plot,
            lambda: self.load_propeller_grid_from_excel(silent=False),
            self.save_propeller_table_to_excel
        )
        self.prop_plot_box = tk.Frame(body, bg=self.bg_dark)
        self.prop_plot_box.grid(row=0, column=1, sticky='nsew')
        self.load_propeller_grid_from_excel(silent=True)

    def load_propeller_dataframe(self):
        try:
            df = self.load_data(self.prop_file)
            if df.shape[1] < 4:
                raise ValueError("propeller open water curve.xlsx must contain J, KT, 10KQ and Eta columns.")
            prop_df = pd.DataFrame({
                "J": pd.to_numeric(df.iloc[:, 0], errors='coerce'),
                "KT": pd.to_numeric(df.iloc[:, 1], errors='coerce'),
                "10KQ": pd.to_numeric(df.iloc[:, 2], errors='coerce'),
                "Eta_o": pd.to_numeric(df.iloc[:, 3], errors='coerce')
            }).dropna()
            if prop_df.empty:
                raise ValueError("No numeric propeller rows found.")
            return prop_df.sort_values("J")
        except Exception:
            return pd.DataFrame({
                "J": [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.88, 0.92],
                "KT": [0.435, 0.395, 0.355, 0.312, 0.268, 0.222, 0.174, 0.123, 0.07, 0.024, 0],
                "10KQ": [0.65, 0.595, 0.54, 0.482, 0.422, 0.362, 0.3, 0.235, 0.165, 0.105, 0.075],
                "Eta_o": [0, 0.106, 0.209, 0.31, 0.404, 0.488, 0.554, 0.584, 0.54, 0.322, 0]
            })

    def load_propeller_grid_from_excel(self, silent=False):
        try:
            df = self.load_propeller_dataframe()
            self._grid_populate("prop_grid_entries", df)
            self.refresh_propeller_plot()
            if not silent:
                messagebox.showinfo("Propeller Table", "Propeller table reloaded from Excel.")
        except Exception as e:
            if hasattr(self, "_prop_grid_entries_status"):
                self._prop_grid_entries_status.config(text=f"Propeller load failed: {e}", fg='red')
            if not silent:
                messagebox.showerror("Propeller Load Failed", str(e))

    def parse_propeller_table(self):
        if not hasattr(self, "prop_grid_entries") or not self.prop_grid_entries:
            return self.load_propeller_dataframe()
        df = self._grid_to_dataframe("prop_grid_entries", ["J", "KT", "10KQ", "Eta_o"])
        return df.groupby("J", as_index=False).mean().sort_values("J")

    def save_propeller_table_to_excel(self):
        try:
            df = self.parse_propeller_table()
            out_df = pd.DataFrame({
                "Advance Coefficient (J)": df["J"],
                "Thrust Coefficient (KT)": df["KT"],
                "Torque Coefficient (10KQ)": df["10KQ"],
                "Open Water Efficiency (Eta_o)": df["Eta_o"]
            })
            out_df.to_excel(self.prop_file, index=False)
            messagebox.showinfo("Propeller Table Saved", f"Saved {len(df)} rows to {self.prop_file}.")
        except Exception as e:
            messagebox.showerror("Propeller Save Failed", str(e))

    def build_propeller_panel(self):
        self.refresh_propeller_plot()

    def refresh_propeller_plot(self):
        for w in self.prop_plot_box.winfo_children(): w.destroy()
        try:
            df = self.parse_propeller_table()
            j = df["J"].to_numpy(dtype=float)
            kt = df["KT"].to_numpy(dtype=float)
            kq10 = df["10KQ"].to_numpy(dtype=float)
            eta = df["Eta_o"].to_numpy(dtype=float)

            j_smooth = np.linspace(j.min(), j.max(), 300)
            spline_order = min(3, len(j) - 1)
            spline_kt = make_interp_spline(j, kt, k=spline_order)
            spline_kq10 = make_interp_spline(j, kq10, k=spline_order)
            spline_eta = make_interp_spline(j, eta, k=spline_order)

            kt_smooth = spline_kt(j_smooth)
            kq10_smooth = spline_kq10(j_smooth)
            eta_smooth = spline_eta(j_smooth)

            fig = Figure(figsize=(7.8, 4.6), facecolor=self.bg_card)
            ax1 = fig.add_subplot(111, facecolor=self.bg_dark)

            ax1.plot(j_smooth, kt_smooth, color=self.fg_green, linewidth=2.5, label='$K_T$')
            ax1.plot(j_smooth, kq10_smooth, color=self.plot_alt, linewidth=2.5, label='$10K_Q$')
            ax1.scatter(j, kt, color=self.bg_dark, edgecolors=self.fg_green, s=40, zorder=5)
            ax1.scatter(j, kq10, color=self.bg_dark, edgecolors=self.plot_alt, s=40, zorder=5)

            ax1.set_xlabel("Advance Coefficient (J)", color=self.fg_white, fontweight='bold')
            ax1.set_ylabel("Force Performance Vector Scale", color=self.fg_white)
            ax1.grid(True, color=self.grid_color, linestyle=':')
            ax1.tick_params(colors=self.fg_white)

            ax2 = ax1.twinx()
            ax2.plot(j_smooth, eta_smooth, color=self.fg_blue, linewidth=2.5, label=r'$\eta_o$')
            ax2.scatter(j, eta, color=self.bg_dark, edgecolors=self.fg_blue, s=40, zorder=5)
            ax2.set_ylabel(r"Open Water Efficiency Profile ($\eta_o$)", color=self.fg_blue, fontweight='bold')
            ax2.tick_params(colors=self.fg_blue)

            fig.tight_layout()
            lines1, labels1 = ax1.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            ax1.legend(lines1 + lines2, labels1 + labels2, loc='center left', facecolor=self.bg_card,
                       edgecolor=self.fg_blue, labelcolor=self.fg_white)

            canvas = FigureCanvasTkAgg(fig, master=self.prop_plot_box)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)
            if hasattr(self, "_prop_grid_entries_status"):
                self._prop_grid_entries_status.config(text=f"Rows: {len(df)} | J range: {df['J'].min():g} to {df['J'].max():g}",
                                        fg=self.fg_blue)
        except Exception as e:
            tk.Label(self.prop_plot_box, text=f"Pipeline error connecting to excel file: {e}", fg='red',
                     bg=self.bg_dark).pack(pady=20)
            if hasattr(self, "_prop_grid_entries_status"):
                self._prop_grid_entries_status.config(text=f"Propeller curve error: {e}", fg='red')

    # =========================================================================
    # RESISTANCE CURVE
    # =========================================================================
    def setup_resistance_ui_layout(self):
        top = tk.Frame(self.frame_resistance, bg=self.bg_dark)
        top.pack(side='top', fill='x', padx=15, pady=5)
        tk.Label(top, text="DATA LAYER INTERFACE: ship speed vs total resistance (columns A and E)",
                 bg=self.bg_dark, fg=self.fg_blue, font=('Segoe UI', 9, 'italic')).pack(side='left')

        body = tk.Frame(self.frame_resistance, bg=self.bg_dark)
        body.pack(fill='both', expand=True, padx=15, pady=10)
        body.grid_columnconfigure(0, minsize=480, weight=0)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        self.build_grid_table_panel(
            body,
            " EDITABLE RESISTANCE TABLE ",
            "Edit resistance data below. Columns: Speed (knots), RT (kN).",
            ["Speed (knots)", "RT (kN)"],
            "resistance_grid_entries",
            self.refresh_resistance_plot,
            lambda: self.load_resistance_grid_from_excel(silent=False),
            self.save_resistance_table_to_excel
        )
        self.resistance_plot_box = tk.Frame(body, bg=self.bg_dark)
        self.resistance_plot_box.grid(row=0, column=1, sticky='nsew')
        self.load_resistance_grid_from_excel(silent=True)

    def load_resistance_dataframe(self):
        try:
            df = self.load_data(self.resistance_file)
            if df.shape[1] < 5:
                raise ValueError("ship total resistance.xlsx must contain at least columns A through E.")

            resistance_df = pd.DataFrame({
                "Speed (knots)": pd.to_numeric(df.iloc[:, 0], errors='coerce'),
                "Total Resistance (kN)": pd.to_numeric(df.iloc[:, 4], errors='coerce')
            }).dropna()

            if resistance_df.empty:
                raise ValueError("No numeric speed/resistance rows found in columns A and E.")

            resistance_df = resistance_df.groupby("Speed (knots)", as_index=False).mean()
            resistance_df = resistance_df.sort_values("Speed (knots)")
            return resistance_df
        except Exception:
            return pd.DataFrame({
                "Speed (knots)": [0, 2, 4, 6, 8, 10, 12, 14, 16, 18],
                "Total Resistance (kN)": [0, 15, 45, 80, 125, 135, 240, 300, 520, 760]
            })

    def load_resistance_grid_from_excel(self, silent=False):
        try:
            df = self.load_resistance_dataframe()
            self._grid_populate("resistance_grid_entries", df)
            self.refresh_resistance_plot(show_popup=False)
            if not silent:
                messagebox.showinfo("Resistance Table", "Resistance table reloaded from Excel columns A and E.")
        except Exception as e:
            if hasattr(self, "_resistance_grid_entries_status"):
                self._resistance_grid_entries_status.config(text=f"Resistance load failed: {e}", fg='red')
            if not silent:
                messagebox.showerror("Resistance Load Failed", str(e))

    def parse_resistance_table(self):
        if not hasattr(self, "resistance_grid_entries") or not self.resistance_grid_entries:
            return self.load_resistance_dataframe()

        df = self._grid_to_dataframe("resistance_grid_entries", ["Speed (knots)", "Total Resistance (kN)"])
        df = df.groupby("Speed (knots)", as_index=False).mean()
        df = df.sort_values("Speed (knots)")
        return df

    def get_total_resistance(self, vs_knots):
        df = self.parse_resistance_table()
        speed = df["Speed (knots)"].to_numpy(dtype=float)
        resistance_kn = df["Total Resistance (kN)"].to_numpy(dtype=float)

        if len(speed) == 1:
            return resistance_kn[0] * 1000.0

        resistance_curve = interp1d(speed, resistance_kn, fill_value='extrapolate')
        return float(resistance_curve(vs_knots)) * 1000.0

    def refresh_resistance_plot(self, show_popup=False):
        for w in self.resistance_plot_box.winfo_children():
            w.destroy()
        try:
            df = self.parse_resistance_table()
            speed = df["Speed (knots)"].to_numpy(dtype=float)
            resistance_kn = df["Total Resistance (kN)"].to_numpy(dtype=float)

            fig = Figure(figsize=(7.8, 4.6), facecolor=self.bg_card)
            ax = fig.add_subplot(111, facecolor=self.bg_dark)
            ax.plot(speed, resistance_kn, color=self.fg_green, linewidth=2.5, marker='o',
                    markerfacecolor=self.bg_dark, markeredgecolor=self.fg_green, label='RT curve')

            # Only show operating point if a case has been run
            if self.case_has_run:
                try:
                    vs_knots = float(self.entries["Desired Ship Speed, vs (knots)"].get())
                    rt_kn = self.get_total_resistance(vs_knots) / 1000.0
                    case_str = f"Case {self.active_case_no}: " if hasattr(self, 'active_case_no') else ""
                    ax.scatter([vs_knots], [rt_kn], color='#FFFF00', edgecolors='black', s=120,
                               zorder=10, label=f'{case_str}Input speed RT = {rt_kn:.3f} kN')
                except Exception:
                    pass

            ax.set_xlabel("Ship Speed (knots)", color=self.fg_white, fontweight='bold')
            ax.set_ylabel("Total Resistance RT (kN)", color=self.fg_white, fontweight='bold')
            ax.set_title("Ship Total Resistance Curve", color=self.fg_green, fontweight='bold', pad=12)
            ax.grid(True, color=self.grid_color, linestyle=':')
            ax.tick_params(colors=self.fg_white)
            ax.legend(facecolor=self.bg_card, edgecolor=self.fg_green, labelcolor=self.fg_white)
            fig.tight_layout()

            canvas = FigureCanvasTkAgg(fig, master=self.resistance_plot_box)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)

            status_text = f"Rows: {len(df)}"
            if self.case_has_run:
                try:
                    vs_knots = float(self.entries["Desired Ship Speed, vs (knots)"].get())
                    rt_kn = self.get_total_resistance(vs_knots) / 1000.0
                    status_text += f" | Input speed: {vs_knots:g} knots | Interpolated RT: {rt_kn:.3f} kN"
                except Exception:
                    pass

            if hasattr(self, "_resistance_grid_entries_status"):
                self._resistance_grid_entries_status.config(text=status_text, fg=self.fg_blue)

            if show_popup:
                messagebox.showinfo("Resistance Curve", "Resistance table applied.")
        except Exception as e:
            tk.Label(self.resistance_plot_box, text=f"Resistance curve error: {e}", fg='red',
                     bg=self.bg_dark).pack(pady=20)
            if hasattr(self, "_resistance_grid_entries_status"):
                self._resistance_grid_entries_status.config(text=f"Resistance curve error: {e}", fg='red')

    def save_resistance_table_to_excel(self):
        try:
            df = self.parse_resistance_table()
            out_df = pd.DataFrame({
                "Speed (V in knots)": df["Speed (knots)"],
                "Total Resistance without WAP (RT in kN)": df["Total Resistance (kN)"],
                "Sail Resistance (kN)": np.zeros(len(df)),
                "Rudder resistance": np.zeros(len(df)),
                "Total resistance with WAP (kN)": df["Total Resistance (kN)"]
            })
            out_df.to_excel(self.resistance_file, index=False)
            messagebox.showinfo("Resistance Table Saved", f"Saved {len(df)} rows to {self.resistance_file}.")
        except Exception as e:
            messagebox.showerror("Resistance Save Failed", str(e))

    # =========================================================================
    # AOA OPTIMISER TAB (6th tab — separate sub-window)
    # =========================================================================
    def setup_optimization_ui_layout(self):
        dash_strip = tk.Frame(self.frame_opt, bg=self.bg_card, bd=1, relief='solid')
        dash_strip.pack(side='top', fill='x', padx=15, pady=8)

        btn_solve = tk.Button(dash_strip, text="Run optimisation", command=self.compute_3d_convergence,
                              bg=self.bg_dark, fg=self.fg_green, font=('Segoe UI', 10, 'bold'),
                              bd=1, relief='solid', activebackground=self.fg_green, activeforeground=self.bg_dark)
        btn_solve.pack(side='left', padx=15, pady=10)

        self.hud_summary = tk.Label(dash_strip, text="System Ready. Run a voyage case to view results.",
                                    bg=self.bg_card, fg=self.fg_blue, font=('Consolas', 10, 'bold'))
        self.hud_summary.pack(side='left', padx=15)
        self.opt_plot_box.pack(fill='both', expand=True, padx=15, pady=10)

        # 3D interaction controls below graph
        self.graph_controls = tk.Frame(self.frame_opt, bg=self.bg_dark)
        self.graph_controls.pack(side='bottom', fill='x', pady=8)

        tk.Label(self.graph_controls, text="3D view", bg=self.bg_dark, fg=self.fg_white,
                 font=('Segoe UI', 9, 'bold')).pack(side='left', padx=10)

        tk.Button(self.graph_controls, text="Zoom in",
                  command=lambda: self.trigger_hud_zoom(0.75),
                  bg=self.bg_card, fg=self.fg_blue,
                  font=('Segoe UI', 9, 'bold')).pack(side='left', padx=4)

        tk.Button(self.graph_controls, text="Zoom out",
                  command=lambda: self.trigger_hud_zoom(1.3),
                  bg=self.bg_card, fg=self.fg_blue,
                  font=('Segoe UI', 9, 'bold')).pack(side='left', padx=4)

        tk.Button(self.graph_controls, text="Reset view",
                  command=self.reset_hud_view,
                  bg=self.bg_card, fg=self.fg_green,
                  font=('Segoe UI', 9, 'bold')).pack(side='left', padx=4)

    def trigger_hud_zoom(self, factor):
        if self.ax_3d is None or self.canvas_3d is None: return
        xlim, ylim, zlim = self.ax_3d.get_xlim3d(), self.ax_3d.get_ylim3d(), self.ax_3d.get_zlim3d()
        x_m, y_m, z_m = sum(xlim) / 2, sum(ylim) / 2, sum(zlim) / 2
        x_s, y_s, z_s = (xlim[1] - xlim[0]) * factor, (ylim[1] - ylim[0]) * factor, (zlim[1] - zlim[0]) * factor

        self.ax_3d.set_xlim3d([x_m - x_s / 2, x_m + x_s / 2])
        self.ax_3d.set_ylim3d([y_m - y_s / 2, y_m + y_s / 2])
        self.ax_3d.set_zlim3d([z_m - z_s / 2, z_m + z_s / 2])
        self.canvas_3d.draw()

    def reset_hud_view(self):
        if self.ax_3d is None or self.canvas_3d is None or self.orig_xlim is None: return
        self.ax_3d.set_xlim3d(self.orig_xlim)
        self.ax_3d.set_ylim3d(self.orig_ylim)
        self.ax_3d.set_zlim3d(self.orig_zlim)
        self.ax_3d.view_init(elev=22, azim=-45)
        self.canvas_3d.draw()


    # =============================================================================
    # FUEL SAVING CALCULATOR TAB
    # =============================================================================
    def safe_float(self, value, default=0.0):
        try:
            if value is None or pd.isna(value):
                return default
            return float(value)
        except Exception:
            return default

    def load_voyage_settings(self):
        """Read global voyage settings such as duration, SFOC, and efficiencies.
        Now also checks the in-UI SFOC / duration entries first.
        """
        # Prefer UI values if available
        sfoc_ui = None
        total_hours_ui = None
        if hasattr(self, "sfoc_entry"):
            try:
                sfoc_ui = float(self.sfoc_entry.get())
            except (ValueError, tk.TclError):
                pass
        if hasattr(self, "duration_entry"):
            try:
                dur_val = float(self.duration_entry.get())
                if hasattr(self, "duration_unit_var"):
                    unit = self.duration_unit_var.get()
                    total_hours_ui = dur_val * 24.0 if unit == "days" else dur_val
                else:
                    total_hours_ui = dur_val
            except (ValueError, tk.TclError):
                pass

        if not os.path.exists(self.voyage_file):
            return {
                "total_hours": total_hours_ui or 0.0,
                "sfoc": sfoc_ui or 150.0,
                "diameter": float(self.entries["Propeller Diameter, D (m)"].get()),
                "eta_without": 1.0,
                "transmission_eff": 0.9,
            }

        df = pd.read_excel(self.voyage_file, header=None)
        total_days = self.safe_float(df.iat[0, 2] if df.shape[0] > 0 and df.shape[1] > 2 else 0.0)
        total_hours = self.safe_float(df.iat[0, 3] if df.shape[0] > 0 and df.shape[1] > 3 else total_days * 24.0)
        if total_hours <= 0:
            total_hours = total_days * 24.0

        sfoc = self.safe_float(df.iat[1, 2] if df.shape[0] > 1 and df.shape[1] > 2 else 150.0, 150.0)
        diameter = self.safe_float(df.iat[2, 2] if df.shape[0] > 2 and df.shape[1] > 2 else self.entries["Propeller Diameter, D (m)"].get(), 4.0)
        eta_without = self.safe_float(df.iat[3, 7] if df.shape[0] > 3 and df.shape[1] > 7 else 1.0, 1.0)
        transmission_eff = self.safe_float(df.iat[4, 7] if df.shape[0] > 4 and df.shape[1] > 7 else 0.9, 0.9)

        if eta_without <= 0:
            eta_h = self.safe_float(df.iat[0, 7] if df.shape[1] > 7 else 1.0, 1.0)
            eta_r = self.safe_float(df.iat[1, 7] if df.shape[1] > 7 else 1.0, 1.0)
            eta_o = self.safe_float(df.iat[2, 7] if df.shape[1] > 7 else 1.0, 1.0)
            eta_without = max(eta_h * eta_r * eta_o, 1e-9)

        if transmission_eff <= 0:
            transmission_eff = 1.0

        # Override with UI values if present
        if sfoc_ui is not None:
            sfoc = sfoc_ui
        if total_hours_ui is not None:
            total_hours = total_hours_ui

        return {
            "total_hours": total_hours,
            "sfoc": sfoc,
            "diameter": diameter,
            "eta_without": eta_without,
            "transmission_eff": transmission_eff,
        }

    def load_voyage_cases(self):
        """Read voyage case rows from the voyage details workbook."""
        if not os.path.exists(self.voyage_file):
            return [
                {
                    "case_no": 1,
                    "wind_speed": 15.0,
                    "wind_angle": 60.0,
                    "ship_speed": 12.0,
                    "percentage": 5.0,
                },
                {
                    "case_no": 2,
                    "wind_speed": 20.0,
                    "wind_angle": 50.0,
                    "ship_speed": 14.0,
                    "percentage": 10.0,
                },
                {
                    "case_no": 3,
                    "wind_speed": 21.0,
                    "wind_angle": 40.0,
                    "ship_speed": 15.0,
                    "percentage": 25.0,
                },
                {
                    "case_no": 4,
                    "wind_speed": 25.0,
                    "wind_angle": 20.0,
                    "ship_speed": 15.0,
                    "percentage": 40.0,
                },
                {
                    "case_no": 5,
                    "wind_speed": 22.0,
                    "wind_angle": 30.0,
                    "ship_speed": 15.0,
                    "percentage": 20.0,
                }
            ]

        raw = pd.read_excel(self.voyage_file, header=None)
        header_row = None

        for idx in range(len(raw)):
            row_values = [str(x).strip().lower() for x in raw.iloc[idx, :4].tolist()]
            if (
                len(row_values) >= 4
                and "true wind speed" in row_values[0]
                and "true wind angle" in row_values[1]
                and "ship speed" in row_values[2]
                and "percentage" in row_values[3]
            ):
                header_row = idx
                break

        if header_row is None:
            raise ValueError("Could not find the voyage case header row in voyage details.xlsx.")

        cases = []
        for row_idx in range(header_row + 1, len(raw)):
            wind_speed = self.safe_float(raw.iat[row_idx, 0], None)
            wind_angle = self.safe_float(raw.iat[row_idx, 1], None)
            ship_speed = self.safe_float(raw.iat[row_idx, 2], None)
            percentage = self.safe_float(raw.iat[row_idx, 3], None)

            if wind_speed is None or wind_angle is None or ship_speed is None or percentage is None:
                row_text = " ".join(str(v).lower() for v in raw.iloc[row_idx, :12].tolist())
                if "total voyage" in row_text:
                    break
                if cases:
                    break
                continue

            cases.append({
                "case_no": len(cases) + 1,
                "wind_speed": float(wind_speed),
                "wind_angle": float(wind_angle),
                "ship_speed": float(ship_speed),
                "percentage": float(percentage),
            })

        if not cases:
            raise ValueError("No numeric voyage cases found. Expected columns: true wind speed, angle, ship speed, percentage duration.")

        return cases

    def calculate_case_solution(self, case):
        """Calculate one voyage case and return all values displayed in the UI."""
        self.set_case_inputs(case)

        vw = float(case["wind_speed"])
        tw_deg = float(case["wind_angle"])
        vs_knots = float(case["ship_speed"])

        w = float(self.entries["Wake Fraction, w"].get())
        t = float(self.entries["Thrust Deduction Fraction, t"].get())
        rho_a = float(self.entries["Density of Air, rho_a (kg/m3)"].get())
        rho_w = float(self.entries["Density of Water, rho_w (kg/m3)"].get())
        As = float(self.entries["Sail Area, As (m2)"].get())
        D = float(self.entries["Propeller Diameter, D (m)"].get())
        eta_r = float(self.entries["Relative Rotative Efficiency, eta_R"].get())

        vs_ms = vs_knots * 0.5144
        tw_rad = np.radians(tw_deg)
        vx = -vs_ms + vw * np.cos(tw_rad)
        vy = vw * np.sin(tw_rad)
        va = np.sqrt(vx ** 2 + vy ** 2)
        theta_deg = np.degrees(np.arctan2(vy, vx))
        theta_rad = np.radians(theta_deg)
        vadv = vs_ms * (1 - w)
        eta_h = (1 - t) / (1 - w)

        df_sail = self.parse_sail_table()
        aoa_raw = df_sail["AOA (deg)"].to_numpy(dtype=float)
        cl_raw = df_sail["Lift Coefficient (CL)"].to_numpy(dtype=float)
        cd_raw = df_sail["Drag Coefficient (CD)"].to_numpy(dtype=float)

        if np.max(aoa_raw) < 180.0:
            aoa_raw = np.append(aoa_raw, 180.0)
            cl_raw = np.append(cl_raw, 0.0)
            cd_raw = np.append(cd_raw, cd_raw[-1])

        df_prop = self.parse_propeller_table()
        j_raw = df_prop["J"].to_numpy(dtype=float)
        kt_raw = df_prop["KT"].to_numpy(dtype=float)
        eta_raw = df_prop["Eta_o"].to_numpy(dtype=float)

        cl_spline = make_interp_spline(aoa_raw, cl_raw, k=min(3, len(aoa_raw) - 1))
        cd_spline = make_interp_spline(aoa_raw, cd_raw, k=min(3, len(aoa_raw) - 1))
        kt_spline = make_interp_spline(j_raw, kt_raw, k=min(3, len(j_raw) - 1))
        eta_spline = make_interp_spline(j_raw, eta_raw, k=min(3, len(j_raw) - 1))

        RT = self.get_total_resistance(vs_knots)
        aoa_dense = np.linspace(0, 180, 50)
        j_scanner = np.linspace(0.01, np.max(j_raw), 1000)
        kt_prop_scanned = kt_spline(j_scanner)

        result_rows = []

        for angle in aoa_dense:
            lift = 0.5 * rho_a * As * (va ** 2) * float(cl_spline(angle))
            drag = 0.5 * rho_a * As * (va ** 2) * float(cd_spline(angle))
            t_sail = lift * np.sin(theta_rad) - drag * np.cos(theta_rad)
            t_prop = (RT - t_sail) / (1 - t)

            kt_j2_ship = t_prop / (rho_w * (D ** 2) * (vadv ** 2))
            kt_ship_scanned = kt_j2_ship * (j_scanner ** 2)
            match_idx = int(np.argmin(np.abs(kt_prop_scanned - kt_ship_scanned)))

            current_j = float(j_scanner[match_idx])
            current_kt = float(kt_prop_scanned[match_idx])
            eta_o = float(eta_spline(current_j))
            eta_d = eta_o * eta_h * eta_r
            PE = (RT - t_sail) * vs_ms
            PD = PE / eta_d if eta_d > 0 else np.nan

            result_rows.append({
                "aoa": float(angle),
                "j": current_j,
                "kt": current_kt,
                "eta_d": float(eta_d),
                "pd_w": float(PD) if not np.isnan(PD) else np.nan,
                "rt_n": float(RT),
                "t_sail_n": float(t_sail),
            })

        eta_values = np.array([row["eta_d"] for row in result_rows], dtype=float)
        if np.all(np.isnan(eta_values)):
            raise ValueError("Could not calculate a valid propulsive efficiency for this case.")

        best_idx = int(np.nanargmax(eta_values))
        best = result_rows[best_idx]

        settings = self.load_voyage_settings()
        total_hours = settings["total_hours"]
        sfoc = settings["sfoc"]
        transmission_eff = settings["transmission_eff"]
        eta_without = settings["eta_without"]

        case_hours = total_hours * float(case["percentage"]) / 100.0

        delivered_power_with_kw = best["pd_w"] / 1000.0
        fuel_with_sail_t = (delivered_power_with_kw / transmission_eff) * sfoc * case_hours / 1_000_000.0

        effective_power_without_kw = (best["rt_n"] * vs_ms) / 1000.0
        delivered_power_without_kw = effective_power_without_kw / eta_without
        fuel_without_sail_t = (delivered_power_without_kw / transmission_eff) * sfoc * case_hours / 1_000_000.0

        fuel_saving_t = fuel_without_sail_t - fuel_with_sail_t
        percent_fuel_saving = (fuel_saving_t / fuel_without_sail_t * 100.0) if fuel_without_sail_t else 0.0

        # This follows the existing voyage-details workbook rpm calculation style.
        rpm = (vs_ms / best["j"] / D) / (2 * np.pi) * 60 if best["j"] > 0 and D > 0 else np.nan

        return {
            **case,
            "aoa": best["aoa"],
            "j": best["j"],
            "eta_d": best["eta_d"],
            "pd_kw": delivered_power_with_kw,
            "rt_kn": best["rt_n"] / 1000.0,
            "propeller_rpm": rpm,
            "fuel_required_with_sail_t": fuel_with_sail_t,
            "fuel_required_without_sail_t": fuel_without_sail_t,
            "fuel_saving_t": fuel_saving_t,
            "percent_fuel_saving": percent_fuel_saving,
        }

    def setup_fuel_saving_ui_layout(self):
        """Build the primary Fuel Saving Calculator tab.

        The tab is intentionally arranged from top to bottom:
        1. Header actions with SFOC/duration inputs
        2. Voyage case table (editable)
        3. Final voyage summary cards
        """
        # Header / action bar -------------------------------------------------
        header = tk.Frame(self.frame_fuel, bg=self.bg_card, bd=0,
                          highlightthickness=1, highlightbackground=self.border_color)
        header.pack(fill='x', padx=0, pady=(0, 10))

        header_text = tk.Frame(header, bg=self.bg_card)
        header_text.pack(side='left', fill='x', expand=True, padx=16, pady=12)
        self.make_label(header_text, "Fuel saving calculator", role="title").pack(anchor='w')
        self.make_label(header_text, "Run individual voyage cases and view the AOA optimisation graph in the AOA Optimiser tab.",
                        role="muted").pack(anchor='w', pady=(2, 0))

        action_box = tk.Frame(header, bg=self.bg_card)
        action_box.pack(side='right', padx=16, pady=12)
        self.make_button(action_box, "Reload voyage details", command=self.refresh_fuel_window,
                         variant="secondary").pack(side='right', padx=(8, 0))
        self.make_button(action_box, "Calculate all cases", command=self.calculate_all_fuel_cases,
                         variant="primary").pack(side='right')

        # SFOC and Duration input bar -----------------------------------------
        param_bar = tk.Frame(self.frame_fuel, bg=self.bg_card, bd=0,
                             highlightthickness=1, highlightbackground=self.border_color)
        param_bar.pack(fill='x', padx=0, pady=(0, 10))

        # SFOC
        sfoc_frame = tk.Frame(param_bar, bg=self.bg_card)
        sfoc_frame.pack(side='left', padx=16, pady=8)
        self.make_label(sfoc_frame, "SFOC (g/kW·hr):", role="body").pack(side='left', padx=(0, 6))
        self.sfoc_entry = tk.Entry(sfoc_frame, bg=self.input_bg, fg=self.fg_green,
                                   insertbackground=self.fg_green,
                                   font=('Consolas', 10, 'bold'), bd=1, relief='solid', width=10)
        self.sfoc_entry.insert(0, "150.0")
        self.sfoc_entry.pack(side='left')

        # Duration
        dur_frame = tk.Frame(param_bar, bg=self.bg_card)
        dur_frame.pack(side='left', padx=16, pady=8)
        self.make_label(dur_frame, "Total voyage duration:", role="body").pack(side='left', padx=(0, 6))
        self.duration_entry = tk.Entry(dur_frame, bg=self.input_bg, fg=self.fg_green,
                                       insertbackground=self.fg_green,
                                       font=('Consolas', 10, 'bold'), bd=1, relief='solid', width=10)
        self.duration_entry.insert(0, "0")
        self.duration_entry.pack(side='left', padx=(0, 6))

        self.duration_unit_var = tk.StringVar(value="hours")
        unit_frame = tk.Frame(dur_frame, bg=self.bg_card)
        unit_frame.pack(side='left')
        for text, val in [("Hours", "hours"), ("Days", "days")]:
            rb = tk.Radiobutton(unit_frame, text=text, value=val, variable=self.duration_unit_var,
                                bg=self.bg_card, fg=self.fg_white, selectcolor=self.input_bg,
                                activebackground=self.bg_card, activeforeground=self.fg_green,
                                font=('Segoe UI', 9), anchor='w')
            rb.pack(side='left', padx=2)

        # Try loading SFOC and duration from Excel
        try:
            settings = self.load_voyage_settings()
            self.sfoc_entry.delete(0, tk.END)
            self.sfoc_entry.insert(0, f"{settings['sfoc']:g}")
            if settings['total_hours'] > 0:
                self.duration_entry.delete(0, tk.END)
                self.duration_entry.insert(0, f"{settings['total_hours']:g}")
        except Exception:
            pass

        self.fuel_status = self.make_label(self.frame_fuel, "", role="muted", bg=self.bg_dark,
                                           font=("Consolas", 10, "bold"), justify='left')
        self.fuel_status.pack(anchor='w', padx=2, pady=(0, 8))

        # Main grid ----------------------------------------------------------
        main = tk.Frame(self.frame_fuel, bg=self.bg_dark)
        main.pack(fill='both', expand=True)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=0)

        # Voyage cases table (EDITABLE) --------------------------------------
        table_holder = self.make_section(main, " VOYAGE CASES ", padx=10, pady=10)
        table_holder.grid(row=0, column=0, sticky='nsew')
        table_holder.grid_rowconfigure(0, weight=1)
        table_holder.grid_columnconfigure(0, weight=1)

        self.fuel_table_canvas = tk.Canvas(table_holder, bg=self.bg_card, highlightthickness=0)
        y_scroll = tk.Scrollbar(table_holder, orient='vertical', command=self.fuel_table_canvas.yview)
        x_scroll = tk.Scrollbar(table_holder, orient='horizontal', command=self.fuel_table_canvas.xview)
        self.fuel_table_canvas.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.fuel_table_canvas.grid(row=0, column=0, sticky='nsew')
        y_scroll.grid(row=0, column=1, sticky='ns')
        x_scroll.grid(row=1, column=0, sticky='ew')

        self.fuel_table_frame = tk.Frame(self.fuel_table_canvas, bg=self.bg_card)
        self.fuel_table_window_id = self.fuel_table_canvas.create_window((0, 0), window=self.fuel_table_frame, anchor='nw')
        self.fuel_table_frame.bind(
            "<Configure>",
            lambda event: self.fuel_table_canvas.configure(scrollregion=self.fuel_table_canvas.bbox("all"))
        )

        def resize_fuel_table_canvas(event):
            """Keep the table at least as wide as its canvas, but allow overflow scroll."""
            try:
                required_width = self.fuel_table_frame.winfo_reqwidth()
                self.fuel_table_canvas.itemconfig(self.fuel_table_window_id, width=max(event.width, required_width))
            except Exception:
                pass

        self.fuel_table_canvas.bind("<Configure>", resize_fuel_table_canvas)

        # Add row / Sync to Excel controls below table
        table_controls = tk.Frame(table_holder, bg=self.bg_card)
        table_controls.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(6, 0))
        self.make_button(table_controls, "+ Add Row", command=self.add_voyage_row,
                         variant="secondary", padx=10, pady=5).pack(side='left', padx=4)
        self.make_button(table_controls, "− Delete Last Row", command=self.delete_last_voyage_row,
                         variant="danger", padx=10, pady=5).pack(side='left', padx=4)
        self.make_button(table_controls, "Sync to Excel", command=self.sync_voyage_to_excel,
                         variant="primary", padx=10, pady=5).pack(side='right', padx=4)

        # Voyage summary metric cards ---------------------------------------
        self.fuel_summary = self.make_section(main, " FINAL VOYAGE OUTPUT ", padx=10, pady=10)
        self.fuel_summary.grid(row=1, column=0, sticky='ew', pady=(10, 10))
        for col in range(4):
            self.fuel_summary.grid_columnconfigure(col, weight=1)

        self.fuel_summary_labels = {}
        summary_metrics = [
            "Total fuel required with sail",
            "Total fuel required without sail",
            "Total fuel saving",
            "Overall percentage fuel saving",
        ]
        for col, label in enumerate(summary_metrics):
            metric = tk.Frame(self.fuel_summary, bg=self.input_bg, bd=0,
                              highlightthickness=1, highlightbackground=self.border_color)
            metric.grid(row=0, column=col, sticky='nsew', padx=6, pady=2)
            self.make_label(metric, label, role="muted", bg=self.input_bg).pack(anchor='w', padx=10, pady=(8, 0))
            value_label = self.make_label(metric, "--", role="metric", bg=self.input_bg)
            value_label.pack(anchor='w', padx=10, pady=(2, 8))
            self.fuel_summary_labels[label] = value_label

        self.fuel_case_labels = {}
        self.fuel_case_editable_entries = {}  # {case_no: {"wind_speed": Entry, ...}}
        self.fuel_results = {}
        self.refresh_fuel_window()

    def open_fuel_saving_calculator(self):
        # Backwards-compatible wrapper: fuel calculator is now a main-window tab.
        self.show_fuel_tab()

    def refresh_fuel_window(self):
        """Reload voyage cases from Excel and rebuild the case table."""
        self.fuel_case_labels = {}
        self.fuel_case_editable_entries = {}
        self.clear_children(self.fuel_table_frame)

        try:
            self.voyage_cases = self.load_voyage_cases()
            total_percentage = sum(case["percentage"] for case in self.voyage_cases)
            if abs(total_percentage - 100.0) < 1e-6:
                status_text = f"Voyage cases loaded: {len(self.voyage_cases)} | Time allocation: {total_percentage:.2f}%"
                status_color = self.fg_green
            else:
                status_text = f"Voyage cases loaded: {len(self.voyage_cases)} | Time allocation: {total_percentage:.2f}%; expected 100%"
                status_color = self.warning_color

            self.fuel_status.config(text=status_text, fg=status_color)

            headers = [
                "Case", "Wind speed", "Wind angle", "Ship speed", "Time %", "Action",
                "RPM", "Fuel with sail", "Fuel without sail", "Fuel saving", "Saving %"
            ]

            for col, header_text in enumerate(headers):
                cell = self.make_table_cell(self.fuel_table_frame, header_text, header=True, width=14)
                cell.grid(row=0, column=col, sticky='nsew', padx=0, pady=0)

            for row_idx, case in enumerate(self.voyage_cases, start=1):
                alt = row_idx % 2 == 0
                bg = self.table_alt if alt else self.input_bg

                # Column 0: Case number (read-only label)
                self.make_table_cell(self.fuel_table_frame, f"Case {case['case_no']}", alt=alt, width=14).grid(
                    row=row_idx, column=0, sticky='nsew', padx=0, pady=0
                )

                # Columns 1-4: Editable Entry widgets
                editable_keys = ["wind_speed", "wind_angle", "ship_speed", "percentage"]
                editable_values = [
                    f"{case['wind_speed']:g}",
                    f"{case['wind_angle']:g}",
                    f"{case['ship_speed']:g}",
                    f"{case['percentage']:g}",
                ]
                units = [" m/s", "°", " kn", "%"]

                case_entries = {}
                for col_offset, (key, val, unit) in enumerate(zip(editable_keys, editable_values, units)):
                    cell_frame = tk.Frame(self.fuel_table_frame, bg=bg, bd=1, relief='solid')
                    cell_frame.grid(row=row_idx, column=col_offset + 1, sticky='nsew', padx=0, pady=0)

                    ent = tk.Entry(cell_frame, bg=bg, fg=self.fg_white,
                                   insertbackground=self.fg_green,
                                   font=('Consolas', 9), bd=0, relief='flat',
                                   width=8, justify='center')
                    ent.insert(0, val)
                    ent.pack(side='left', fill='both', expand=True, padx=2, pady=4)

                    unit_lbl = tk.Label(cell_frame, text=unit, bg=bg, fg=self.text_muted,
                                        font=('Consolas', 8))
                    unit_lbl.pack(side='right', padx=(0, 4))
                    case_entries[key] = ent

                self.fuel_case_editable_entries[case["case_no"]] = case_entries

                run_button = self.make_button(
                    self.fuel_table_frame,
                    "Run case",
                    command=lambda c=case: self._run_editable_case(c["case_no"]),
                    variant="secondary",
                    padx=8,
                    pady=6,
                )
                run_button.grid(row=row_idx, column=5, sticky='nsew', padx=0, pady=0)

                self.fuel_case_labels[case["case_no"]] = {}
                for col, key in enumerate([
                    "propeller_rpm",
                    "fuel_required_with_sail_t",
                    "fuel_required_without_sail_t",
                    "fuel_saving_t",
                    "percent_fuel_saving",
                ], start=6):
                    lbl = self.make_table_cell(self.fuel_table_frame, "--", alt=alt, width=14, fg=self.fg_blue)
                    lbl.config(font=("Consolas", 9, "bold"))
                    lbl.grid(row=row_idx, column=col, sticky='nsew', padx=0, pady=0)
                    self.fuel_case_labels[case["case_no"]][key] = lbl

            for col in range(len(headers)):
                self.fuel_table_frame.grid_columnconfigure(col, weight=1)

            for existing_result in getattr(self, "fuel_results", {}).values():
                self.update_fuel_row(existing_result)

            self.update_fuel_summary()
            self.highlight_active_case_row()

        except Exception as e:
            self.fuel_status.config(text=f"Fuel calculator load error: {e}", fg=self.danger_color)

    def highlight_active_case_row(self):
        if not hasattr(self, 'active_case_no'): return
        theme = getattr(self, 'theme_name', 'dark')
        active_bg = "#e0f2fe" if theme == 'light' else "#1e3a8a"
        
        for c_no, entries_dict in self.fuel_case_editable_entries.items():
            is_active = (c_no == self.active_case_no)
            row_idx = list(self.fuel_case_editable_entries.keys()).index(c_no) + 1
            default_bg = self.table_alt if row_idx % 2 == 0 else self.input_bg
            bg = active_bg if is_active else default_bg 

            for key, ent in entries_dict.items():
                ent.config(bg=bg)
                if ent.master:
                    ent.master.config(bg=bg)
                    for child in ent.master.winfo_children():
                        if isinstance(child, tk.Label):
                            child.config(bg=bg)

        for c_no, labels_dict in self.fuel_case_labels.items():
            is_active = (c_no == self.active_case_no)
            row_idx = list(self.fuel_case_labels.keys()).index(c_no) + 1
            default_bg = self.table_alt if row_idx % 2 == 0 else self.input_bg
            bg = active_bg if is_active else default_bg 
            for lbl in labels_dict.values():
                lbl.config(bg=bg)

    def _get_case_from_entries(self, case_no):
        """Read a voyage case from the editable Entry widgets."""
        entries = self.fuel_case_editable_entries.get(case_no, {})
        if not entries:
            # Fallback to stored case
            for c in self.voyage_cases:
                if c["case_no"] == case_no:
                    return c
            return None

        return {
            "case_no": case_no,
            "wind_speed": float(entries["wind_speed"].get()),
            "wind_angle": float(entries["wind_angle"].get()),
            "ship_speed": float(entries["ship_speed"].get()),
            "percentage": float(entries["percentage"].get()),
        }

    def _run_editable_case(self, case_no):
        """Run a case using current editable table values."""
        try:
            case = self._get_case_from_entries(case_no)
            if case is None:
                messagebox.showerror("Error", f"Could not find case {case_no}")
                return
            self.run_fuel_case(case, show_popup=False, render_graph=False)
        except Exception as e:
            messagebox.showerror("Fuel saving calculator", f"Case calculation failed:\n{e}")

    def add_voyage_row(self):
        """Add a new editable row to the voyage cases table."""
        # Determine next case number
        existing_cases = list(self.fuel_case_editable_entries.keys())
        next_case_no = max(existing_cases) + 1 if existing_cases else 1

        # Count current rows
        row_idx = next_case_no  # 1-based (row 0 is header)
        alt = row_idx % 2 == 0
        bg = self.table_alt if alt else self.input_bg

        # Case label
        self.make_table_cell(self.fuel_table_frame, f"Case {next_case_no}", alt=alt, width=14).grid(
            row=row_idx, column=0, sticky='nsew', padx=0, pady=0
        )

        # Editable entries
        editable_keys = ["wind_speed", "wind_angle", "ship_speed", "percentage"]
        default_values = ["0", "0", "0", "0"]
        units = [" m/s", "°", " kn", "%"]

        case_entries = {}
        for col_offset, (key, val, unit) in enumerate(zip(editable_keys, default_values, units)):
            cell_frame = tk.Frame(self.fuel_table_frame, bg=bg, bd=1, relief='solid')
            cell_frame.grid(row=row_idx, column=col_offset + 1, sticky='nsew', padx=0, pady=0)

            ent = tk.Entry(cell_frame, bg=bg, fg=self.fg_white,
                           insertbackground=self.fg_green,
                           font=('Consolas', 9), bd=0, relief='flat',
                           width=8, justify='center')
            ent.insert(0, val)
            ent.pack(side='left', fill='both', expand=True, padx=2, pady=4)

            unit_lbl = tk.Label(cell_frame, text=unit, bg=bg, fg=self.text_muted,
                                font=('Consolas', 8))
            unit_lbl.pack(side='right', padx=(0, 4))
            case_entries[key] = ent

        self.fuel_case_editable_entries[next_case_no] = case_entries

        # Run button
        run_button = self.make_button(
            self.fuel_table_frame,
            "Run case",
            command=lambda c=next_case_no: self._run_editable_case(c),
            variant="secondary",
            padx=8,
            pady=6,
        )
        run_button.grid(row=row_idx, column=5, sticky='nsew', padx=0, pady=0)

        # Result labels
        self.fuel_case_labels[next_case_no] = {}
        for col, key in enumerate([
            "propeller_rpm",
            "fuel_required_with_sail_t",
            "fuel_required_without_sail_t",
            "fuel_saving_t",
            "percent_fuel_saving",
        ], start=6):
            lbl = self.make_table_cell(self.fuel_table_frame, "--", alt=alt, width=14, fg=self.fg_blue)
            lbl.config(font=("Consolas", 9, "bold"))
            lbl.grid(row=row_idx, column=col, sticky='nsew', padx=0, pady=0)
            self.fuel_case_labels[next_case_no][key] = lbl

        # Add to voyage_cases
        new_case = {
            "case_no": next_case_no,
            "wind_speed": 0.0,
            "wind_angle": 0.0,
            "ship_speed": 0.0,
            "percentage": 0.0,
        }
        if hasattr(self, "voyage_cases"):
            self.voyage_cases.append(new_case)
        else:
            self.voyage_cases = [new_case]

    def delete_last_voyage_row(self):
        """Remove the last row from the voyage cases table."""
        if not hasattr(self, "voyage_cases") or not self.voyage_cases:
            return

        last_case = self.voyage_cases[-1]
        case_no = last_case["case_no"]

        # Remove widgets from grid
        row_idx = case_no
        for widget in self.fuel_table_frame.grid_slaves(row=row_idx):
            widget.destroy()

        # Clean up tracking dicts
        self.fuel_case_editable_entries.pop(case_no, None)
        self.fuel_case_labels.pop(case_no, None)
        self.fuel_results.pop(case_no, None)
        self.voyage_cases.pop()

        self.update_fuel_summary()

    def sync_voyage_to_excel(self):
        """Write the current editable voyage table back to voyage details.xlsx."""
        try:
            cases = []
            for case_no in sorted(self.fuel_case_editable_entries.keys()):
                case = self._get_case_from_entries(case_no)
                if case:
                    cases.append(case)

            if not cases:
                messagebox.showwarning("Sync", "No voyage cases to sync.")
                return

            if not os.path.exists(self.voyage_file):
                messagebox.showerror("Sync Failed", f"Voyage file not found: {self.voyage_file}")
                return

            raw = pd.read_excel(self.voyage_file, header=None)

            # Find header row
            header_row = None
            for idx in range(len(raw)):
                row_values = [str(x).strip().lower() for x in raw.iloc[idx, :4].tolist()]
                if (
                    len(row_values) >= 4
                    and "true wind speed" in row_values[0]
                    and "true wind angle" in row_values[1]
                    and "ship speed" in row_values[2]
                    and "percentage" in row_values[3]
                ):
                    header_row = idx
                    break

            if header_row is None:
                messagebox.showerror("Sync Failed", "Could not find voyage case header row in Excel.")
                return

            # Find end of data rows
            data_start = header_row + 1
            data_end = data_start
            for row_idx in range(data_start, len(raw)):
                try:
                    row_text = " ".join(str(v).lower() for v in raw.iloc[row_idx, :12].tolist())
                    if "total voyage" in row_text:
                        break
                    vals = [self.safe_float(raw.iat[row_idx, c], None) for c in range(4)]
                    if any(v is None for v in vals):
                        break
                    data_end = row_idx + 1
                except Exception:
                    break

            # Build new data rows
            new_rows = []
            for case in cases:
                row = [None] * raw.shape[1]
                row[0] = case["wind_speed"]
                row[1] = case["wind_angle"]
                row[2] = case["ship_speed"]
                row[3] = case["percentage"]
                new_rows.append(row)

            # Replace data rows
            top_part = raw.iloc[:data_start]
            bottom_part = raw.iloc[data_end:]

            new_data_df = pd.DataFrame(new_rows, columns=raw.columns)
            result = pd.concat([top_part, new_data_df, bottom_part], ignore_index=True)
            result.to_excel(self.voyage_file, index=False, header=False)

            messagebox.showinfo("Sync Complete", f"Synced {len(cases)} voyage cases to {self.voyage_file}.")

        except Exception as e:
            messagebox.showerror("Sync Failed", f"Could not sync to Excel:\n{e}")

    def update_fuel_row(self, result):
        """Write one calculated case result into the voyage table."""
        labels = self.fuel_case_labels.get(result["case_no"], {})
        if not labels:
            return

        labels["propeller_rpm"].config(text=self.format_number(result.get('propeller_rpm'), 2))
        labels["fuel_required_with_sail_t"].config(text=self.format_number(result.get('fuel_required_with_sail_t'), 3, " t"))
        labels["fuel_required_without_sail_t"].config(text=self.format_number(result.get('fuel_required_without_sail_t'), 3, " t"))
        labels["fuel_saving_t"].config(text=self.format_number(result.get('fuel_saving_t'), 3, " t"))
        labels["percent_fuel_saving"].config(text=self.format_number(result.get('percent_fuel_saving'), 2, "%"))

    def update_fuel_summary(self):
        """Update final voyage totals from all calculated case results."""
        if not hasattr(self, "fuel_summary_labels"):
            return

        results = list(getattr(self, "fuel_results", {}).values())
        if not results:
            for label in self.fuel_summary_labels.values():
                label.config(text="--", fg=self.fg_white)
            return

        total_with = sum(r["fuel_required_with_sail_t"] for r in results)
        total_without = sum(r["fuel_required_without_sail_t"] for r in results)
        total_saving = sum(r["fuel_saving_t"] for r in results)
        overall_pct = (total_saving / total_without * 100.0) if total_without else 0.0

        self.fuel_summary_labels["Total fuel required with sail"].config(text=f"{total_with:.3f} t")
        self.fuel_summary_labels["Total fuel required without sail"].config(text=f"{total_without:.3f} t")
        self.fuel_summary_labels["Total fuel saving"].config(text=f"{total_saving:.3f} t", fg=self.fg_green)
        self.fuel_summary_labels["Overall percentage fuel saving"].config(text=f"{overall_pct:.2f}%", fg=self.fg_green)

    def run_fuel_case(self, case, show_popup=False, render_graph=True):
        """Run one case, update the table/summary, and optionally render the graph."""
        try:
            self.active_case_no = case["case_no"]
            self.highlight_active_case_row()
            if hasattr(self, "kinematics_active_case_label"):
                self.kinematics_active_case_label.config(text=f"Case: {self.active_case_no}")
                
            result = self.calculate_case_solution(case)
            self.fuel_results[result["case_no"]] = result
            self.update_fuel_row(result)
            self.update_fuel_summary()

            # Unlock frozen tabs
            self._unlock_frozen_tabs()

            if render_graph:
                self.show_optimization_tab()

            if show_popup:
                messagebox.showinfo(
                    "Fuel saving calculator",
                    f"Case {result['case_no']} results\n\n"
                    f"Propeller rpm: {result['propeller_rpm']:.3f}\n"
                    f"Fuel saving: {result['fuel_saving_t']:.3f} t\n"
                    f"Fuel required with sail: {result['fuel_required_with_sail_t']:.3f} t\n"
                    f"Fuel required without sail: {result['fuel_required_without_sail_t']:.3f} t\n"
                    f"Percentage fuel savings: {result['percent_fuel_saving']:.2f}%"
                )

            return result

        except Exception as e:
            messagebox.showerror("Fuel saving calculator", f"Case calculation failed:\n{e}")
            return None

    def calculate_all_fuel_cases(self):
        """Run every voyage case and show the graph for the last calculated case."""
        try:
            # Re-read from editable entries
            cases = []
            for case_no in sorted(self.fuel_case_editable_entries.keys()):
                case = self._get_case_from_entries(case_no)
                if case:
                    cases.append(case)

            if not cases:
                cases = getattr(self, "voyage_cases", None) or self.load_voyage_cases()

            self.fuel_results = {}
            last_result = None

            for case in cases:
                last_result = self.run_fuel_case(case, show_popup=False, render_graph=False)

            if last_result:
                self.set_case_inputs(last_result)
                self._unlock_frozen_tabs()

            self.update_fuel_summary()
            total_percentage = sum(case["percentage"] for case in cases)

            if abs(total_percentage - 100.0) > 1e-6:
                messagebox.showwarning(
                    "Fuel saving calculator",
                    f"All cases calculated, but the voyage percentage total is {total_percentage:.2f}% instead of 100%."
                )
            else:
                messagebox.showinfo("Fuel saving calculator", "All voyage cases calculated successfully.")

        except Exception as e:
            messagebox.showerror("Fuel saving calculator", f"Could not calculate all cases:\n{e}")

    # -------------------------------------------------------------------------
    # AOA optimisation solver and graph renderer
    # -------------------------------------------------------------------------
    def compute_3d_convergence(self, target_plot_box=None):
        plot_box = target_plot_box or self.opt_plot_box
        self.clear_children(plot_box)
        vectors = self.calculate_kinematics()
        if not vectors: return

        vx, vy, va, theta_deg, vadv, vs_ms, eta_h = vectors

        try:
            df_sail = self.parse_sail_table()
            aoa_raw = df_sail["AOA (deg)"].to_numpy(dtype=float)
            cl_raw = df_sail["Lift Coefficient (CL)"].to_numpy(dtype=float)
            cd_raw = df_sail["Drag Coefficient (CD)"].to_numpy(dtype=float)
            if np.max(aoa_raw) < 180.0:
                aoa_raw = np.append(aoa_raw, 180.0);
                cl_raw = np.append(cl_raw, 0.0);
                cd_raw = np.append(cd_raw, cd_raw[-1])

            df_prop = self.parse_propeller_table()
            j_raw = df_prop["J"].to_numpy(dtype=float)
            kt_raw = df_prop["KT"].to_numpy(dtype=float)
            eta_raw = df_prop["Eta_o"].to_numpy(dtype=float)
        except Exception as err:
            messagebox.showerror("Matrix Stream Exception", f"Error accessing workbook datasets:\n{err}")
            return

        # Build cubic splines for sail and propeller data.  The spline order is
        # lowered automatically if the user provides fewer than four rows.
        sail_spline_order = min(3, len(aoa_raw) - 1)
        prop_spline_order = min(3, len(j_raw) - 1)
        cl_spline = make_interp_spline(aoa_raw, cl_raw, k=sail_spline_order)
        cd_spline = make_interp_spline(aoa_raw, cd_raw, k=sail_spline_order)
        kt_spline = make_interp_spline(j_raw, kt_raw, k=prop_spline_order)
        eta_spline = make_interp_spline(j_raw, eta_raw, k=prop_spline_order)

        rho_a = float(self.entries["Density of Air, rho_a (kg/m3)"].get())
        rho_w = float(self.entries["Density of Water, rho_w (kg/m3)"].get())
        As = float(self.entries["Sail Area, As (m2)"].get())
        vs_knots = float(self.entries["Desired Ship Speed, vs (knots)"].get())
        RT = self.get_total_resistance(vs_knots)
        t = float(self.entries["Thrust Deduction Fraction, t"].get())
        D = float(self.entries["Propeller Diameter, D (m)"].get())
        eta_r = float(self.entries["Relative Rotative Efficiency, eta_R"].get())

        theta_rad = np.radians(theta_deg)

        # Grid parameters used to display the surface sheets gracefully
        aoa_dense = np.linspace(0, 180, 50)
        j_dense = np.linspace(0.01, np.max(j_raw), 50)
        J_mesh, AOA_mesh = np.meshgrid(j_dense, aoa_dense)

        KT_prop_mesh = kt_spline(J_mesh)
        KT_ship_mesh = np.zeros_like(J_mesh)

        intersect_J=[]
        intersect_AOA=[]
        intersect_KT=[]

        eta_open_curve=[]
        eta_propulsive_curve=[]
        power_delivered_curve=[]

        # Use a high-resolution J scanner so the intersection curve is smooth
        # and not visibly limited by the display grid resolution.
        j_high_res_scanner = np.linspace(0.01, np.max(j_raw), 1000)
        kt_prop_scanned = kt_spline(j_high_res_scanner)

        for idx, angle in enumerate(aoa_dense):
            lift = 0.5 * rho_a * As * (va ** 2) * cl_spline(angle)
            drag = 0.5 * rho_a * As * (va ** 2) * cd_spline(angle)
            t_sail = lift * np.sin(theta_rad) - drag * np.cos(theta_rad)
            t_prop = (RT - t_sail) / (1 - t)

            # Populate low-res mesh surface
            kt_j2_ship_surface = t_prop / (rho_w * (D ** 2) * (vadv ** 2))
            KT_ship_mesh[idx, :] = kt_j2_ship_surface * (j_dense ** 2)

            # Track high-res precise intersection coordinates
            kt_ship_scanned = kt_j2_ship_surface * (j_high_res_scanner ** 2)
            idx_match = np.argmin(np.abs(kt_prop_scanned - kt_ship_scanned))

            intersect_J.append(j_high_res_scanner[idx_match])
            intersect_AOA.append(angle)
            intersect_KT.append(kt_prop_scanned[idx_match])
            current_J = j_high_res_scanner[idx_match]

            eta_o = eta_spline(current_J)
            eta_d = eta_o * eta_h * eta_r
            PE=(RT-t_sail)*vs_ms

            if eta_d>0:
                PD=PE/eta_d
            else:
                PD=np.nan

            eta_open_curve.append(eta_o)
            eta_propulsive_curve.append(eta_d)
            power_delivered_curve.append(PD)

        eta_propulsive_curve=np.array(eta_propulsive_curve)

        best_opt_idx=np.nanargmax(eta_propulsive_curve)

        opt_aoa_val = intersect_AOA[best_opt_idx]
        opt_eta_val = eta_propulsive_curve[best_opt_idx]
        opt_j_val = intersect_J[best_opt_idx]
        opt_kt_val = intersect_KT[best_opt_idx]
        opt_PD_val = power_delivered_curve[best_opt_idx]

        case_prefix = f"Case {self.active_case_no} | " if hasattr(self, 'active_case_no') else ""
        self.hud_summary.config(
            text=
            f"{case_prefix}Optimal AOA {opt_aoa_val:.2f}°  ·  "
            f"J {opt_j_val:.3f}  ·  "
            f"RT {RT/1000:.3f} kN  ·  "
            f"ηD {opt_eta_val*100:.2f}%  ·  "
            f"PD {opt_PD_val/1e3:.2f} kW",
            fg=self.fg_green
        )

        fig=Figure(
            figsize=(14,6),
            facecolor=self.bg_card
        )

        # Left panel
        self.ax_3d=fig.add_subplot(
        121,
        projection='3d'
        )

        # Right panel
        ax_eff=fig.add_subplot(
        122
        )
        self.ax_3d.set_facecolor(self.bg_dark)

        surf1 = self.ax_3d.plot_surface(J_mesh, AOA_mesh, KT_prop_mesh, cmap='cool', alpha=0.35, edgecolor='none')
        surf2 = self.ax_3d.plot_surface(J_mesh, AOA_mesh, KT_ship_mesh, cmap='winter', alpha=0.35, edgecolor='none')

        # Smooth high-res interaction track rendering
        self.ax_3d.plot(intersect_J, intersect_AOA, intersect_KT, color='#FF007F', linewidth=3.5,
                        label='Operating Curve', zorder=10)
        self.ax_3d.scatter([opt_j_val], [opt_aoa_val], [opt_kt_val], color='#FFFF00', s=160, edgecolors='black',
                           label=f'Max efficient AOA ({opt_aoa_val:.1f}°)', zorder=15)

        self.ax_3d.set_xlabel('Advance Coefficient ($J$)', color=self.fg_blue, fontweight='bold', labelpad=12)
        self.ax_3d.set_ylabel(r'Angle of Attack ($AOA^\circ$)', color=self.fg_blue, fontweight='bold', labelpad=12)
        self.ax_3d.set_zlabel('Thrust Coefficient ($K_T$)', color=self.fg_blue, fontweight='bold', labelpad=12)

        self.ax_3d.tick_params(colors=self.fg_green, labelsize=9)
        self.ax_3d.xaxis.set_pane_color((0.0, 0.0, 0.0, 0.0))
        self.ax_3d.yaxis.set_pane_color((0.0, 0.0, 0.0, 0.0))
        self.ax_3d.zaxis.set_pane_color((0.0, 0.0, 0.0, 0.0))

        self.ax_3d.set_title("Propeller operating point by AOA", color=self.fg_green,
                             fontweight='bold', pad=15)
        self.ax_3d.legend(facecolor=self.bg_dark, edgecolor=self.fg_green, labelcolor=self.fg_white, loc='upper left')

        self.ax_3d.view_init(elev=22, azim=-45)

        # =====================================
        # STEP 7: Efficiency & Power curves
        # =====================================

        intersect_AOA=np.array(intersect_AOA)
        power_delivered_curve=np.array(power_delivered_curve)

        sort_idx=np.argsort(intersect_AOA)

        intersect_AOA=intersect_AOA[sort_idx]
        eta_propulsive_curve=eta_propulsive_curve[sort_idx]
        power_delivered_curve=power_delivered_curve[sort_idx]

        aoa_fine=np.linspace(
            min(intersect_AOA),
            max(intersect_AOA),
            300
        )

        eta_smooth=make_interp_spline(
            intersect_AOA,
            eta_propulsive_curve,
            k=3
        )(aoa_fine)

        power_smooth=make_interp_spline(
            intersect_AOA,
            power_delivered_curve/1e3,
            k=3
        )(aoa_fine)

        ax_eff.set_facecolor(self.bg_dark)
        fig.subplots_adjust(wspace=0.38)

        ax_eff.plot(
            aoa_fine,
            eta_smooth,
            linewidth=2,
            color=self.fg_green,
            label='ηD'
        )

        ax_eff.set_xlabel(
            "AOA (deg)",
            color=self.fg_white,
            fontweight='bold'
        )

        ax_eff.set_ylabel(
            "Propulsive Efficiency (ηD)",
            color=self.fg_green,
            fontweight='bold'
        )

        ax_eff.tick_params(
            axis='x',
            colors=self.fg_white
        )

        ax_eff.tick_params(
            axis='y',
            colors=self.fg_green
        )

        ax_eff.grid(
            True,
            color=self.grid_color,
            linestyle=':'
        )

        for spine in ax_eff.spines.values():
            spine.set_color(self.fg_white)

        ax_power=ax_eff.twinx()

        ax_power.plot(
            aoa_fine,
            power_smooth,
            linewidth=2,
            color=self.fg_blue,
            label='PD'
        )

        ax_power.set_facecolor(self.bg_dark)

        ax_power.set_ylabel(
            "Delivered Power (kW)",
            color=self.fg_blue,
            fontweight='bold'
        )

        ax_power.tick_params(
            axis='y',
            colors=self.fg_blue
        )

        for spine in ax_power.spines.values():
            spine.set_color(self.fg_white)

        ax_eff.scatter([opt_aoa_val], [opt_eta_val], color='#FFFF00', edgecolors='black', s=120, zorder=15, label='Max efficient AOA')

        ax_eff.set_title(
            "AOA vs propulsive efficiency and delivered power",
            color=self.fg_green,
            fontweight='bold',
            pad=15
        )

        lines1,labels1=ax_eff.get_legend_handles_labels()
        lines2,labels2=ax_power.get_legend_handles_labels()

        ax_eff.legend(
            lines1+lines2,
            labels1+labels2,
            facecolor=self.bg_card,
            edgecolor=self.fg_green,
            labelcolor=self.fg_white,
            loc='upper right'
        )

        fig.canvas.draw()

        self.orig_xlim=self.ax_3d.get_xlim3d()
        self.orig_ylim=self.ax_3d.get_ylim3d()
        self.orig_zlim=self.ax_3d.get_zlim3d()

        self.canvas_3d=FigureCanvasTkAgg(
            fig,
            master=plot_box
        )

        self.canvas_3d.draw()

        self.canvas_3d.get_tk_widget().pack(
            fill='both',
            expand=True
        )



if __name__ == "__main__":
    root = tk.Tk()
    app = CyberMarineOptimizer(root)
    root.mainloop()
