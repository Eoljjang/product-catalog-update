import os
import sys
import traceback
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from tkinter.scrolledtext import ScrolledText
from tkinterdnd2 import DND_FILES, TkinterDnD
import xlwings as xlw


class CatalogUpdateApp(TkinterDnD.Tk):

    def __init__(self):
        super().__init__()

        self.title("Catalog Update Automation (Debug Mode)")
        self.geometry("620x600")
        self.resizable(False, False)

        # Apply dark theme palette
        self.configure(bg="#1e1e1e")
        self._apply_dark_theme()

        self.zone1_files = []
        self.zone2_file = None

        self._build_ui()

    def _apply_dark_theme(self):
        """Configures ttk styles for a sleek dark mode look."""
        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        # Color definitions
        BG_DARK = "#1e1e1e"
        FRAME_BG = "#2d2d2d"
        TEXT_LIGHT = "#e0e0e0"
        ACCENT_BLUE = "#007acc"
        ACCENT_HOVER = "#005999"

        # General TTK Styles
        self.style.configure(".", background=BG_DARK, foreground=TEXT_LIGHT)

        # LabelFrame
        self.style.configure(
            "TLabelframe",
            background=BG_DARK,
            foreground=TEXT_LIGHT,
            borderwidth=1,
            relief="solid",
        )
        self.style.configure(
            "TLabelframe.Label",
            background=BG_DARK,
            foreground="#a0a0a0",
            font=("Segoe UI", 12, "bold"),
        )

        # Buttons
        self.style.configure(
            "Accent.TButton",
            font=("Segoe UI", 14, "bold"),
            background=ACCENT_BLUE,
            foreground="#ffffff",
            borderwidth=0,
            focusthickness=0,
            padding=8,
        )
        self.style.map(
            "Accent.TButton",
            background=[("active", ACCENT_HOVER), ("pressed", "#004080")],
        )

        self.style.configure(
            "TButton",
            font=("Segoe UI", 12),
            background="#3c3c3c",
            foreground=TEXT_LIGHT,
            borderwidth=0,
            padding=5,
        )
        self.style.map(
            "TButton",
            background=[("active", "#4a4a4a")],
        )

        # Custom Entry style to ensure text is black
        self.style.configure(
            "BlackText.TEntry",
            fieldbackground="#ffffff",
            foreground="#000000",
        )

    def _build_ui(self):
        header = ttk.Label(
            self,
            text="Catalog Update Automation",
            font=("Segoe UI", 16, "bold"),
            background="#1e1e1e",
            foreground="#ffffff",
        )
        header.pack(pady=(20, 10))

        # Zone 1
        z1_frame = ttk.LabelFrame(
            self, text="Product Catalog Files (Multiple) - Drag & Drop Here"
        )
        z1_frame.pack(fill="x", padx=25, pady=8)

        self.z1_box = tk.Listbox(
            z1_frame,
            height=5,
            selectmode=tk.MULTIPLE,
            relief="flat",
            bg="#252526",
            fg="#d4d4d4",
            selectbackground="#04395e",
            selectforeground="#ffffff",
            highlightthickness=1,
            highlightbackground="#3c3c3c",
            font=("Segoe UI", 9),
        )
        self.z1_box.pack(fill="x", padx=12, pady=8)
        self.z1_box.drop_target_register(DND_FILES)
        self.z1_box.dnd_bind("<<Drop>>", self._on_drop_zone1)

        z1_hint = ttk.Label(
            z1_frame,
            text="Drag and drop Excel files here (double-click to clear)",
            font=("Segoe UI", 10, "italic"),
            background="#1e1e1e",
            foreground="#858585",
        )
        z1_hint.pack(pady=(0, 8))
        self.z1_box.bind("<Double-Button-1>", lambda e: self._clear_zone1())

        # Zone 2
        z2_frame = ttk.LabelFrame(self, text=" MMS Product List File - Drag & Drop. Only column A should contain data \nand it must be numeric!")
        z2_frame.pack(fill="x", padx=25, pady=8)

        self.z2_label = tk.Label(
            z2_frame,
            text="Drag & Drop Product List Excel File Here",
            bg="#252526",
            fg="#858585",
            height=3,
            relief="flat",
            highlightthickness=1,
            highlightbackground="#3c3c3c",
            font=("Segoe UI", 10, "italic"),
        )
        self.z2_label.pack(fill="x", padx=12, pady=10)
        self.z2_label.drop_target_register(DND_FILES)
        self.z2_label.dnd_bind("<<Drop>>", self._on_drop_zone2)

        # Markup Percentage Input Frame
        markup_frame = ttk.Frame(self)
        markup_frame.pack(fill="x", padx=25, pady=8)

        lbl_markup = ttk.Label(
            markup_frame,
            text="Markup %",
            font=("Segoe UI", 14, "bold"),
            background="#1e1e1e",
            foreground="#e0e0e0",
        )
        lbl_markup.pack(side="left", padx=(5, 10))

        # Using standard tk.Entry instead of ttk.Entry for direct control over text and cursor colors
        self.markup_entry = tk.Entry(
            markup_frame,
            font=("Segoe UI", 14),
            width=10,
            bg="#ffffff",
            fg="#000000",
            insertbackground="#000000",
            relief="flat",
            highlightthickness=1,
            highlightbackground="#3c3c3c",
        )
        self.markup_entry.pack(side="left", ipady=3)
        self.markup_entry.insert(0, "2.5")  # Default value
        
        # Process Button
        self.btn_run = ttk.Button(
            self,
            text="Start Processing",
            style="Accent.TButton",
            command=self.process_files,
        )
        self.btn_run.pack(pady=(10, 5), ipadx=15)

        # Status Label
        self.status_var = tk.StringVar(value="Ready")
        self.lbl_status = ttk.Label(
            self,
            textvariable=self.status_var,
            font=("Segoe UI", 9),
            background="#1e1e1e",
            foreground="#a0a0a0",
        )
        self.lbl_status.pack(pady=5)

    def _show_error_dialog(self, title, error_msg, tb_text):
        """Prints error to terminal AND pops up a detailed scrollable dark dialog."""
        print(f"\n{'='*20} ERROR DETECTED {'='*20}\n", file=sys.stderr)
        print(tb_text, file=sys.stderr)
        print(f"{'='*56}\n", file=sys.stderr)

        err_win = tk.Toplevel(self)
        err_win.title(title)
        err_win.geometry("600x400")
        err_win.configure(bg="#1e1e1e")

        lbl = ttk.Label(
            err_win,
            text=error_msg,
            font=("Segoe UI", 10, "bold"),
            background="#1e1e1e",
            foreground="#f44747",
        )
        lbl.pack(anchor="w", padx=12, pady=(12, 6))

        txt = ScrolledText(
            err_win,
            wrap="none",
            bg="#252526",
            fg="#f44747",
            insertbackground="#ffffff",
            relief="flat",
            font=("Consolas", 9),
        )
        txt.pack(fill="both", expand=True, padx=12, pady=6)
        txt.insert("1.0", tb_text)
        txt.config(state="disabled")

        btn = ttk.Button(err_win, text="Close", command=err_win.destroy)
        btn.pack(pady=8)

    def _parse_drop_files(self, event_data):
        files = []
        raw_files = self.tk.splitlist(event_data)
        for f in raw_files:
            f_clean = f.strip("{}")
            if f_clean.endswith((".xlsx", ".xls", ".xlsm")):
                files.append(f_clean)
        return files

    def _on_drop_zone1(self, event):
        files = self._parse_drop_files(event.data)
        for f in files:
            if f not in self.zone1_files:
                self.zone1_files.append(f)
                self.z1_box.insert(tk.END, os.path.basename(f))

    def _clear_zone1(self):
        self.zone1_files.clear()
        self.z1_box.delete(0, tk.END)

    def _on_drop_zone2(self, event):
        files = self._parse_drop_files(event.data)
        if not files:
            messagebox.showwarning(
                "Invalid File", "Please drop a valid Excel file."
            )
            return

        file_path = files[0]
        if self._validate_product_list(file_path):
            self.zone2_file = file_path
            self.z2_label.config(
                text=f"Loaded: {os.path.basename(file_path)}",
                bg="#1e3a1e",
                fg="#4ec9b0",
                font=("Segoe UI", 9, "bold"),
            )

    def _validate_product_list(self, file_path):
        """Validates that only Column A contains data."""
        abs_path = os.path.abspath(file_path)
        app = xlw.App(visible=False)
        app.display_alerts = False

        try:
            wb = app.books.open(abs_path, read_only=True)
            sheet = wb.sheets[0]

            used_range = sheet.used_range
            if used_range is None or used_range.value is None:
                messagebox.showerror("Validation Error", "File appears to be empty.")
                wb.close()
                return False

            if used_range.last_cell.column > 1:
                messagebox.showerror(
                    "Validation Error",
                    "Validation failed!\nColumns beyond Column A contain data.",
                )
                wb.close()
                return False

            wb.close()
            return True

        except Exception as e:
            tb_str = traceback.format_exc()
            self._show_error_dialog(
                "Validation Error", f"Failed to validate file: {e}", tb_str
            )
            return False
        finally:
            app.quit()

    def process_files(self):
        if not self.zone1_files:
            messagebox.showwarning(
                "Missing Input", "Please add at least one Excel file in Zone 1."
            )
            return
        if not self.zone2_file:
            messagebox.showwarning(
                "Missing Input", "Please select a valid Product List in Zone 2."
            )
            return

        # Validate and retrieve Markup Percentage
        markup_text = self.markup_entry.get().strip()
        if not markup_text:
            messagebox.showwarning(
                "Missing Input", "Markup % field is required."
            )
            return
        try:
            markup_val = float(markup_text)
        except ValueError:
            messagebox.showerror(
                "Invalid Input", "Markup % must be a valid numeric number."
            )
            return

        # Prompt user to choose destination directory
        selected_dir = filedialog.askdirectory(
            title="Select Output Directory for Updated Catalog Folder"
        )
        if not selected_dir:
            self.status_var.set("Processing cancelled.")
            return

        self.status_var.set("Processing files...")
        self.update()

        app = xlw.App(visible=False)
        app.display_alerts = False

        try:
            abs_p_list = os.path.abspath(self.zone2_file)
            p_list_wb = app.books.open(abs_p_list, read_only=True)
            p_list_name = p_list_wb.name
            p_sheet_name = p_list_wb.sheets[0].name

            ext_ref = f"'[{p_list_name}]{p_sheet_name}'!$A:$A"

            # Format current date as MM-DD-YYYY for clean directory naming
            date_str = datetime.now().strftime("%m-%d-%Y")
            folder_name = f"catalog update - {date_str}"
            output_folder = os.path.join(selected_dir, folder_name)
            os.makedirs(output_folder, exist_ok=True)

            for target_path in self.zone1_files:
                abs_target = os.path.abspath(target_path)
                self.status_var.set(
                    f"Processing: {os.path.basename(abs_target)}"
                )
                self.update()

                wb = app.books.open(abs_target)
                sheet = wb.sheets[0]

                # Find headers on row 6
                row_6_vals = sheet.range("6:6").value
                oh1_col = None
                oo1_col = None
                oh2_col = None
                cost_col = None

                if row_6_vals:
                    for idx, val in enumerate(row_6_vals, start=1):
                        if val == "OH1":
                            oh1_col = idx
                        elif val == "OO1":
                            oo1_col = idx
                        elif val == "OH2":
                            oh2_col = idx
                        elif val == "COST":
                            cost_col = idx

                if not all([oh1_col, oo1_col, oh2_col, cost_col]):
                    missing = []
                    if not oh1_col: missing.append("OH1")
                    if not oo1_col: missing.append("OO1")
                    if not oh2_col: missing.append("OH2")
                    if not cost_col: missing.append("COST")
                    
                    messagebox.showerror(
                        "Header Error",
                        f"Missing headers {missing} in row 6 of:\n{target_path}",
                    )
                    wb.close()
                    continue

                # Calculate multiplier from user input (e.g. 2.5% -> 1 + 2.5/100 = 1.025)
                multiplier = 1 + (markup_val / 100)

                # Prepare structured formulas
                vlookup_formula = f'=IFERROR(VLOOKUP([@CODE], {ext_ref}, 1, FALSE), "")'
                oo1_formula = f'=IF([@OH1]="", [@COST], [@COST]*{multiplier})'
                oh2_formula = '=IF([@OO1]=0, "", [@OO1])'

                # Target cells on Row 7
                oh1_target = sheet.cells(7, oh1_col)
                oo1_target = sheet.cells(7, oo1_col)
                oh2_target = sheet.cells(7, oh2_col)

                try:
                    oh1_target.formula2 = vlookup_formula
                except AttributeError:
                    oh1_target.formula = vlookup_formula

                try:
                    oo1_target.formula2 = oo1_formula
                except AttributeError:
                    oo1_target.formula = oo1_formula

                try:
                    oh2_target.formula2 = oh2_formula
                except AttributeError:
                    oh2_target.formula = oh2_formula

                # Paste values vertically from OH2 column into COST column
                table = sheet.tables[0] if sheet.tables else None
                if table:
                    oh2_range = sheet.range(table.api.ListColumns("OH2").DataBodyRange.Address)
                    cost_range = sheet.range(table.api.ListColumns("COST").DataBodyRange.Address)
                    
                    oh2_range.api.Copy()
                    cost_range.api.PasteSpecial(Paste=-4163)  # -4163 = xlPasteValues
                    app.api.CutCopyMode = False

                    sheet.range(table.api.ListColumns("OH1").DataBodyRange.Address).clear_contents()
                    sheet.range(table.api.ListColumns("OO1").DataBodyRange.Address).clear_contents()
                    sheet.range(table.api.ListColumns("OH2").DataBodyRange.Address).clear_contents()
                else:
                    last_row = sheet.used_range.last_cell.row
                    if last_row >= 7:
                        oh2_range = sheet.range((7, oh2_col), (last_row, oh2_col))
                        cost_range = sheet.range((7, cost_col), (last_row, cost_col))
                        
                        oh2_range.api.Copy()
                        cost_range.api.PasteSpecial(Paste=-4163)
                        app.api.CutCopyMode = False

                        sheet.range((7, oh1_col), (last_row, oh1_col)).clear_contents()
                        sheet.range((7, oo1_col), (last_row, oo1_col)).clear_contents()
                        sheet.range((7, oh2_col), (last_row, oh2_col)).clear_contents()

                # Highlight cells A3, B3 yellow
                sheet.range("A3:B3").color = (255, 255, 0)

                # Save copy inside the folder
                _, full_filename = os.path.split(abs_target)
                filename, ext = os.path.splitext(full_filename)

                new_filepath = os.path.join(
                    output_folder, f"{filename} - updated{ext}"
                )

                wb.save(new_filepath)
                wb.close()

            p_list_wb.close()
            messagebox.showinfo(
                "Success", "All target files processed and saved successfully!"
            )
            self.status_var.set("Processing completed successfully.")

        except Exception as e:
            tb_str = traceback.format_exc()
            self._show_error_dialog(
                "Execution Error", f"An exception occurred: {e}", tb_str
            )
            self.status_var.set("Error encountered.")

        finally:
            app.quit()


if __name__ == "__main__":
    app = CatalogUpdateApp()
    app.mainloop()