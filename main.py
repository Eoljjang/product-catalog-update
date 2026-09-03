import os
import sys
import traceback
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from tkinterdnd2 import DND_FILES, TkinterDnD
import xlwings as xlw


class ExcelProcessorApp(TkinterDnD.Tk):

    def __init__(self):
        super().__init__()

        self.title("Excel Batch Formula Processor (Debug Mode)")
        self.geometry("600x520")
        self.resizable(False, False)

        self.zone1_files = []
        self.zone2_file = None

        self._build_ui()

    def _build_ui(self):
        header = ttk.Label(
            self, text="Excel Batch Processor", font=("Segoe UI", 16, "bold")
        )
        header.pack(pady=(15, 10))

        # Zone 1
        z1_frame = ttk.LabelFrame(
            self, text=" Zone 1: Target Excel Files (Multiple) "
        )
        z1_frame.pack(fill="x", padx=20, pady=5)

        self.z1_box = tk.Listbox(
            z1_frame, height=5, selectmode=tk.MULTIPLE, relief="flat", bg="#f8f9fa"
        )
        self.z1_box.pack(fill="x", padx=10, pady=5)
        self.z1_box.drop_target_register(DND_FILES)
        self.z1_box.dnd_bind("<<Drop>>", self._on_drop_zone1)

        z1_hint = ttk.Label(
            z1_frame,
            text="Drag and drop Excel files here (double-click to clear)",
            font=("Segoe UI", 8, "italic"),
            foreground="gray",
        )
        z1_hint.pack(pady=(0, 5))
        self.z1_box.bind("<Double-Button-1>", lambda e: self._clear_zone1())

        # Zone 2
        z2_frame = ttk.LabelFrame(self, text=" Zone 2: Product List File (Single) ")
        z2_frame.pack(fill="x", padx=20, pady=10)

        self.z2_label = tk.Label(
            z2_frame,
            text="Drag & Drop Product List Excel File Here",
            bg="#e9ecef",
            height=3,
            relief="groove",
            font=("Segoe UI", 9, "italic"),
        )
        self.z2_label.pack(fill="x", padx=10, pady=5)
        self.z2_label.drop_target_register(DND_FILES)
        self.z2_label.dnd_bind("<<Drop>>", self._on_drop_zone2)

        # Process Button
        self.btn_run = ttk.Button(
            self, text="Start Processing", command=self.process_files
        )
        self.btn_run.pack(pady=15, ipadx=10, ipady=5)

        # Status Label
        self.status_var = tk.StringVar(value="Ready")
        self.lbl_status = ttk.Label(
            self, textvariable=self.status_var, font=("Segoe UI", 9)
        )
        self.lbl_status.pack(pady=5)

    def _show_error_dialog(self, title, error_msg, tb_text):
        """Prints error to terminal AND pops up a detailed scrollable dialog."""
        print(f"\n{'='*20} ERROR DETECTED {'='*20}\n", file=sys.stderr)
        print(tb_text, file=sys.stderr)
        print(f"{'='*56}\n", file=sys.stderr)

        err_win = tk.Toplevel(self)
        err_win.title(title)
        err_win.geometry("600x400")

        lbl = ttk.Label(
            err_win, text=error_msg, font=("Segoe UI", 10, "bold"), foreground="red"
        )
        lbl.pack(anchor="w", padx=10, pady=(10, 5))

        txt = ScrolledText(err_win, wrap="none")
        txt.pack(fill="both", expand=True, padx=10, pady=5)
        txt.insert("1.0", tb_text)
        txt.config(state="disabled")

        btn = ttk.Button(err_win, text="Close", command=err_win.destroy)
        btn.pack(pady=5)

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
                bg="#d1e7dd",
                fg="#0f5132",
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

                # Prepare structured formulas
                vlookup_formula = f'=IFERROR(VLOOKUP([@CODE], {ext_ref}, 1, FALSE), "")'
                oo1_formula = '=IF([@OH1]="", [@COST], [@COST]*1.025)'
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
                    
                    # Native COM copy and value-paste keeps vertical orientation
                    oh2_range.api.Copy()
                    cost_range.api.PasteSpecial(Paste=-4163)  # -4163 = xlPasteValues
                    app.api.CutCopyMode = False

                    # Clear data contents for OH1, OO1, and OH2 (leaving headers)
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

                        # Clear data rows 7 down to last_row for OH1, OO1, and OH2
                        sheet.range((7, oh1_col), (last_row, oh1_col)).clear_contents()
                        sheet.range((7, oo1_col), (last_row, oo1_col)).clear_contents()
                        sheet.range((7, oh2_col), (last_row, oh2_col)).clear_contents()

                # Highlight cells A3, B3 yellow
                sheet.range("A3:B3").color = (255, 255, 0)

                # Save copy into an "updated" folder in the same directory
                dir_name, full_filename = os.path.split(abs_target)
                filename, ext = os.path.splitext(full_filename)

                output_folder = os.path.join(dir_name, "updated")
                os.makedirs(output_folder, exist_ok=True)

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
    app = ExcelProcessorApp()
    app.mainloop()