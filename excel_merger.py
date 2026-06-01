"""
Excel Merger — modern GUI
Requires: customtkinter, pandas, openpyxl
"""

import threading
from pathlib import Path

import customtkinter as ctk
import pandas as pd
from tkinter import filedialog, messagebox

# ── Theme ──────────────────────────────────────────────────────────────── #
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".xlsb"}

ACCENT   = "#3B82F6"   # blue-500
SUCCESS  = "#22C55E"   # green-500
WARNING  = "#F59E0B"   # amber-500
ERROR    = "#EF4444"   # red-500
MUTED    = "#6B7280"   # gray-500
BG_CARD  = "#1E293B"   # slate-800
BG_INPUT = "#0F172A"   # slate-900


# ════════════════════════════════════════════════════════════════════════ #
#  Helper widgets                                                          #
# ════════════════════════════════════════════════════════════════════════ #

class Card(ctk.CTkFrame):
    """A slightly raised card panel."""
    def __init__(self, master, **kw):
        super().__init__(master, fg_color=BG_CARD, corner_radius=12, **kw)


class SectionLabel(ctk.CTkLabel):
    def __init__(self, master, text, **kw):
        super().__init__(
            master, text=text,
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=MUTED,
            **kw,
        )


class PathRow(ctk.CTkFrame):
    """Label + entry + browse button in one row."""

    def __init__(self, master, label: str, placeholder: str,
                 browse_cmd, **kw):
        super().__init__(master, fg_color="transparent", **kw)
        self.columnconfigure(1, weight=1)

        ctk.CTkLabel(self, text=label,
                     font=ctk.CTkFont(size=13),
                     width=110, anchor="w").grid(row=0, column=0, sticky="w")

        self.var = ctk.StringVar()
        self._entry = ctk.CTkEntry(
            self, textvariable=self.var,
            placeholder_text=placeholder,
            height=36, corner_radius=8,
            fg_color=BG_INPUT, border_color="#334155",
        )
        self._entry.grid(row=0, column=1, sticky="ew", padx=(8, 8))

        ctk.CTkButton(
            self, text="Browse", width=80, height=36,
            corner_radius=8, fg_color="#334155",
            hover_color="#475569", command=browse_cmd,
        ).grid(row=0, column=2)

    def get(self) -> str:
        return self.var.get().strip()


class OptionRow(ctk.CTkFrame):
    """Label + entry/spinbox + hint in one row."""

    def __init__(self, master, label: str, hint: str = "",
                 spinbox: bool = False, **kw):
        super().__init__(master, fg_color="transparent", **kw)
        self.columnconfigure(2, weight=1)

        ctk.CTkLabel(self, text=label,
                     font=ctk.CTkFont(size=13),
                     width=130, anchor="w").grid(row=0, column=0, sticky="w")

        self.var = ctk.StringVar()
        if spinbox:
            self.var.set("0")
            self._widget = ctk.CTkEntry(
                self, textvariable=self.var,
                width=70, height=36, corner_radius=8,
                fg_color=BG_INPUT, border_color="#334155",
                justify="center",
            )
        else:
            self._widget = ctk.CTkEntry(
                self, textvariable=self.var,
                width=220, height=36, corner_radius=8,
                fg_color=BG_INPUT, border_color="#334155",
            )
        self._widget.grid(row=0, column=1, padx=(8, 8))

        if hint:
            ctk.CTkLabel(self, text=hint,
                         font=ctk.CTkFont(size=11),
                         text_color=MUTED, anchor="w").grid(
                row=0, column=2, sticky="w")

    def get(self) -> str:
        return self.var.get().strip()


# ════════════════════════════════════════════════════════════════════════ #
#  Main application                                                        #
# ════════════════════════════════════════════════════════════════════════ #

class ExcelMergerApp(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Excel Merger")
        self.geometry("780x760")
        self.minsize(700, 680)
        self._build_ui()

    # ------------------------------------------------------------------ #
    #  UI                                                                  #
    # ------------------------------------------------------------------ #

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # ── Header ─────────────────────────────────────────────────────── #
        header = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=0, height=64)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.columnconfigure(1, weight=1)

        ctk.CTkLabel(
            header,
            text="⊞  Excel Merger",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white",
        ).grid(row=0, column=0, padx=20, pady=16, sticky="w")

        # Theme toggle
        self._theme_mode = ctk.StringVar(value="dark")
        ctk.CTkSegmentedButton(
            header,
            values=["dark", "light"],
            variable=self._theme_mode,
            command=self._toggle_theme,
            width=120, height=30,
            font=ctk.CTkFont(size=12),
        ).grid(row=0, column=1, padx=20, sticky="e")

        # ── Scrollable body ────────────────────────────────────────────── #
        body = ctk.CTkScrollableFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        body.columnconfigure(0, weight=1)

        pad = {"padx": 16, "pady": (0, 12)}

        # ┌─ Paths card ──────────────────────────────────────────────────┐
        paths_card = Card(body)
        paths_card.grid(row=0, column=0, sticky="ew", **pad)
        paths_card.columnconfigure(0, weight=1)

        SectionLabel(paths_card, "📁  FILES").grid(
            row=0, column=0, sticky="w", padx=16, pady=(12, 6))

        self._folder_row = PathRow(
            paths_card, "Input folder",
            "Select folder containing Excel files",
            self._browse_folder,
        )
        self._folder_row.grid(row=1, column=0, sticky="ew", padx=16, pady=4)

        self._output_row = PathRow(
            paths_card, "Output file",
            "Where to save the merged file",
            self._browse_output,
        )
        self._output_row.grid(row=2, column=0, sticky="ew",
                              padx=16, pady=(4, 14))

        # ┌─ Options card ────────────────────────────────────────────────┐
        opt_card = Card(body)
        opt_card.grid(row=1, column=0, sticky="ew", **pad)
        opt_card.columnconfigure(0, weight=1)

        SectionLabel(opt_card, "⚙  OPTIONS").grid(
            row=0, column=0, sticky="w", padx=16, pady=(12, 6))

        self._sheet_row = OptionRow(
            opt_card, "Sheet name",
            hint="leave empty → first sheet",
        )
        self._sheet_row.grid(row=1, column=0, sticky="ew", padx=16, pady=4)

        self._skip_rows_row = OptionRow(
            opt_card, "Skip rows",
            hint="rows to skip before the header",
            spinbox=True,
        )
        self._skip_rows_row.grid(row=2, column=0, sticky="ew", padx=16, pady=4)

        self._skip_cols_row = OptionRow(
            opt_card, "Skip columns",
            hint="e.g.  A,C  or  0,2  or  Name,Date",
        )
        self._skip_cols_row.grid(row=3, column=0, sticky="ew",
                                 padx=16, pady=(4, 14))

        # ┌─ Cleaning card ───────────────────────────────────────────────┐
        clean_card = Card(body)
        clean_card.grid(row=2, column=0, sticky="ew", **pad)
        clean_card.columnconfigure(0, weight=1)

        SectionLabel(clean_card, "🧹  DATA CLEANING").grid(
            row=0, column=0, sticky="w", padx=16, pady=(12, 6))

        checks = ctk.CTkFrame(clean_card, fg_color="transparent")
        checks.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 14))

        self._dedup_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            checks, text="Remove duplicate rows",
            variable=self._dedup_var,
            font=ctk.CTkFont(size=13),
            checkbox_width=20, checkbox_height=20,
            corner_radius=4, border_width=2,
        ).grid(row=0, column=0, padx=(0, 24), pady=4, sticky="w")

        self._drop_empty_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            checks, text="Drop fully-empty rows",
            variable=self._drop_empty_var,
            font=ctk.CTkFont(size=13),
            checkbox_width=20, checkbox_height=20,
            corner_radius=4, border_width=2,
        ).grid(row=0, column=1, pady=4, sticky="w")

        # ┌─ Action row ──────────────────────────────────────────────────┐
        action_frame = ctk.CTkFrame(body, fg_color="transparent")
        action_frame.grid(row=3, column=0, sticky="ew",
                          padx=16, pady=(0, 12))
        action_frame.columnconfigure(0, weight=1)

        self._merge_btn = ctk.CTkButton(
            action_frame,
            text="  Merge Excel Files",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=46, corner_radius=10,
            fg_color=ACCENT, hover_color="#2563EB",
            command=self._start_merge,
        )
        self._merge_btn.grid(row=0, column=0, sticky="ew")

        # ┌─ Progress ────────────────────────────────────────────────────┐
        self._progress = ctk.CTkProgressBar(
            body, mode="indeterminate",
            height=6, corner_radius=3,
            progress_color=ACCENT, fg_color="#1E293B",
        )
        self._progress.grid(row=4, column=0, sticky="ew",
                            padx=16, pady=(0, 12))
        self._progress.set(0)

        # ┌─ Log card ────────────────────────────────────────────────────┐
        log_card = Card(body)
        log_card.grid(row=5, column=0, sticky="nsew", **pad)
        log_card.columnconfigure(0, weight=1)
        log_card.rowconfigure(1, weight=1)

        log_header = ctk.CTkFrame(log_card, fg_color="transparent")
        log_header.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 6))
        log_header.columnconfigure(0, weight=1)

        SectionLabel(log_header, "📋  LOG").grid(row=0, column=0, sticky="w")

        ctk.CTkButton(
            log_header, text="Clear", width=60, height=26,
            corner_radius=6, fg_color="#334155",
            hover_color="#475569",
            font=ctk.CTkFont(size=11),
            command=self._clear_log,
        ).grid(row=0, column=1, sticky="e")

        self._log_box = ctk.CTkTextbox(
            log_card, height=200,
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=BG_INPUT, corner_radius=8,
            border_width=1, border_color="#334155",
            state="disabled", wrap="word",
        )
        self._log_box.grid(row=1, column=0, sticky="nsew",
                           padx=16, pady=(0, 14))

    # ------------------------------------------------------------------ #
    #  Theme toggle                                                        #
    # ------------------------------------------------------------------ #

    def _toggle_theme(self, value: str):
        ctk.set_appearance_mode(value)

    # ------------------------------------------------------------------ #
    #  Browse dialogs                                                      #
    # ------------------------------------------------------------------ #

    def _browse_folder(self):
        path = filedialog.askdirectory(title="Select folder with Excel files")
        if path:
            self._folder_row.var.set(path)

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save merged file as",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self._output_row.var.set(path)

    # ------------------------------------------------------------------ #
    #  Log helpers                                                         #
    # ------------------------------------------------------------------ #

    def _log(self, msg: str, color: str | None = None):
        self._log_box.configure(state="normal")
        self._log_box.insert("end", msg + "\n")
        self._log_box.see("end")
        self._log_box.configure(state="disabled")
        self.update_idletasks()

    def _clear_log(self):
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

    # ------------------------------------------------------------------ #
    #  Column parsing                                                      #
    # ------------------------------------------------------------------ #

    @staticmethod
    def _col_letter_to_index(letter: str) -> int:
        idx = 0
        for ch in letter.upper():
            idx = idx * 26 + (ord(ch) - ord("A") + 1)
        return idx - 1

    def _resolve_skip_columns(self, raw: str, columns) -> list:
        if not raw.strip():
            return []
        cols = list(columns)
        to_drop = []
        for token in raw.split(","):
            token = token.strip()
            if not token:
                continue
            if token.replace(" ", "").isalpha() and len(token) <= 3:
                idx = self._col_letter_to_index(token)
                if 0 <= idx < len(cols):
                    to_drop.append(cols[idx])
                else:
                    self._log(f"  ⚠  Column letter '{token}' out of range")
            elif token.lstrip("-").isdigit():
                idx = int(token)
                if 0 <= idx < len(cols):
                    to_drop.append(cols[idx])
                else:
                    self._log(f"  ⚠  Column index {idx} out of range")
            elif token in cols:
                to_drop.append(token)
            else:
                self._log(f"  ⚠  Column '{token}' not found")
        return to_drop

    # ------------------------------------------------------------------ #
    #  Merge (background thread)                                           #
    # ------------------------------------------------------------------ #

    def _start_merge(self):
        self._merge_btn.configure(state="disabled", text="  Merging…")
        self._progress.configure(mode="indeterminate")
        self._progress.start()
        threading.Thread(target=self._merge_thread, daemon=True).start()

    def _merge_thread(self):
        try:
            self._do_merge()
        except Exception as exc:
            self._log(f"\n✗  Unexpected error: {exc}")
            messagebox.showerror("Error", str(exc))
        finally:
            self._progress.stop()
            self._progress.set(0)
            self._merge_btn.configure(state="normal",
                                      text="  Merge Excel Files")

    def _do_merge(self):
        folder      = self._folder_row.get()
        output      = self._output_row.get()
        sheet_name  = self._sheet_row.get() or 0
        skip_cols_r = self._skip_cols_row.get()

        try:
            skip_rows = int(self._skip_rows_row.get() or 0)
        except ValueError:
            messagebox.showerror("Validation", "Skip Rows must be a whole number.")
            return

        if not folder:
            messagebox.showerror("Validation", "Please select an input folder.")
            return
        if not output:
            messagebox.showerror("Validation", "Please specify an output file.")
            return

        self._log("─" * 52)
        self._log(f"  Folder  : {folder}")
        self._log(f"  Output  : {output}")
        self._log(f"  Sheet   : {'first sheet' if sheet_name == 0 else sheet_name}")
        self._log(f"  Skip rows: {skip_rows}   Skip cols: {skip_cols_r or '—'}")
        self._log(f"  Dedup   : {'yes' if self._dedup_var.get() else 'no'}")
        self._log("─" * 52 + "\n")

        excel_files = sorted(
            f for f in Path(folder).iterdir()
            if f.suffix.lower() in EXCEL_EXTENSIONS
        )

        if not excel_files:
            messagebox.showwarning(
                "No files found",
                "No Excel files (.xlsx / .xls / .xlsm / .xlsb) found in that folder.",
            )
            return

        self._log(f"  Found {len(excel_files)} file(s)\n")

        frames: list[pd.DataFrame] = []
        errors: list[str] = []

        for i, fp in enumerate(excel_files, 1):
            self._log(f"  [{i}/{len(excel_files)}]  {fp.name}")
            try:
                df = pd.read_excel(
                    fp,
                    sheet_name=sheet_name,
                    skiprows=skip_rows,
                    dtype=str,
                )

                if self._drop_empty_var.get():
                    df.dropna(how="all", inplace=True)

                to_drop = self._resolve_skip_columns(skip_cols_r, df.columns)
                if to_drop:
                    df.drop(columns=to_drop, inplace=True, errors="ignore")
                    self._log(f"         dropped cols: {to_drop}")

                df.insert(0, "Source File", fp.name)
                frames.append(df)
                self._log(
                    f"         ✓  {len(df):,} rows × "
                    f"{len(df.columns) - 1:,} data cols"
                )

            except Exception as exc:
                msg = f"         ✗  {exc}"
                self._log(msg)
                errors.append(f"{fp.name}: {exc}")

        if not frames:
            messagebox.showerror("Error", "Could not read any data from the files.")
            return

        self._log(f"\n  Merging {len(frames)} frame(s)…")
        merged = pd.concat(frames, ignore_index=True, sort=False)
        self._log(f"  Combined : {len(merged):,} rows")

        if self._dedup_var.get():
            before = len(merged)
            data_cols = [c for c in merged.columns if c != "Source File"]
            merged.drop_duplicates(subset=data_cols, inplace=True)
            self._log(
                f"  Deduped  : removed {before - len(merged):,}, "
                f"kept {len(merged):,}"
            )

        self._log(f"\n  Saving → {output}")
        merged.to_excel(output, index=False)
        self._log("  Done ✓\n")

        summary = (
            f"Merged {len(frames)} file(s)\n"
            f"{len(merged):,} rows saved to:\n{output}"
        )
        if errors:
            summary += f"\n\n⚠  {len(errors)} file(s) had errors (see log)."
        messagebox.showinfo("Merge complete", summary)


# ════════════════════════════════════════════════════════════════════════ #

if __name__ == "__main__":
    app = ExcelMergerApp()
    app.mainloop()
