"""
GUI wrapper that connects ai_reviewer.py and document_flagger.py to a
window-based interface. Neither of the two tool files is touched — this
file just imports them and calls their functions directly.
"""

"""
os       - used to check paths, build file lists, and create folders
sys      - lets us swap out sys.stdout so print() writes to the log panel
threading - runs the actual processing in the background so the window
            doesn't freeze while the tool is working
tkinter  - Python's built-in GUI library. tk is the core, ttk gives nicer
            looking widgets, filedialog handles browse popups, scrolledtext
            is the log panel, and messagebox shows error dialogs
datetime - used to put a timestamp on output filenames so nothing gets
           overwritten accidentally
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
from datetime import datetime

import ai_reviewer
import document_flagger


"""
TextRedirector

Both tool scripts report their progress entirely through print() statements.
Normally those would go to a terminal, but since we're running as a GUI
there's no terminal to see. This class tricks Python into sending all those
print() calls to the log panel instead.

How it works: Python's print() always writes to whatever object is stored
in sys.stdout. By replacing sys.stdout with this class before calling any
tool functions, every print() inside those functions ends up in write()
here instead of the terminal. We restore sys.stdout afterward in a
finally block so nothing else in the program is affected.

Inside write(), we briefly unlock the read-only log widget, paste the text
in, scroll to the bottom, then lock it again. update_idletasks() forces
the window to redraw right away — without it, the log would only refresh
between button clicks, which makes it look frozen during a long run.
"""

class TextRedirector:
    """Redirects print() calls into a tkinter Text widget."""

    def __init__(self, widget, status_var):
        self.widget = widget
        self.status_var = status_var

    def write(self, text):
        self.widget.configure(state="normal")
        self.widget.insert(tk.END, text)
        self.widget.see(tk.END)
        self.widget.configure(state="disabled")
        self.widget.update_idletasks()

    def flush(self):
        # Python expects stdout to have a flush() method. Nothing to do here.
        pass


"""
ContractToolsApp

The main application class. Inheriting from tk.Tk means this class IS the
window — there's no separate window object to manage. Everything lives here:
the layout, the input fields, the button handlers, and the worker threads.

Startup order matters: _build_ui() creates all the widgets first, then
_load_defaults() fills them with values. If you swapped the order, the
fields wouldn't exist yet when we tried to write into them.

_stop_event is a threading.Event — basically a thread-safe on/off flag.
The Stop button sets it to "on" and the background workers check it between
files. It gets reset to "off" at the start of every new run.
"""

class ContractToolsApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AI Contract Tools")
        self.resizable(True, True)
        self.minsize(740, 680)
        self._stop_event = threading.Event()
        self._build_ui()
        self._load_defaults()


    def _build_ui(self):
        """
        Builds the overall window layout from top to bottom:
          1. Dark blue header banner with title and subtitle
          2. Two-tab panel (one tab per tool)
          3. Row of shared action buttons (Run, Stop, Open Output, Clear Log)
          4. Scrollable dark-themed log panel
          5. Thin status bar pinned to the bottom edge

        The top-level sections use pack() because they just stack vertically.
        The individual tab contents use grid() because they need labels, fields,
        and buttons to line up in columns — pack() can't do that cleanly.

        run_btn and stop_btn are saved as instance variables so the worker
        threads can enable/disable them from the background. stop_btn starts
        disabled and only becomes clickable once a run is in progress.

        status_var is a tk.StringVar — a special string container that
        automatically updates its linked label whenever the value changes,
        with no manual refresh needed.
        """

        # Header banner
        header = tk.Frame(self, bg="#154360", pady=10)
        header.pack(fill=tk.X)
        tk.Label(
            header, text="AI Contract Tools",
            font=("Segoe UI", 16, "bold"), bg="#154360", fg="white",
        ).pack()
        tk.Label(
            header, text="Auburn University — Contract Review & Compliance Flagging",
            font=("Segoe UI", 9), bg="#154360", fg="#AED6F1",
        ).pack()

        # Two-tab panel
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.X, padx=12, pady=(10, 0))

        tab1 = ttk.Frame(notebook, padding=10)
        tab2 = ttk.Frame(notebook, padding=10)
        notebook.add(tab1, text="  AI Contract Reviewer  ")
        notebook.add(tab2, text="  Compliance Matrix Generator  ")

        self._build_reviewer_tab(tab1)
        self._build_flagger_tab(tab2)

        # Shared action buttons
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=(6, 2))

        self.run_btn = ttk.Button(
            btn_frame, text="▶  Run Selected Tool", command=self._run_selected, width=26
        )
        self.run_btn.pack(side=tk.LEFT, padx=6)

        # Stop starts disabled — it only makes sense when something is running
        self.stop_btn = ttk.Button(
            btn_frame, text="⏹  Stop", command=self._request_stop,
            width=10, state="disabled"
        )
        self.stop_btn.pack(side=tk.LEFT, padx=6)

        ttk.Button(
            btn_frame, text="Open Output Folder", command=self._open_output, width=20
        ).pack(side=tk.LEFT, padx=6)

        ttk.Button(
            btn_frame, text="Clear Log", command=self._clear_log, width=12
        ).pack(side=tk.LEFT, padx=6)

        # Save notebook so other methods can check which tab is currently open
        self.notebook = notebook

        # Log panel — dark background, read-only, auto-scrolls as text comes in
        log_frame = ttk.LabelFrame(self, text="Log", padding=6)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(4, 8))

        self.log = scrolledtext.ScrolledText(
            log_frame, state="disabled", font=("Consolas", 9),
            bg="#1C2833", fg="#ECF0F1", insertbackground="white",
            relief=tk.FLAT, wrap=tk.WORD
        )
        self.log.pack(fill=tk.BOTH, expand=True)

        # Status bar at the very bottom of the window
        self.status_var = tk.StringVar(value="Ready")
        tk.Label(
            self, textvariable=self.status_var,
            anchor=tk.W, relief=tk.SUNKEN,
            font=("Segoe UI", 8), bg="#D5D8DC", padx=6
        ).pack(fill=tk.X, side=tk.BOTTOM)


    def _build_reviewer_tab(self, parent):
        """
        Input form for the AI Contract Reviewer tool (ai_reviewer.py).
        Fields: Gemini API key, Ts&Cs Matrix file, input source, output folder.

        Layout uses a 3-column grid: labels in col 0, entry fields in col 1
        (which stretches with the window), browse buttons in col 2.

        The API key field uses show="*" to hide the key on screen. The actual
        value in api_key_var is always the real text — show="*" only affects
        what gets displayed. The "Show" checkbox calls _toggle_key_visibility
        to switch between masked and plain text.

        The mode toggle (Entire Folder vs Single File) uses two radio buttons
        that share rev_mode_var. Whichever is selected writes "folder" or "file"
        into that variable and swaps the visible row below it. Both rows sit in
        the exact same grid position (row 3) — only one is visible at a time.
        grid_remove() hides a widget without destroying it, so its saved value
        and grid position are both preserved when you switch back.
        """

        parent.columnconfigure(1, weight=1)

        # API key row
        ttk.Label(parent, text="Gemini API Key:").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.api_key_var = tk.StringVar()
        self.api_entry = ttk.Entry(parent, textvariable=self.api_key_var, show="*", width=52)
        self.api_entry.grid(row=0, column=1, sticky=tk.EW, padx=(8, 4))
        self.show_key_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            parent, text="Show", variable=self.show_key_var,
            command=self._toggle_key_visibility
        ).grid(row=0, column=2)

        # Ts&Cs matrix file row
        ttk.Label(parent, text="Ts&Cs Matrix (.xlsm):").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.matrix_var = tk.StringVar()
        ttk.Entry(parent, textvariable=self.matrix_var, width=52).grid(row=1, column=1, sticky=tk.EW, padx=(8, 4))
        ttk.Button(parent, text="Browse…", command=self._browse_matrix).grid(row=1, column=2)

        # Mode toggle — folder or single file
        self.rev_mode_var = tk.StringVar(value="folder")
        mode_frame = ttk.Frame(parent)
        mode_frame.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=(6, 2))
        ttk.Label(mode_frame, text="Process:").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Radiobutton(
            mode_frame, text="Entire Solicitations Folder", variable=self.rev_mode_var,
            value="folder", command=self._toggle_rev_mode
        ).pack(side=tk.LEFT, padx=(0, 16))
        ttk.Radiobutton(
            mode_frame, text="Single File", variable=self.rev_mode_var,
            value="file", command=self._toggle_rev_mode
        ).pack(side=tk.LEFT)

        # Solicitations folder row — shown by default
        self.rev_sol_label = ttk.Label(parent, text="Solicitations Folder:")
        self.rev_sol_label.grid(row=3, column=0, sticky=tk.W, pady=4)
        self.rev_sol_var = tk.StringVar()
        self.rev_sol_entry = ttk.Entry(parent, textvariable=self.rev_sol_var, width=52)
        self.rev_sol_entry.grid(row=3, column=1, sticky=tk.EW, padx=(8, 4))
        self.rev_sol_btn = ttk.Button(parent, text="Browse…", command=lambda: self._browse_dir(self.rev_sol_var))
        self.rev_sol_btn.grid(row=3, column=2)

        # Single file row — hidden by default, swaps into row 3 when mode = "file"
        # Must be placed in the grid before grid_remove() will work on it
        self.rev_file_label = ttk.Label(parent, text="Contract File (.docx):")
        self.rev_file_var = tk.StringVar()
        self.rev_file_entry = ttk.Entry(parent, textvariable=self.rev_file_var, width=52)
        self.rev_file_btn = ttk.Button(parent, text="Browse…", command=self._browse_rev_file)
        self.rev_file_label.grid(row=3, column=0, sticky=tk.W, pady=4)
        self.rev_file_entry.grid(row=3, column=1, sticky=tk.EW, padx=(8, 4))
        self.rev_file_btn.grid(row=3, column=2)
        self.rev_file_label.grid_remove()
        self.rev_file_entry.grid_remove()
        self.rev_file_btn.grid_remove()

        # Output folder row
        ttk.Label(parent, text="Output Folder:").grid(row=4, column=0, sticky=tk.W, pady=4)
        self.rev_out_var = tk.StringVar()
        ttk.Entry(parent, textvariable=self.rev_out_var, width=52).grid(row=4, column=1, sticky=tk.EW, padx=(8, 4))
        ttk.Button(parent, text="Browse…", command=lambda: self._browse_dir(self.rev_out_var)).grid(row=4, column=2)

        ttk.Label(
            parent,
            text="Processes .docx files — flags risky clauses and suggests redlines using Gemini AI.",
            font=("Segoe UI", 8), foreground="#555"
        ).grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=(6, 0))


    def _build_flagger_tab(self, parent):
        """
        Input form for the Compliance Matrix Generator (document_flagger.py).
        Simpler than Tab 1 — no API key needed since this tool runs offline.
        Fields: Database folder, input source (folder or single file), output folder.

        Same 3-column grid layout and same folder/file toggle pattern as Tab 1,
        but using entirely separate variables (flag_mode_var, flag_sol_var, etc.)
        so the two tabs are fully independent of each other.

        Unlike Tab 1 which only accepts .docx, this tab also accepts PDFs.
        The file picker and the validator in _start_flagger both allow both types.
        """

        parent.columnconfigure(1, weight=1)

        # Database folder row — where the FAR/DFARS clause spreadsheets live
        ttk.Label(parent, text="Database Folder:").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.db_var = tk.StringVar()
        ttk.Entry(parent, textvariable=self.db_var, width=52).grid(row=0, column=1, sticky=tk.EW, padx=(8, 4))
        ttk.Button(parent, text="Browse…", command=lambda: self._browse_dir(self.db_var)).grid(row=0, column=2)

        # Mode toggle — folder or single file
        self.flag_mode_var = tk.StringVar(value="folder")
        mode_frame = ttk.Frame(parent)
        mode_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(6, 2))
        ttk.Label(mode_frame, text="Process:").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Radiobutton(
            mode_frame, text="Entire Solicitations Folder", variable=self.flag_mode_var,
            value="folder", command=self._toggle_flag_mode
        ).pack(side=tk.LEFT, padx=(0, 16))
        ttk.Radiobutton(
            mode_frame, text="Single File", variable=self.flag_mode_var,
            value="file", command=self._toggle_flag_mode
        ).pack(side=tk.LEFT)

        # Solicitations folder row — shown by default
        self.flag_sol_label = ttk.Label(parent, text="Solicitations Folder:")
        self.flag_sol_label.grid(row=2, column=0, sticky=tk.W, pady=4)
        self.flag_sol_var = tk.StringVar()
        self.flag_sol_entry = ttk.Entry(parent, textvariable=self.flag_sol_var, width=52)
        self.flag_sol_entry.grid(row=2, column=1, sticky=tk.EW, padx=(8, 4))
        self.flag_sol_btn = ttk.Button(parent, text="Browse…", command=lambda: self._browse_dir(self.flag_sol_var))
        self.flag_sol_btn.grid(row=2, column=2)

        # Single file row — hidden by default, swaps into row 2 when mode = "file"
        self.flag_file_label = ttk.Label(parent, text="Contract File (.pdf / .docx):")
        self.flag_file_var = tk.StringVar()
        self.flag_file_entry = ttk.Entry(parent, textvariable=self.flag_file_var, width=52)
        self.flag_file_btn = ttk.Button(parent, text="Browse…", command=self._browse_flag_file)
        self.flag_file_label.grid(row=2, column=0, sticky=tk.W, pady=4)
        self.flag_file_entry.grid(row=2, column=1, sticky=tk.EW, padx=(8, 4))
        self.flag_file_btn.grid(row=2, column=2)
        self.flag_file_label.grid_remove()
        self.flag_file_entry.grid_remove()
        self.flag_file_btn.grid_remove()

        # Output folder row
        ttk.Label(parent, text="Output Folder:").grid(row=3, column=0, sticky=tk.W, pady=4)
        self.flag_out_var = tk.StringVar()
        ttk.Entry(parent, textvariable=self.flag_out_var, width=52).grid(row=3, column=1, sticky=tk.EW, padx=(8, 4))
        ttk.Button(parent, text="Browse…", command=lambda: self._browse_dir(self.flag_out_var)).grid(row=3, column=2)

        ttk.Label(
            parent,
            text="Processes .pdf and .docx files — generates a color-coded compliance matrix and highlighted document.",
            font=("Segoe UI", 8), foreground="#555"
        ).grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=(6, 0))


    def _toggle_rev_mode(self):
        """
        Called when the user clicks one of the "Process:" radio buttons on Tab 1.
        Shows the folder row and hides the file row, or vice versa.

        grid_remove() hides a widget but remembers its position. grid() with no
        arguments puts it back in exactly the same spot. This is why we don't
        need to re-specify the row/column when showing the widget again.
        """
        if self.rev_mode_var.get() == "folder":
            self.rev_sol_label.grid(); self.rev_sol_entry.grid(); self.rev_sol_btn.grid()
            self.rev_file_label.grid_remove(); self.rev_file_entry.grid_remove(); self.rev_file_btn.grid_remove()
        else:
            self.rev_file_label.grid(); self.rev_file_entry.grid(); self.rev_file_btn.grid()
            self.rev_sol_label.grid_remove(); self.rev_sol_entry.grid_remove(); self.rev_sol_btn.grid_remove()

    def _toggle_flag_mode(self):
        """
        Same as _toggle_rev_mode but for Tab 2. Completely independent —
        changing the mode on one tab has no effect on the other.
        """
        if self.flag_mode_var.get() == "folder":
            self.flag_sol_label.grid(); self.flag_sol_entry.grid(); self.flag_sol_btn.grid()
            self.flag_file_label.grid_remove(); self.flag_file_entry.grid_remove(); self.flag_file_btn.grid_remove()
        else:
            self.flag_file_label.grid(); self.flag_file_entry.grid(); self.flag_file_btn.grid()
            self.flag_sol_label.grid_remove(); self.flag_sol_entry.grid_remove(); self.flag_sol_btn.grid_remove()


    def _load_defaults(self):
        """
        Fills in all the input fields when the app first opens, using the same
        folder paths that are already hardcoded in ai_reviewer.py and
        document_flagger.py. This way the user doesn't have to type anything
        for a standard run — the fields just match what the scripts expect.

        When Python imports a module, the top-level code runs immediately, so
        things like ai_reviewer.MATRIX_FILE are already set by the time we get
        here. StringVar.set() pushes a value into the field linked to that variable.
        """
        # Tab 1 — from ai_reviewer.py
        self.api_key_var.set(ai_reviewer.GEMINI_API_KEY)
        self.matrix_var.set(ai_reviewer.MATRIX_FILE)
        self.rev_sol_var.set(ai_reviewer.SOLICITATIONS_DIR)
        self.rev_out_var.set(ai_reviewer.OUTPUT_DIR)

        # Tab 2 — from document_flagger.py
        self.db_var.set(document_flagger.DATABASE_DIR)
        self.flag_sol_var.set(document_flagger.SOLICITATIONS_DIR)
        self.flag_out_var.set(document_flagger.OUTPUT_DIR)


    def _browse_matrix(self):
        """Opens a file picker filtered to Excel formats for the matrix field."""
        path = filedialog.askopenfilename(
            title="Select Ts&Cs Matrix",
            filetypes=[("Excel Macro-Enabled", "*.xlsm"), ("Excel", "*.xlsx *.xls"), ("All", "*.*")]
        )
        if path:
            self.matrix_var.set(path)

    def _browse_rev_file(self):
        """Tab 1 single-file picker — .docx only since ai_reviewer can't handle PDFs."""
        path = filedialog.askopenfilename(
            title="Select Contract File",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if path:
            self.rev_file_var.set(path)

    def _browse_flag_file(self):
        """Tab 2 single-file picker — allows both PDF and DOCX."""
        path = filedialog.askopenfilename(
            title="Select Contract File",
            filetypes=[("Supported Files", "*.pdf *.docx"), ("PDF", "*.pdf"), ("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if path:
            self.flag_file_var.set(path)

    def _browse_dir(self, string_var):
        """Generic folder picker reused by every directory field in both tabs."""
        path = filedialog.askdirectory()
        if path:
            string_var.set(path)

    def _toggle_key_visibility(self):
        """Flips the API key entry between masked (show='*') and visible (show='')."""
        self.api_entry.configure(show="" if self.show_key_var.get() else "*")

    def _open_output(self):
        """Opens the output folder for whichever tab is currently active."""
        tab_idx = self.notebook.index(self.notebook.select())
        folder = self.rev_out_var.get() if tab_idx == 0 else self.flag_out_var.get()
        if os.path.isdir(folder):
            os.startfile(folder)
        else:
            messagebox.showwarning("Not Found", f"Output folder does not exist yet:\n{folder}")

    def _clear_log(self):
        """Briefly unlocks the read-only log, wipes all text, then locks it again."""
        self.log.configure(state="normal")
        self.log.delete("1.0", tk.END)
        self.log.configure(state="disabled")

    def _request_stop(self):
        """
        Called when the user clicks the Stop button mid-run.
        Sets the stop flag to on — the background worker checks this between
        files and exits cleanly when it sees it. The current file always finishes
        before stopping, so nothing gets left half-written.
        """
        self._stop_event.set()
        self.stop_btn.configure(state="disabled")
        self.status_var.set("Stopping after current file…")


    def _run_selected(self):
        """Checks which tab is open and hands off to the right launcher."""
        tab_idx = self.notebook.index(self.notebook.select())
        if tab_idx == 0:
            self._start_reviewer()
        else:
            self._start_flagger()

    def _start_reviewer(self):
        """
        Reads and validates the Tab 1 form fields, then starts the reviewer
        worker on a background thread if everything looks good.

        Validation runs here on the main thread because messagebox dialogs
        can crash tkinter if called from a background thread. Once validation
        passes, _begin_run() preps the UI and the worker thread is started
        with daemon=True so it dies automatically if the window is closed.
        """
        api_key     = self.api_key_var.get().strip()
        matrix_path = self.matrix_var.get().strip()
        out_dir     = self.rev_out_var.get().strip()
        mode        = self.rev_mode_var.get()

        if not api_key or api_key in ("YOUR_API_KEY_HERE", "API KEY HERE"):
            messagebox.showerror("Missing API Key", "Please enter your Gemini API key.")
            return
        if not os.path.isfile(matrix_path):
            messagebox.showerror("File Not Found", f"Matrix file not found:\n{matrix_path}")
            return

        if mode == "folder":
            sol_dir = self.rev_sol_var.get().strip()
            if not os.path.isdir(sol_dir):
                messagebox.showerror("Folder Not Found", f"Solicitations folder not found:\n{sol_dir}")
                return
            target = sol_dir
        else:
            file_path = self.rev_file_var.get().strip()
            if not os.path.isfile(file_path):
                messagebox.showerror("File Not Found", f"Contract file not found:\n{file_path}")
                return
            if not file_path.lower().endswith(".docx"):
                messagebox.showerror("Wrong File Type", "The AI Reviewer only supports .docx files.")
                return
            target = file_path

        self._begin_run()
        threading.Thread(
            target=self._reviewer_worker,
            args=(api_key, matrix_path, target, out_dir, mode),
            daemon=True,
        ).start()

    def _start_flagger(self):
        """
        Reads and validates the Tab 2 form fields, then starts the flagger
        worker on a background thread if everything looks good.
        """
        db_dir  = self.db_var.get().strip()
        out_dir = self.flag_out_var.get().strip()
        mode    = self.flag_mode_var.get()

        if not os.path.isdir(db_dir):
            messagebox.showerror("Folder Not Found", f"Database folder not found:\n{db_dir}")
            return

        if mode == "folder":
            sol_dir = self.flag_sol_var.get().strip()
            if not os.path.isdir(sol_dir):
                messagebox.showerror("Folder Not Found", f"Solicitations folder not found:\n{sol_dir}")
                return
            target = sol_dir
        else:
            file_path = self.flag_file_var.get().strip()
            if not os.path.isfile(file_path):
                messagebox.showerror("File Not Found", f"Contract file not found:\n{file_path}")
                return
            if not file_path.lower().endswith(('.pdf', '.docx')):
                messagebox.showerror("Wrong File Type", "Please select a .pdf or .docx file.")
                return
            target = file_path

        self._begin_run()
        threading.Thread(
            target=self._flagger_worker,
            args=(db_dir, target, out_dir, mode),
            daemon=True,
        ).start()

    def _begin_run(self):
        """Preps the UI when a run starts — disables Run, enables Stop, clears the log."""
        self._stop_event.clear()  # reset from any previous stop
        self.run_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.status_var.set("Running… please wait")
        self._clear_log()

    def _end_run(self, status_msg):
        """Resets the UI when a run finishes — re-enables Run, disables Stop."""
        self.run_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.status_var.set(status_msg)


    def _validate_api_key(self, api_key):
        """
        Sends a tiny test message to Gemini before doing anything real. If the
        key is rejected, we catch the error and print a clear message right away
        rather than letting the user wait through the whole playbook-building
        step before finding out their key doesn't work.

        The error message check distinguishes between two failure types:
          - Auth errors (bad/expired key): contain words like "invalid", "401", "403"
          - Network errors (no internet, API down): everything else
        That way the user knows whether to fix their key or check their connection.
        """
        print("Validating API key…")
        try:
            from google import genai
            client = genai.Client(api_key=api_key)
            client.models.generate_content(
                model='gemini-2.5-flash',
                contents=["Reply with the single word OK."]
            )
            print("API key accepted.\n")
            return True
        except Exception as e:
            error_str = str(e).lower()
            if any(word in error_str for word in ("api key", "api_key", "invalid", "permission", "credential", "401", "403", "unauthenticated")):
                print("[!] API key is not valid or not authorized.")
                print(f"    Detail: {e}")
            else:
                print(f"[!] Could not reach the Gemini API: {e}")
            return False


    def _reviewer_worker(self, api_key, matrix_path, target, out_dir, mode):
        """
        The actual processing for Tab 1, running on a background thread.
        Steps: validate API key → build playbook → find contracts → review each one.

        stdout redirection: sys.stdout is swapped for a TextRedirector at the start
        and restored in the finally block. This means every print() inside this
        function AND inside the ai_reviewer functions it calls will show up in the
        log panel. The finally block ensures stdout always gets restored even if
        something crashes partway through.

        Runtime patching: we overwrite the module-level variables in ai_reviewer
        (e.g. ai_reviewer.GEMINI_API_KEY = api_key) so the functions inside that
        module pick up the values the user entered in the GUI. Python modules are
        objects, so their variables can be changed from outside just like any other
        object's attributes — no need to touch the source file.

        File list: in folder mode we scan the folder for .docx files; in single
        file mode we just wrap the one path in a list. Either way the loop below
        is identical — no special cases needed.

        Stop check: self._stop_event.is_set() is checked at the top of every loop
        iteration. If the user clicked Stop, we exit before starting the next file.
        The file currently being processed always finishes first.
        """
        old_stdout = sys.stdout
        sys.stdout = TextRedirector(self.log, self.status_var)
        try:
            ai_reviewer.GEMINI_API_KEY = api_key
            ai_reviewer.MATRIX_FILE    = matrix_path
            ai_reviewer.OUTPUT_DIR     = out_dir

            print("\n--- Starting AI Semantic Contract Reviewer ---")
            os.makedirs(out_dir, exist_ok=True)

            # Bail out early if the key is wrong — no point reading anything else
            if not self._validate_api_key(api_key):
                self._end_run("Stopped — invalid API key.")
                return

            # Parse the Excel matrix into text the AI can understand
            playbook_text = ai_reviewer.build_ai_playbook(matrix_path)
            if not playbook_text:
                print("[!] Could not build playbook. Aborting.")
                self._end_run("Failed — see log.")
                return

            # Build the file list — either a folder scan or a single file in a list
            if mode == "folder":
                sol_dir = target
                os.makedirs(sol_dir, exist_ok=True)
                doc_files = [
                    os.path.join(sol_dir, f)
                    for f in os.listdir(sol_dir)
                    if f.lower().endswith(".docx")
                ]
                if not doc_files:
                    print(f"\nNo .docx contracts found in:\n  {sol_dir}")
                    self._end_run("No contracts found.")
                    return
            else:
                doc_files = [target]

            for file_path in doc_files:
                if self._stop_event.is_set():
                    print("\n[!] Stopped by user.")
                    self._end_run("Stopped by user.")
                    return

                file_name = os.path.basename(file_path)
                print(f"\n--- Processing: {file_name} ---")
                contract_text = ai_reviewer.extract_text_from_docx(file_path)
                if not contract_text:
                    continue

                ai_report = ai_reviewer.review_contract_with_ai(contract_text, playbook_text)
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                report_filename = f"AI_Risk_Report_{os.path.splitext(file_name)[0]}_{timestamp}.docx"
                ai_reviewer.save_report_to_word(ai_report, os.path.join(out_dir, report_filename))

            print("\n--- All tasks completed. ---")
            self._end_run("Done — check the Output folder.")

        except Exception as exc:
            print(f"\n[!] Unexpected error: {exc}")
            self._end_run("Error — see log.")
        finally:
            sys.stdout = old_stdout
            self.run_btn.configure(state="normal")
            self.stop_btn.configure(state="disabled")


    def _flagger_worker(self, db_dir, target, out_dir, mode):
        """
        The actual processing for Tab 2, running on a background thread.
        Steps: load clause database → find contracts → scan each one → save outputs.

        Same stdout redirection, runtime patching, and stop check pattern as
        _reviewer_worker — see those notes above for how those parts work.

        Per-file steps:
          1. Extract text from the contract (PDF or DOCX, with OCR fallback for scans)
          2. Search the text for every known federal clause number in the database
          3. Look up each found clause in the master dataframe to get its full details
          4. Save a color-coded Excel spreadsheet (green/yellow/red by compliance status)
          5. Save a highlighted copy of the original document with clause numbers marked
        """
        old_stdout = sys.stdout
        sys.stdout = TextRedirector(self.log, self.status_var)
        try:
            document_flagger.DATABASE_DIR = db_dir
            document_flagger.OUTPUT_DIR   = out_dir

            print("\n--- Starting Automated Compliance Matrix Generator ---")
            os.makedirs(db_dir, exist_ok=True)
            os.makedirs(out_dir, exist_ok=True)

            # Load and merge all agency clause spreadsheets into one big table
            master_df = document_flagger.load_databases(db_dir)
            if master_df is None:
                print("[!] Could not load databases. Aborting.")
                self._end_run("Failed — see log.")
                return

            known_clauses = master_df[document_flagger.CLAUSE_COL_NAME].unique().tolist()

            # Build the file list — either a folder scan or a single file in a list
            if mode == "folder":
                sol_dir = target
                os.makedirs(sol_dir, exist_ok=True)
                doc_files = [
                    os.path.join(sol_dir, f)
                    for f in os.listdir(sol_dir)
                    if f.lower().endswith(('.pdf', '.docx'))
                ]
                if not doc_files:
                    print(f"\nNo PDF or DOCX files found in:\n  {sol_dir}")
                    self._end_run("No contracts found.")
                    return
            else:
                doc_files = [target]

            for file_path in doc_files:
                if self._stop_event.is_set():
                    print("\n[!] Stopped by user.")
                    self._end_run("Stopped by user.")
                    return

                file_name = os.path.basename(file_path)
                print(f"\n--- Processing: {file_name} ---")

                if file_name.lower().endswith('.pdf'):
                    text = document_flagger.extract_text_from_pdf(file_path)
                else:
                    text = document_flagger.extract_text_from_docx(file_path)

                if not text:
                    continue

                found_clauses = document_flagger.find_clauses_from_db(text, known_clauses)
                print(f"Found {len(found_clauses)} unique federal clauses.")

                if not found_clauses:
                    print("Warning: No matching clauses found. Skipping file.")
                    continue

                compliance_df = document_flagger.generate_compliance_matrix(found_clauses, master_df)
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                base = os.path.splitext(file_name)[0]

                # Save the Excel matrix, then paint the green/yellow/red colors on top
                excel_path = os.path.join(out_dir, f"Compliance_Matrix_{base}_{timestamp}.xlsx")
                try:
                    compliance_df.to_excel(excel_path, index=False)
                    document_flagger.apply_color_coding(excel_path)
                    print(f"  -> Saved compliance matrix.")
                except Exception as e:
                    print(f"  ! Error saving Excel: {e}")

                # Save a highlighted copy of the original document
                if file_name.lower().endswith('.pdf'):
                    out_path = os.path.join(out_dir, f"Executed_Highlights_{base}_{timestamp}.pdf")
                    document_flagger.highlight_pdf(file_path, out_path, found_clauses)
                else:
                    out_path = os.path.join(out_dir, f"Executed_Highlights_{base}_{timestamp}.docx")
                    document_flagger.highlight_docx(file_path, out_path, found_clauses)

            print("\n--- All tasks completed. ---")
            self._end_run("Done — check the Output folder.")

        except Exception as exc:
            print(f"\n[!] Unexpected error: {exc}")
            self._end_run("Error — see log.")
        finally:
            sys.stdout = old_stdout
            self.run_btn.configure(state="normal")
            self.stop_btn.configure(state="disabled")


"""
Only runs when this file is executed directly (python gui.py or pythonw gui.py).
If someone were to import gui as a module, this block is skipped so the window
doesn't pop open unexpectedly.
"""
if __name__ == "__main__":
    app = ContractToolsApp()
    app.mainloop()
