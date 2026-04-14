import tkinter as tk
import subprocess
import threading
import os
import sys
 
# ============================================================
# CONFIG — update to final_capstone.py pathing moves or if we are trying to run it on different systems. 
# ============================================================

SCRIPT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'final_capstone.py')
if sys.platform == 'win32':
    VENV_PYTHON = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.venv', 'Scripts', 'python.exe')
else:
    VENV_PYTHON = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.venv', 'bin', 'python')


PYTHON_EXEC = VENV_PYTHON if os.path.exists(VENV_PYTHON) else sys.executable
 
 
# ============================================================
# CORE LOGIC
# ============================================================
def run_script(button, status_label, log_box):
    """Run final_capstone.py in a subprocess and stream output to the log box."""
 
    def task():
        # --- UI: when it's running ---
        button.config(state='disabled', text='RUNNING...', bg='#C8A951')
        status_label.config(text='Status: Running', fg='#C8A951')
        log_box.config(state='normal')
        log_box.delete('1.0', tk.END)
        log_insert(log_box, f'▶  Starting {os.path.basename(SCRIPT_PATH)}\n\n')
 
        try:
            process = subprocess.Popen(
                [PYTHON_EXEC, SCRIPT_PATH],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1
            )
 
            for line in process.stdout:
                log_insert(log_box, line)
 
            process.wait()
 
            if process.returncode == 0:
                log_insert(log_box, '\n✓  Completed successfully.\n')
                button.config(state='normal', text='RUN SCAN', bg='#4A7C59')
                status_label.config(text='Status: Complete', fg='#4A7C59')
            else:
                log_insert(log_box, f'\n✗  Exited with code {process.returncode}\n')
                button.config(state='normal', text='RUN SCAN', bg='#A63D2F')
                status_label.config(text='Status: Error', fg='#A63D2F')
 
        except FileNotFoundError:
            log_insert(log_box, f'✗  Script not found:\n   {SCRIPT_PATH}\n')
            button.config(state='normal', text='RUN SCAN', bg='#A63D2F')
            status_label.config(text='Status: Error', fg='#A63D2F')
 
        log_box.config(state='disabled')
 
    # Run in background thread so the UI stays responsive
    threading.Thread(target=task, daemon=True).start()
 
 
def log_insert(log_box, text):
    """Thread-safe insert into the log text box."""
    log_box.after(0, lambda: _do_insert(log_box, text))
 
 
def _do_insert(log_box, text):
    log_box.config(state='normal')
    log_box.insert(tk.END, text)
    log_box.see(tk.END)
    log_box.config(state='disabled')
 
 
# ============================================================
# UI - colors, fonts, etc. 
# ============================================================
def build_ui():
    root = tk.Tk()
    root.title('PDF Bill Scanner')
    root.resizable(False, False)
    root.configure(bg='#1A1A1A')
 
    # ── Fonts ──────────────────────────────────────────────
    font_title  = ('Courier New', 13, 'bold')
    font_label  = ('Courier New', 9)
    font_button = ('Courier New', 14, 'bold')
    font_log    = ('Courier New', 9)
 
    # ── Header ─────────────────────────────────────────────
    tk.Label(
        root,
        text='PDF BILL SCANNER',
        font=font_title,
        bg='#1A1A1A',
        fg='#E8E0D0',
        pady=0
    ).pack(pady=(22, 0))
 
    tk.Label(
        root,
        text='energy data extraction utility',
        font=font_label,
        bg='#1A1A1A',
        fg='#666660'
    ).pack(pady=(2, 18))
 
    # Divider
    tk.Frame(root, bg='#333330', height=1, width=340).pack(pady=(0, 20))
 
    # ── Run Button ─────────────────────────────────────────
    button = tk.Button(
        root,
        text='RUN SCAN',
        font=font_button,
        bg='#4A7C59',
        fg='#2D4B87',
        activebackground='#3A6347',
        activeforeground='#F0EBE0',
        relief='flat',
        cursor='hand2',
        width=18,
        pady=12
    )
    button.pack(pady=(0, 14))
 
    # ── Status Label ───────────────────────────────────────
    status_label = tk.Label(
        root,
        text='Status: Idle',
        font=font_label,
        bg='#1A1A1A',
        fg='#666660'
    )
    status_label.pack(pady=(0, 14))
 
    # Divider
    tk.Frame(root, bg='#333330', height=1, width=340).pack(pady=(0, 14))
 
    # ── Log Box ────────────────────────────────────────────
    tk.Label(
        root,
        text='OUTPUT LOG',
        font=('Courier New', 8, 'bold'),
        bg='#1A1A1A',
        fg='#555550',
        anchor='w'
    ).pack(padx=24, fill='x')
 
    log_frame = tk.Frame(root, bg='#111111', bd=0)
    log_frame.pack(padx=24, pady=(4, 22), fill='both')
 
    log_box = tk.Text(
        log_frame,
        height=12,
        width=48,
        font=font_log,
        bg='#111111',
        fg='#99BF88',
        insertbackground='#99BF88',
        relief='flat',
        state='disabled',
        wrap='word',
        pady=8,
        padx=10
    )
    log_box.pack(side='left', fill='both', expand=True)
 
    scrollbar = tk.Scrollbar(log_frame, command=log_box.yview, bg='#1A1A1A', troughcolor='#1A1A1A')
    scrollbar.pack(side='right', fill='y')
    log_box.config(yscrollcommand=scrollbar.set)
 
    button.config(command=lambda: run_script(button, status_label, log_box))
 
    root.mainloop()
 
 
if __name__ == '__main__':
    build_ui()
