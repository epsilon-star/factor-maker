import traceback
import tkinter as tk
from tkinter import messagebox

def show_exception_popup(exc_info):
    # Format full traceback info
    formatted = ''.join(traceback.format_exception(*exc_info))
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    messagebox.showerror("Exception Occurred", formatted)

try:
    # Example error
    x = 1 / 0

except Exception as e:
    import sys
    exc_info = sys.exc_info()  # Get exception type, value, traceback
    show_exception_popup(exc_info)
