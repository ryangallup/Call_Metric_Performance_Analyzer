import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

DEFAULT_WEIGHTS = {
    'AHT (Minutes)': 40,
    'Phone Efficiency': 30,
    'Calls Answered': 25,
    'ACW Time (Seconds)': 5
}

BAR_COLORS = ['#1f77b4', '#ff7f0e']
ICON_PATH = "Rocket_Logo.ico"

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)

def calculate_performance(event=None):
    file_path = file_entry.get()

    if file_path:
        try:
            df = pd.read_excel(file_path, sheet_name="Results", skiprows=4, usecols="B:J")
            weights = {metric: entry.get().rstrip('%') for metric, entry in weights_entries.items()}
            weights = {metric: float(weight) for metric, weight in weights.items()}
            df['Performance Score'] = df[list(weights.keys())].dot(pd.Series(weights))
            df['Performance Score'] = df['Performance Score'].abs() / 10000
            df = df[df['Performance Score'] != 0].sort_values(by='Performance Score', ascending=False)[::-1]

            plot_window = tk.Toplevel(root)
            plot_window.title('Performance Analysis')
            plot_window.iconbitmap(ICON_PATH)

            plt.style.use('dark_background')

            fig, ax = plt.subplots(figsize=(12, 8))
            bars = ax.barh(df['Team Member'], df['Performance Score'], height=0.6,  # Adjust height for better spacing
                           color=[BAR_COLORS[i % len(BAR_COLORS)] for i in range(len(df))])
            ax.set_xlabel('Performance Score')
            ax.set_title('Team Member Phone Metric Performance Analysis')
            ax.title.set_color('white')
            ax.tick_params(axis='x', colors='white')
            ax.tick_params(axis='y', colors='white')

            for idx, bar in enumerate(bars):
                width = bar.get_width()
                ax.annotate(f'{width:.2f}', xy=(width, bar.get_y() + bar.get_height() / 2), xytext=(3, 0),
                            textcoords="offset points", ha='left', va='center', color='white')

            canvas = FigureCanvasTkAgg(fig, master=plot_window)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    else:
        messagebox.showerror("Error", "Please select an Excel file.")


def show_about():
    about_window = tk.Toplevel(root)
    about_window.title("About")
    about_window.iconbitmap(ICON_PATH)

    with open("about_info.txt", "r") as f:
        about_text = f.read()

    about_label = ttk.Label(about_window, text=about_text, wraplength=400)
    about_label.pack(padx=10, pady=10)

    # Center the about window
    center_window(about_window)


def center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry('{}x{}+{}+{}'.format(width, height, x, y))

def validate_entry(entry):
    value = entry.get().strip('%')
    if value.isdigit() or (value.startswith('-') and value[1:].isdigit()):
        entry.delete(0, tk.END)
        entry.insert(0, f"{value}%")

root = tk.Tk()
root.title("The Guy Phone Metric Performance Analysis")
root.iconbitmap(ICON_PATH)
root.after(0, lambda: center_window(root))

file_label = ttk.Label(root, text="Select Excel file:")
file_label.grid(row=0, column=0, padx=5, pady=5)

file_entry = ttk.Entry(root, width=50)
file_entry.grid(row=0, column=1, padx=5, pady=5)

browse_button = ttk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, padx=5, pady=5)

weight_label = ttk.Label(root, text="Weight Percentages")
weight_label.grid(row=1, column=1, padx=5, pady=5)

weights_entries = {}

for i, (metric, default_weight) in enumerate(DEFAULT_WEIGHTS.items()):
    ttk.Label(root, text=metric).grid(row=i + 2, column=0, padx=5, pady=5)
    weight_entry = ttk.Entry(root, justify='center')
    weight_entry.insert(0, f"{default_weight}%")
    weight_entry.grid(row=i + 2, column=1, padx=5, pady=5)
    weights_entries[metric] = weight_entry
    # Binding validation to weight entry
    weight_entry.bind('<FocusOut>', lambda event, entry=weight_entry: validate_entry(entry))

calculate_button = ttk.Button(root, text="Calculate Performance", command=calculate_performance)
calculate_button.grid(row=len(DEFAULT_WEIGHTS) + 2, columnspan=3, padx=5, pady=10)

about_button = ttk.Button(root, text="About", command=show_about)
about_button.grid(row=len(DEFAULT_WEIGHTS) + 3, columnspan=3, padx=5, pady=5)

root.bind('<Return>', calculate_performance)

for i in range(len(DEFAULT_WEIGHTS) + 4):
    root.grid_rowconfigure(i, weight=1)
for i in range(3):
    root.grid_columnconfigure(i, weight=1)

root.mainloop()
