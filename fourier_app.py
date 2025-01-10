import tkinter as tk
from tkinter import ttk, messagebox
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import io
import threading


class LoadingWindow:
    def __init__(self, parent):
        self.top = tk.Toplevel(parent)
        self.top.title("Processing")
        self.top.geometry("300x150")
        self.top.transient(parent)
        self.top.grab_set()

        self.center_window()

        frame = ttk.Frame(self.top, padding="20")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.progress = ttk.Progressbar(frame, mode='indeterminate', length=200)
        self.progress.grid(row=0, column=0, pady=20)

        self.label = ttk.Label(frame, text="Generating results...")
        self.label.grid(row=1, column=0)

        self.progress.start(10)

    def center_window(self):
        self.top.update_idletasks()
        width = self.top.winfo_width()
        height = self.top.winfo_height()
        x = (self.top.winfo_screenwidth() // 2) - (width // 2)
        y = (self.top.winfo_screenheight() // 2) - (height // 2)
        self.top.geometry(f'+{x}+{y}')

    def destroy(self):
        self.top.destroy()


class SaveOptionsWindow:
    def __init__(self, parent, df, figures):
        self.top = tk.Toplevel(parent)
        self.top.title("Save Options")
        self.top.geometry("300x200")
        self.top.transient(parent)
        self.top.grab_set()

        self.center_window()

        self.df = df
        self.figures = figures

        frame = ttk.Frame(self.top, padding="20")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        label = ttk.Label(frame, text="Choose save format:", font=('Arial', 12))
        label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        excel_btn = ttk.Button(frame, text="Save as Excel",
                               command=lambda: self.save_file('excel'))
        excel_btn.grid(row=1, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

        csv_btn = ttk.Button(frame, text="Save as CSV",
                             command=lambda: self.save_file('csv'))
        csv_btn.grid(row=2, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

        close_btn = ttk.Button(frame, text="Close", command=self.top.destroy)
        close_btn.grid(row=3, column=0, columnspan=2, pady=(20, 0), sticky=(tk.W, tk.E))

    def center_window(self):
        self.top.update_idletasks()
        width = self.top.winfo_width()
        height = self.top.winfo_height()
        x = (self.top.winfo_screenwidth() // 2) - (width // 2)
        y = (self.top.winfo_screenheight() // 2) - (height // 2)
        self.top.geometry(f'+{x}+{y}')

    def save_file(self, file_type):
        file_types = {
            'excel': [('Excel files', '*.xlsx')],
            'csv': [('CSV files', '*.csv')]
        }
        default_ext = '.xlsx' if file_type == 'excel' else '.csv'

        file_path = tk.filedialog.asksaveasfilename(
            defaultextension=default_ext,
            filetypes=file_types[file_type],
            title=f"Save as {file_type.upper()}"
        )

        if not file_path:
            return

        try:
            if file_type == 'excel':
                self.save_excel(file_path)
            else:
                self.save_csv(file_path)
            messagebox.showinfo("Success", "File saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving file: {str(e)}")

    def save_excel(self, filename):
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            self.df.to_excel(writer, sheet_name="Peaks", index=False)
            workbook = writer.book

            sheet_names = ["X1 Signal", "X2 Signal", "Combined Signal", "DFT Spectrum"]
            for fig_buf, sheet_name in zip(self.figures, sheet_names):
                sheet = workbook.create_sheet(title=sheet_name)
                img = Image(fig_buf)
                sheet.add_image(img, "A1")

    def save_csv(self, filename):
        self.df.to_csv(filename, index=False)


class FourierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Discrete Fourier Transformation")

        self.frame = ttk.Frame(root, padding="10")
        self.frame.grid(column=0, row=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.fields = {
            "N": "5000",
            "FR (Sampling Rate)": "1000",
            "Delta T": "0.001",
            "a1": "1",
            "a2": "0.7",
            "f1": "5",
            "f2": "10",
            "Phase Shift 1": "0",
            "Phase Shift 2": "0",
        }

        self.entries = {}
        for idx, (label, default) in enumerate(self.fields.items()):
            ttk.Label(self.frame, text=label).grid(column=0, row=idx, sticky=tk.W)
            entry = ttk.Entry(self.frame, width=10)
            entry.grid(column=1, row=idx, sticky=(tk.W, tk.E))
            entry.insert(0, default)
            self.entries[label] = entry

        self.button_frame = ttk.Frame(self.frame)
        self.button_frame.grid(column=0, row=len(self.fields), columnspan=2, sticky=(tk.W, tk.E))

        self.close_btn = ttk.Button(self.button_frame, text="Close", command=self.close_app)
        self.close_btn.pack(side=tk.LEFT, padx=5)

        self.generate_btn = ttk.Button(self.button_frame, text="Generate Signal",
                                       command=self.start_generation)
        self.generate_btn.pack(side=tk.RIGHT, padx=5)

        for child in self.frame.winfo_children():
            child.grid_configure(padx=5, pady=5)

    def close_app(self):
        self.root.quit()

    def generate_signal(self):
        N = int(self.entries["N"].get())
        FR = int(self.entries["FR (Sampling Rate)"].get())
        a1 = float(self.entries["a1"].get())
        a2 = float(self.entries["a2"].get())
        f1 = float(self.entries["f1"].get())
        f2 = float(self.entries["f2"].get())
        phase1 = float(self.entries["Phase Shift 1"].get())
        phase2 = float(self.entries["Phase Shift 2"].get())

        t = np.linspace(0, (N - 1) / FR, N)

        X1 = a1 * np.sin(2 * np.pi * f1 * t + phase1)
        X2 = a2 * np.sin(2 * np.pi * f2 * t + phase2)
        X_combined = X1 + X2

        X_padded = np.concatenate((X_combined, np.zeros(N)))
        t_padded = np.linspace(0, (2 * N - 1) / FR, 2 * N)

        def DFT(x):
            N = len(x)
            n = np.arange(N)
            k = n.reshape((N, 1))
            e = np.exp(-2j * np.pi * k * n / N)
            return np.dot(e, x)

        X_dft_padded = DFT(X_padded)
        freqs_padded = np.arange(2 * N) * (FR / (2 * N))

        magnitude = np.abs(X_dft_padded)
        half_N = len(magnitude) // 2
        peak_indices = np.argsort(magnitude[:half_N])[-2:][::-1]

        peak_data = {
            "N": N, "FR": FR, "f1": f1, "f2": f2,
            "fM-1": freqs_padded[peak_indices[0] - 1],
            "fM": freqs_padded[peak_indices[0]],
            "fM+1": freqs_padded[peak_indices[0] + 1],
            "AM-1": magnitude[peak_indices[0] - 1],
            "AM": magnitude[peak_indices[0]],
            "AM+1": magnitude[peak_indices[0] + 1],
            "fm-1": freqs_padded[peak_indices[1] - 1],
            "fm": freqs_padded[peak_indices[1]],
            "fm+1": freqs_padded[peak_indices[1] + 1],
            "Am-1": magnitude[peak_indices[1] - 1],
            "Am": magnitude[peak_indices[1]],
            "Am+1": magnitude[peak_indices[1] + 1],
        }

        df = pd.DataFrame([peak_data])

        figure_buffers = []

        fig1 = plt.figure(figsize=(12, 6))
        plt.plot(t, X1, label=f'X1(t) - f1={f1}Hz, a1={a1}', color='b')
        plt.title('Harmonic Signal X1(t)')
        plt.xlabel('Time (s)')
        plt.ylabel('Amplitude')
        plt.legend()
        buf1 = io.BytesIO()
        fig1.savefig(buf1, format="png")
        buf1.seek(0)
        figure_buffers.append(buf1)
        plt.close(fig1)

        fig2 = plt.figure(figsize=(12, 6))
        plt.plot(t, X2, label=f'X2(t) - f2={f2}Hz, a2={a2}', color='g')
        plt.title('Harmonic Signal X2(t)')
        plt.xlabel('Time (s)')
        plt.ylabel('Amplitude')
        plt.legend()
        buf2 = io.BytesIO()
        fig2.savefig(buf2, format="png")
        buf2.seek(0)
        figure_buffers.append(buf2)
        plt.close(fig2)

        fig3 = plt.figure(figsize=(12, 6))
        plt.plot(t, X_combined, label='Combined Signal X(t)', color='r')
        plt.title('Combined Signal X(t) = X1(t) + X2(t)')
        plt.xlabel('Time (s)')
        plt.ylabel('Amplitude')
        plt.legend()
        buf3 = io.BytesIO()
        fig3.savefig(buf3, format="png")
        buf3.seek(0)
        figure_buffers.append(buf3)
        plt.close(fig3)

        fig4 = plt.figure(figsize=(12, 6))
        plt.stem(freqs_padded[:N], np.abs(X_dft_padded)[:N], 'b', markerfmt=" ", basefmt="-b")
        plt.title('DFT Magnitude Spectrum (Zero-Padded)')
        plt.xlabel('Frequency (Hz)')
        plt.ylabel('|X(freq)|')
        plt.xlim(0, max(f1, f2) * 2)
        buf4 = io.BytesIO()
        fig4.savefig(buf4, format="png")
        buf4.seek(0)
        figure_buffers.append(buf4)
        plt.close(fig4)

        return df, figure_buffers

    def start_generation(self):
        self.loading = LoadingWindow(self.root)
        thread = threading.Thread(target=self.process_generation)
        thread.start()

    def process_generation(self):
        try:
            df, figures = self.generate_signal()
            self.root.after(0, self.complete_generation, df, figures)
        except Exception as e:
            self.root.after(0, self.show_error, str(e))

    def complete_generation(self, df, figures):
        self.loading.destroy()
        SaveOptionsWindow(self.root, df, figures)

    def show_error(self, error_message):
        self.loading.destroy()
        messagebox.showerror("Error", f"An error occurred: {error_message}")


if __name__ == "__main__":
    root = tk.Tk()
    app = FourierApp(root)
    root.mainloop()
