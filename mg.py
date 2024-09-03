import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys

# Example processing functions (replace with your own imports)
from pi import process_pi
from pi_payment import process_pi_payment
from dn_country_seperate import process_dn_seperate
from ws_vin_list import process_ws_vin_list
from cipl import process_cipl


class DataExtractorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MG Data Extractor")
        self.geometry("1000x500")

        # Apply the color scheme
        self.bg_color = "#BBE1FA"  # Light Blue
        self.button_color = "#3282B8"  # Bright Blue
        self.button_active_color = "#0F4C75"  # Dark Blue
        self.text_color = "#1B262C"  # Dark Blue-Grey
        self.selected_button = None  # To track the selected button

        self.configure(bg=self.bg_color)
        self.frames = {}

        # Determine the path to the icon file
        if hasattr(sys, '_MEIPASS'):
            # If running as a PyInstaller bundle
            icon_path = os.path.join(sys._MEIPASS, "mg.ico")
        else:
            # If running as a script
            icon_path = os.path.join(os.path.dirname(__file__), "mg.ico")

        # Set the window icon
        self.iconbitmap(icon_path)

        self.create_widgets()

    def create_widgets(self):
        # Create method selection buttons with improved styling
        button_frame = tk.Frame(self, bg=self.bg_color)
        button_frame.pack(pady=20)

        self.buttons = {}

        buttons = [
            ("PI WS Tracker", "PI"),
            ("PI Payment Tracker", "PI Payment"),
            ("CIPL Data Extractor", "CIPL"),
            ("DN WS VIN Extractor", "WS VIN"),
            ("DN Country Seperator", "DN Country")
        ]

        for text, frame_name in buttons:
            button = tk.Button(button_frame, text=text,
                               command=lambda name=frame_name, btn_name=text: self.show_frame(name, btn_name),
                               bg=self.button_color, fg=self.text_color,
                               activebackground=self.button_active_color,
                               font=("Helvetica", 12), relief="flat", padx=10, pady=5)
            button.pack(side=tk.LEFT, padx=10)
            self.buttons[text] = button

        # Create frames for each method
        self.frames["PI"] = self.create_tab_frame("PI", process_pi)
        self.frames["PI Payment"] = self.create_tab_frame("PI Payment", process_pi_payment)
        self.frames["WS VIN"] = self.create_tab_frame("WS VIN", process_ws_vin_list)
        self.frames["CIPL"] = self.create_tab_frame("CIPL", process_cipl)
        self.frames["DN Country"] = self.create_dn_country_frame()

        # Show the "PI" frame by default
        self.show_frame("PI", "PI WS Tracker")

    def create_tab_frame(self, name, process_function):
        frame = tk.Frame(self, bg=self.bg_color)
        frame.pack_forget()  # Ensure it's not packed at creation

        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        file_listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, yscrollcommand=scrollbar.set,
                                  bg="white", fg=self.text_color, font=("Helvetica", 10))
        file_listbox.pack(side=tk.LEFT, fill='both', expand=True, padx=10, pady=10)

        scrollbar.config(command=file_listbox.yview)

        def browse_files():
            files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
            if files:
                file_listbox.delete(0, tk.END)
                for file in files:
                    file_listbox.insert(tk.END, file)

        def save_file():
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                     filetypes=[("Excel files", "*.xlsx")],
                                                     initialfile=f"{name}_extracted_data.xlsx")
            if save_path:
                files = file_listbox.get(0, tk.END)
                process_function(files, save_path)

        button_frame = tk.Frame(frame, bg=self.bg_color)
        button_frame.pack(pady=10)

        select_files_button = tk.Button(button_frame, text="Select Files", command=browse_files,
                                        bg=self.button_color, fg=self.text_color,
                                        activebackground=self.button_active_color,
                                        font=("Helvetica", 12), relief="flat", padx=10, pady=5)
        select_files_button.pack(side=tk.LEFT, padx=10)

        save_as_button = tk.Button(button_frame, text="Save As", command=save_file,
                                   bg=self.button_color, fg=self.text_color,
                                   activebackground=self.button_active_color,
                                   font=("Helvetica", 12), relief="flat", padx=10, pady=5)
        save_as_button.pack(side=tk.LEFT, padx=10)

        return frame

    def create_dn_country_frame(self):
        frame = tk.Frame(self, bg=self.bg_color)
        frame.pack_forget()

        frame_file = tk.Frame(frame, bg=self.bg_color)
        frame_file.pack(pady=10, padx=10, fill='x')

        tk.Label(frame_file, text="Select File:", bg=self.bg_color, fg=self.text_color,
                 font=("Helvetica", 12)).pack(side=tk.LEFT)
        self.entry_file = tk.Entry(frame_file, width=40, bg="white", fg=self.text_color,
                                   font=("Helvetica", 10))
        self.entry_file.pack(side=tk.LEFT, padx=5)
        tk.Button(frame_file, text="Browse", command=self.browse_file,
                  bg=self.button_color, fg=self.text_color,
                  activebackground=self.button_active_color,
                  font=("Helvetica", 10), relief="flat").pack(side=tk.LEFT)

        frame_output_dir = tk.Frame(frame, bg=self.bg_color)
        frame_output_dir.pack(pady=10, padx=10, fill='x')

        tk.Label(frame_output_dir, text="Output Directory:", bg=self.bg_color, fg=self.text_color,
                 font=("Helvetica", 12)).pack(side=tk.LEFT)
        self.entry_output_dir = tk.Entry(frame_output_dir, width=40, bg="white", fg=self.text_color,
                                         font=("Helvetica", 10))
        self.entry_output_dir.pack(side=tk.LEFT, padx=5)
        tk.Button(frame_output_dir, text="Browse", command=self.browse_output_dir,
                  bg=self.button_color, fg=self.text_color,
                  activebackground=self.button_active_color,
                  font=("Helvetica", 10), relief="flat").pack(side=tk.LEFT)

        tk.Button(frame, text="Process", command=self.run_processing,
                  bg=self.button_color, fg=self.text_color,
                  activebackground=self.button_active_color,
                  font=("Helvetica", 12), relief="flat", padx=10, pady=5).pack(pady=20)

        return frame

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, file_path)

    def browse_output_dir(self):
        output_base_dir = filedialog.askdirectory()
        if output_base_dir:
            self.entry_output_dir.delete(0, tk.END)
            self.entry_output_dir.insert(0, output_base_dir)

    def run_processing(self):
        file_path = self.entry_file.get()
        output_base_dir = self.entry_output_dir.get()
        if not file_path:
            messagebox.showwarning("Input Required", "Please specify an input file.")
            return
        if not output_base_dir:
            messagebox.showwarning("Output Directory Required", "Please specify an output directory.")
            return
        try:
            process_dn_seperate(file_path, output_base_dir)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def show_frame(self, name, btn_name):
        # Hide all frames
        for frame in self.frames.values():
            frame.pack_forget()

        # Show the selected frame
        self.frames[name].pack(pady=10, padx=10, fill='both', expand=True)

        # Highlight the selected button
        if self.selected_button:
            self.selected_button.config(bg=self.button_color)  # Reset previous button color
        self.selected_button = self.buttons[btn_name]
        self.selected_button.config(bg=self.button_active_color)  # Highlight selected button


if __name__ == "__main__":
    app = DataExtractorApp()
    app.mainloop()
