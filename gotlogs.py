import tkinter as tk
from tkinter import scrolledtext, simpledialog, messagebox, filedialog, Label, StringVar, OptionMenu, Menu
from tkinter.font import Font
import subprocess
import sys
import os

class LogViewer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("GOTLOGS Log Viewer")
        self.geometry("800x600")

        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(__file__)
        
        icon_path = os.path.join(base_path, 'log.ico')
        self.iconbitmap(icon_path)

        self.button_font = Font(family="Helvetica", size=10, weight="bold")
        self.current_font = Font(family="Helvetica", size=10)

        self.log_area = scrolledtext.ScrolledText(self, state='disabled', wrap=tk.WORD, font=self.current_font)
        self.log_area.pack(fill=tk.BOTH, expand=True)

        self.auto_refresh = True
        self.highlight_words = {"ERROR": "red", "WARNING": "yellow", "SUCCESSFUL": "light green"}

        self.controls_frame = tk.Frame(self)
        self.controls_frame.pack(fill=tk.X)

        self.create_controls()
        self.create_menu()

        self.filepath_label = Label(self.controls_frame, text="", fg="blue")
        self.filepath_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.log_file_path = None

    def create_controls(self):
        self.pause_button = tk.Button(self.controls_frame, text="Pause Refresh", command=self.toggle_refresh, font=self.button_font, bg="lightblue", fg="black")
        self.pause_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.refresh_button = tk.Button(self.controls_frame, text="Refresh Now", command=self.manual_refresh, font=self.button_font, bg="lightgreen", fg="black")
        self.refresh_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.search_button = tk.Button(self.controls_frame, text="Search", command=self.search_in_logs, font=self.button_font, bg="lightgrey", fg="black")
        self.search_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.file_button = tk.Button(self.controls_frame, text="Select Log File", command=self.select_log_file, font=self.button_font, bg="lightyellow", fg="black")
        self.file_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.restart_button = tk.Button(self.controls_frame, text="Restart Service", command=self.prompt_restart_service, font=self.button_font, bg="orange", fg="black")
        self.restart_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.font_var = StringVar(self)
        self.font_options = ["Helvetica", "Tahoma", "Segoe UI", "Arial", "Courier", "Times"]
        self.font_var.set(self.font_options[0])  # default value
        self.font_dropdown = OptionMenu(self.controls_frame, self.font_var, *self.font_options, command=self.change_font)
        self.font_dropdown.pack(side=tk.LEFT, padx=5, pady=5)

    def create_menu(self):
        menu_bar = Menu(self)
        self.config(menu=menu_bar)

        help_menu = Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="Usage Instructions", command=self.show_usage)

    def show_about(self):
        messagebox.showinfo("About", "GOTLOGS LogViewer\nVersion .7\nDeveloped in Python by: TL")

    def show_usage(self):
        usage_text = (
            "Usage Instructions:\n"
            "- To pause the automatic log refresh, click 'Pause Refresh'.\n"
            "- To manually refresh the logs, click 'Refresh Now'.\n"
            "- To search within the logs, pause refresh, click 'Search' and enter the search term.\n"
            "- To restart a service, click 'Restart Service' and enter the service name. You must launch this tool as an admin.\n"
            "- To change the display font, select from the dropdown menu in the lower right.\n"
            "- To select a new log file, click 'Select Log File'."
        )
        messagebox.showinfo("Usage Instructions", usage_text)

    def toggle_refresh(self):
        self.auto_refresh = not self.auto_refresh
        self.pause_button.config(text="Resume Refresh" if not self.auto_refresh else "Pause Refresh")
        if self.auto_refresh and self.log_file_path:
            self.restart_log_view()

    def manual_refresh(self):
        if self.log_file_path:
            self.update_log_view(force=True)

    def update_log_view(self, force=False):
        if (self.auto_refresh or force) and self.log_file_path:
            try:
                with open(self.log_file_path, 'r') as file:
                    lines = file.readlines()
                    self.log_area.config(state='normal')
                    self.log_area.delete('1.0', tk.END)
                    for line in lines[-100:]:
                        self.insert_to_log(line)
                    self.log_area.config(state='disabled')
            except Exception as e:
                messagebox.showerror("Error", str(e))
        if self.auto_refresh:
            self.after(1000, self.update_log_view)

    def insert_to_log(self, text):
        self.log_area.config(state='normal')
        applied_tag = None
        for keyword, color in self.highlight_words.items():
            if keyword in text.upper():
                self.log_area.tag_config(keyword, background=color)
                applied_tag = keyword
                break
        text = text.rstrip('\n')
        self.log_area.insert(tk.END, text + "\n", applied_tag if applied_tag else None)
        self.log_area.config(state='disabled')
        self.log_area.yview(tk.END)

    def search_in_logs(self):
        search_term = simpledialog.askstring("Search", "Enter search text:")
        if search_term:
            self.log_area.tag_remove('search', '1.0', tk.END)  # Remove any previous highlights
            start = '1.0'
            self.log_area.tag_config('search', background='yellow', foreground='black')  # Ensure tag is configured

            while True:
                start = self.log_area.search(search_term, start, nocase=1, stopindex=tk.END)
                if not start:
                    break
                end = f"{start}+{len(search_term)}c"
                self.log_area.tag_add('search', start, end)
                start = end
            self.log_area.see(tk.END)

    def prompt_restart_service(self):
        service_name = simpledialog.askstring("Restart Service", "Enter the service name:")
        if service_name:
            self.restart_service(service_name)

    def restart_service(self, service_name):
        try:
            subprocess.run(["net", "stop", service_name], check=True, shell=True)
            subprocess.run(["net", "start", service_name], check=True, shell=True)
            messagebox.showinfo("Success", f"Service {service_name} restarted successfully!")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"Failed to restart {service_name}: {str(e)}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def select_log_file(self):
        path = filedialog.askopenfilename(title="Select a log file",
                                          filetypes=[("All files", "*.*"), ("Text files", "*.txt"), ("Log files", "*.log")],
                                          defaultextension="*.*")
        if path:
            self.log_file_path = path
            self.filepath_label.config(text=f"File: {self.log_file_path}")
            self.restart_log_view()

    def restart_log_view(self):
        """Refresh the log view with new file content."""
        self.update_log_view(force=True)

    def change_font(self, choice):
        new_font = Font(family=choice, size=10)
        self.log_area.config(font=new_font)

if __name__ == "__main__":
    app = LogViewer()
    app.mainloop()
