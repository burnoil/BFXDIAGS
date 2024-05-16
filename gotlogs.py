import tkinter as tk
from tkinter import ttk, scrolledtext, simpledialog, messagebox, filedialog, Label, StringVar, OptionMenu, Menu, Entry, Button, colorchooser
from tkinter.font import Font
import subprocess
import sys
import os
import socket

class LogViewer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Log Viewer")
        self.geometry("800x600")

        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(__file__)

        icon_path = os.path.join(base_path, 'log.ico')
        self.iconbitmap(icon_path)

        self.button_font = Font(family="Helvetica", size=8, weight="bold")
        self.label_font = Font(family="Helvetica", size=8)
        self.current_font = Font(family="Helvetica", size=10)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.controls_frame = tk.Frame(self)
        self.controls_frame.pack(fill=tk.X)

        self.status_frame = tk.Frame(self)
        self.status_frame.pack(fill=tk.X)

        self.status_label = Label(self.status_frame, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W, font=self.label_font)
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.filepath_label = Label(self.status_frame, text="", bd=1, relief=tk.SUNKEN, anchor=tk.W, font=self.label_font)
        self.filepath_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.auto_refresh = True
        self.highlight_words = {"ERROR": "red", "WARNING": "yellow", "SUCCESSFUL": "light green"}

        self.create_controls()
        self.create_menu()

    def create_controls(self):
        self.pause_button = tk.Button(self.controls_frame, text="Pause Refresh", command=self.toggle_refresh, font=self.button_font, bg="lightblue", fg="black")
        self.pause_button.pack(side=tk.LEFT, padx=3, pady=3)

        self.refresh_button = tk.Button(self.controls_frame, text="Refresh Now", command=self.manual_refresh, font=self.button_font, bg="lightgreen", fg="black")
        self.refresh_button.pack(side=tk.LEFT, padx=3, pady=3)

        self.search_button = tk.Button(self.controls_frame, text="Search", command=self.search_in_logs, font=self.button_font, bg="lightgrey", fg="black")
        self.search_button.pack(side=tk.LEFT, padx=3, pady=3)

        self.file_button = tk.Button(self.controls_frame, text="Select Log File", command=self.select_log_file, font=self.button_font, bg="lightyellow", fg="black")
        self.file_button.pack(side=tk.LEFT, padx=3, pady=3)

        self.restart_button = tk.Button(self.controls_frame, text="Restart Service", command=self.prompt_restart_service, font=self.button_font, bg="orange", fg="black")
        self.restart_button.pack(side=tk.LEFT, padx=3, pady=3)

        self.host_entry = Entry(self.controls_frame, width=12, font=self.button_font)
        self.host_entry.pack(side=tk.LEFT, padx=3)
        self.host_entry.insert(0, "127.0.0.1")  # Default value

        self.port_entry = Entry(self.controls_frame, width=5, font=self.button_font)
        self.port_entry.pack(side=tk.LEFT, padx=3)
        self.port_entry.insert(0, "80")  # Default value

        self.port_check_button = Button(self.controls_frame, text="Check Port", command=self.check_port, font=self.button_font, bg="lightcoral", fg="black")
        self.port_check_button.pack(side=tk.LEFT, padx=3, pady=3)

        self.clear_button = tk.Button(self.controls_frame, text="Clear Log", command=self.clear_log_display, font=self.button_font, bg="lightpink", fg="black")
        self.clear_button.pack(side=tk.LEFT, padx=3, pady=3)

        self.save_button = tk.Button(self.controls_frame, text="Save Filtered Logs", command=self.save_filtered_logs, font=self.button_font, bg="lightgray", fg="black")
        self.save_button.pack(side=tk.LEFT, padx=3, pady=3)

        self.font_var = StringVar(self)
        self.font_options = ["Helvetica", "Courier", "Times"]
        self.font_var.set(self.font_options[0])
        self.font_dropdown = OptionMenu(self.controls_frame, self.font_var, *self.font_options, command=self.change_font)
        self.font_dropdown.config(font=self.button_font)
        self.font_dropdown.pack(side=tk.RIGHT, padx=3, pady=3)

    def create_menu(self):
        menu_bar = Menu(self)
        self.config(menu=menu_bar)

        help_menu = Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="Usage Instructions", command=self.show_usage)

        settings_menu = Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="Change Colors", command=self.change_colors)

    def show_about(self):
        messagebox.showinfo("About", "LogViewer\nVersion 1.0\nDeveloped by: Your Name")

    def show_usage(self):
        usage_text = (
            "Usage Instructions:\n"
            "- To pause the log refresh, click 'Pause Refresh'.\n"
            "- To manually refresh the logs, click 'Refresh Now'.\n"
            "- To search within the logs, click 'Search' and enter the search term.\n"
            "- To restart a service, click 'Restart Service' and enter the service name.\n"
            "- To check if a port is open, enter the host and port, then click 'Check Port'.\n"
            "- To change the display font, select from the dropdown menu.\n"
            "- To select a new log file, click 'Select Log File'.\n"
            "- To change highlight colors, go to 'Settings' > 'Change Colors'.\n"
            "- To clear the log display, click 'Clear Log'.\n"
            "- To save filtered logs, click 'Save Filtered Logs'."
        )
        messagebox.showinfo("Usage Instructions", usage_text)

    def toggle_refresh(self):
        self.auto_refresh = not self.auto_refresh
        self.pause_button.config(text="Resume Refresh" if not self.auto_refresh else "Pause Refresh")
        self.update_current_log_view(force=True)

    def manual_refresh(self):
        self.update_current_log_view(force=True)

    def update_current_log_view(self, force=False):
        current_tab = self.notebook.select()
        log_area = self.notebook.nametowidget(current_tab).log_area
        log_file_path = self.notebook.nametowidget(current_tab).log_file_path

        if (self.auto_refresh or force) and log_file_path:
            try:
                with open(log_file_path, 'r') as file:
                    lines = file.readlines()
                    log_area.config(state='normal')
                    log_area.delete('1.0', tk.END)
                    for line in lines[-100:]:
                        self.insert_to_log(log_area, line)
                    log_area.config(state='disabled')
            except Exception as e:
                messagebox.showerror("Error", str(e))
        if self.auto_refresh:
            self.after(1000, self.update_current_log_view)

    def insert_to_log(self, log_area, text):
        log_area.config(state='normal')
        applied_tag = None
        for keyword, color in self.highlight_words.items():
            if keyword in text.upper():
                log_area.tag_config(keyword, background=color)
                applied_tag = keyword
                break
        text = text.rstrip('\n')
        log_area.insert(tk.END, text + "\n", applied_tag if applied_tag else None)
        log_area.config(state='disabled')
        log_area.yview(tk.END)

    def search_in_logs(self):
        search_term = simpledialog.askstring("Search", "Enter search text:")
        if search_term:
            current_tab = self.notebook.select()
            log_area = self.notebook.nametowidget(current_tab).log_area

            log_area.tag_remove('search', '1.0', tk.END)
            start = '1.0'
            log_area.tag_config('search', background='yellow', foreground='black')

            while True:
                start = log_area.search(search_term, start, nocase=1, stopindex=tk.END)
                if not start:
                    break
                end = f"{start}+{len(search_term)}c"
                log_area.tag_add('search', start, end)
                start = end
            log_area.see(tk.END)

    def prompt_restart_service(self):
        service_name = simpledialog.askstring("Restart Service", "Enter the service name:")
        if service_name:
            self.restart_service(service_name)

    def restart_service(self, service_name):
        try:
            subprocess.run(["net", "stop", service_name], check=True, shell=True)
            subprocess.run(["net", "start", service_name], check=True, shell=True)
            self.status_label.config(text=f"Service {service_name} restarted successfully!")
        except subprocess.CalledProcessError as e:
            self.status_label.config(text=f"Failed to restart {service_name}: {str(e)}")
        except Exception as e:
            self.status_label.config(text=f"An error occurred: {str(e)}")

    def check_port(self):
        host = self.host_entry.get()
        port_str = self.port_entry.get()
        try:
            port = int(port_str)
            if self.is_port_open(host, port):
                self.status_label.config(text=f"Port {port} on {host} is open.")
            else:
                self.status_label.config(text=f"Port {port} on {host} is closed.")
        except ValueError:
            self.status_label.config(text="Invalid port number.")

    @staticmethod
    def is_port_open(host, port):
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(1)  # Timeout after 1 second
        try:
            sock.connect((host, port))
            sock.close()
            return True
        except (socket.timeout, socket.error):
            return False

    def select_log_file(self):
        path = filedialog.askopenfilename(title="Select a log file",
                                          filetypes=[("All files", "*.*"), ("Text files", "*.txt"), ("Log files", "*.log")],
                                          defaultextension="*.*")
        if path:
            self.create_log_tab(path)

    def create_log_tab(self, log_file_path):
        tab_frame = tk.Frame(self.notebook)
        tab_frame.log_file_path = log_file_path

        log_area = scrolledtext.ScrolledText(tab_frame, state='disabled', wrap=tk.WORD, font=self.current_font, bg="lightgray")
        log_area.pack(fill=tk.BOTH, expand=True)
        tab_frame.log_area = log_area

        tab_title = os.path.basename(log_file_path)
        self.notebook.add(tab_frame, text=tab_title)
        self.notebook.select(tab_frame)

        self.filepath_label.config(text=f"File: {log_file_path}")

        self.update_log_view(log_area, log_file_path)

    def update_log_view(self, log_area, log_file_path):
        if self.auto_refresh and log_file_path:
            try:
                with open(log_file_path, 'r') as file:
                    lines = file.readlines()
                    log_area.config(state='normal')
                    log_area.delete('1.0', tk.END)
                    for line in lines[-100:]:
                        self.insert_to_log(log_area, line)
                    log_area.config(state='disabled')
            except Exception as e:
                messagebox.showerror("Error", str(e))
        if self.auto_refresh:
            self.after(1000, lambda: self.update_log_view(log_area, log_file_path))

    def change_font(self, choice):
        new_font = Font(family=choice, size=10)
        for tab in self.notebook.tabs():
            log_area = self.notebook.nametowidget(tab).log_area
            log_area.config(font=new_font)

    def change_colors(self):
        for keyword in self.highlight_words:
            color = colorchooser.askcolor(title=f"Choose color for {keyword}")[1]
            if color:
                self.highlight_words[keyword] = color

    def clear_log_display(self):
        current_tab = self.notebook.select()
        log_area = self.notebook.nametowidget(current_tab).log_area

        log_area.config(state='normal')
        log_area.delete('1.0', tk.END)
        log_area.config(state='disabled')

    def save_filtered_logs(self):
        current_tab = self.notebook.select()
        log_area = self.notebook.nametowidget(current_tab).log_area

        filtered_text = log_area.get("1.0", tk.END).strip()
        if not filtered_text:
            messagebox.showwarning("Warning", "No logs to save. Please apply a filter first.")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                 filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if save_path:
            try:
                with open(save_path, 'w') as file:
                    file.write(filtered_text)
                messagebox.showinfo("Success", f"Filtered logs saved to {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save logs: {str(e)}")

if __name__ == "__main__":
    app = LogViewer()
    app.mainloop()
