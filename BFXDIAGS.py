import os
import tkinter as tk
from tkinter import ttk, scrolledtext, simpledialog, messagebox, filedialog, Label, StringVar, OptionMenu, Menu, Entry, Button, colorchooser
from tkinter.font import Font
import subprocess
import sys
import socket
import glob
import winreg
import shutil
from datetime import datetime

class LogViewer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BFXDIAGS 1.29")
        self.geometry("1200x600")

        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(__file__)

        icon_path = os.path.join(base_path, 'log.ico')
        self.iconbitmap(icon_path)

        self.button_font = Font(family="Tahoma", size=8, weight="bold")
        self.label_font = Font(family="Tahoma", size=8)
        self.current_font = Font(family="Tahoma", size=10)

        self.create_status_labels()
        self.create_controls()
        self.create_main_layout()
        self.create_menu()

        self.auto_refresh = True
        self.highlight_words = {"ERROR": "red", "WARNING": "yellow", "SUCCESSFUL": "light green"}
        self.tab_frames = {}

        self.update_besclient_status()
        self.update_relay_status()
        self.update_workstation_info()
        self.update_besclient_version()
        self.update_throttling_status()
        self.update_subsidiary_status()
        self.update_data_folder_status()

    def create_status_labels(self):
        self.info_frame = tk.Frame(self)
        self.info_frame.pack(fill=tk.X, pady=3)

        self.besclient_status_label = Label(self.info_frame, text="BESClient Status: Unknown", font=self.button_font, bg="lightgray", fg="black")
        self.besclient_status_label.pack(side=tk.LEFT, padx=3, pady=3)

        self.port_status_label = Label(self.info_frame, text="Port 52311 Status: Unknown", font=self.button_font, bg="lightgray", fg="black")
        self.port_status_label.pack(side=tk.LEFT, padx=3, pady=3)

        self.workstation_name_label = Label(self.info_frame, text="Workstation Name: Unknown", font=self.button_font, bg="lightgray", fg="black")
        self.workstation_name_label.pack(side=tk.LEFT, padx=3, pady=3)

        self.ip_address_label = Label(self.info_frame, text="IP Address: Unknown", font=self.button_font, bg="lightgray", fg="black")
        self.ip_address_label.pack(side=tk.LEFT, padx=3, pady=3)

        self.besclient_version_label = Label(self.info_frame, text="BESClient Version: Unknown", font=self.button_font, bg="lightgray", fg="black")
        self.besclient_version_label.pack(side=tk.LEFT, padx=3, pady=3)

        self.subsidiary_status_label = Label(self.info_frame, text="Subsidiary: Unknown", font=self.button_font, bg="lightgray", fg="black")
        self.subsidiary_status_label.pack(side=tk.LEFT, padx=3, pady=3)

        self.status_frame = tk.Frame(self)
        self.status_frame.pack(fill=tk.X, pady=3)

        self.data_folder_status_label = Label(self.status_frame, text="Data Folder Status: Monitoring...", font=self.button_font, bg="lightgray", fg="black")
        self.data_folder_status_label.pack(side=tk.LEFT, padx=3, pady=3)

        self.relay_status_label = Label(self.status_frame, text="Relay Status: Unknown", font=self.button_font, bg="lightgray", fg="black")
        self.relay_status_label.pack(side=tk.LEFT, padx=3, pady=3)

        self.status_label = Label(self.status_frame, text="Status: Ready", font=self.button_font, bg="lightgray", fg="black")
        self.status_label.pack(side=tk.LEFT, padx=3, pady=3)

    def create_controls(self):
        self.controls_frame = tk.Frame(self)
        self.controls_frame.pack(fill=tk.Y, side=tk.LEFT)

        self.pause_button = tk.Button(self.controls_frame, text="Pause Refresh", command=self.toggle_refresh, font=self.button_font, bg="lightblue", fg="black")
        self.pause_button.pack(fill=tk.X, padx=3, pady=3)

        self.refresh_button = tk.Button(self.controls_frame, text="Refresh Now", command=self.manual_refresh, font=self.button_font, bg="lightgreen", fg="black")
        self.refresh_button.pack(fill=tk.X, padx=3, pady=3)

        self.search_button = tk.Button(self.controls_frame, text="Search", command=self.search_in_logs, font=self.button_font, bg="lightgrey", fg="black")
        self.search_button.pack(fill=tk.X, padx=3, pady=3)

        self.file_button = tk.Button(self.controls_frame, text="Select Log File", command=self.select_log_file, font=self.button_font, bg="lightyellow", fg="black")
        self.file_button.pack(fill=tk.X, padx=3, pady=3)

        self.latest_log_button = tk.Button(self.controls_frame, text="Open Latest BigFix Log", command=self.open_latest_bigfix_log, font=self.button_font, bg="lightblue", fg="black")
        self.latest_log_button.pack(fill=tk.X, padx=3, pady=3)

        self.monitor_data_button = tk.Button(self.controls_frame, text="Monitor BES Data Folder", command=self.monitor_bes_data_folder, font=self.button_font, bg="lightgreen", fg="black")
        self.monitor_data_button.pack(fill=tk.X, padx=3, pady=3)

        self.restart_button = tk.Button(self.controls_frame, text="Restart BESClient Service", command=self.restart_besclient_service, font=self.button_font, bg="orange", fg="black")
        self.restart_button.pack(fill=tk.X, padx=3, pady=3)

        self.clear_cache_button = tk.Button(self.controls_frame, text="Clear Latest Site Cache", command=self.clear_latest_site_cache, font=self.button_font, bg="red", fg="white")
        self.clear_cache_button.pack(fill=tk.X, padx=3, pady=3)

        self.toggle_throttling_button = tk.Button(self.controls_frame, text="Toggle Throttling", command=self.toggle_throttling, font=self.button_font, bg="lightcoral", fg="black")
        self.toggle_throttling_button.pack(fill=tk.X, padx=3, pady=3)

        self.edit_registry_button = tk.Button(self.controls_frame, text="Edit BigFix Registry", command=self.open_registry_editor, font=self.button_font, bg="lightcoral", fg="black")
        self.edit_registry_button.pack(fill=tk.X, padx=3, pady=3)

        self.host_entry = Entry(self.controls_frame, width=12, font=self.button_font)
        self.host_entry.pack(fill=tk.X, padx=3)
        self.host_entry.insert(0, "127.0.0.1")  # Default value

        self.port_entry = Entry(self.controls_frame, width=5, font=self.button_font)
        self.port_entry.pack(fill=tk.X, padx=3)
        self.port_entry.insert(0, "80")  # Default value

        self.port_check_button = Button(self.controls_frame, text="Check Port", command=self.check_port, font=self.button_font, bg="lightcoral", fg="black")
        self.port_check_button.pack(fill=tk.X, padx=3, pady=3)

        self.filter_frame = tk.Frame(self.controls_frame)
        self.filter_frame.pack(fill=tk.X, pady=3)

        self.filter_entry = Entry(self.filter_frame, width=20, font=self.button_font)
        self.filter_entry.pack(fill=tk.X, padx=3, pady=3)

        self.filter_button = tk.Button(self.filter_frame, text="Filter Logs", command=self.filter_logs, font=self.button_font, bg="lightblue", fg="black")
        self.filter_button.pack(fill=tk.X, padx=3, pady=3)

        self.clear_button = tk.Button(self.filter_frame, text="Clear Log", command=self.clear_log_display, font=self.button_font, bg="lightpink", fg="black")
        self.clear_button.pack(fill=tk.X, padx=3, pady=3)

        self.save_button = tk.Button(self.filter_frame, text="Save Filtered Logs", command=self.save_filtered_logs, font=self.button_font, bg="lightgray", fg="black")
        self.save_button.pack(fill=tk.X, padx=3, pady=3)

        self.font_var = StringVar(self)
        self.font_options = ["Tahoma", "Segoe UI", "Arial", "Courier", "Times", "Helvetica"]
        self.font_var.set(self.font_options[0])
        self.font_dropdown = OptionMenu(self.filter_frame, self.font_var, *self.font_options, command=self.change_font)
        self.font_dropdown.config(font=self.button_font)
        self.font_dropdown.pack(fill=tk.X, padx=3, pady=3)

    def create_main_layout(self):
        self.main_frame = tk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=True, side=tk.RIGHT)

        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.filepath_label = Label(self.main_frame, text="", bd=1, relief=tk.SUNKEN, anchor=tk.W, font=self.label_font)
        self.filepath_label.pack(fill=tk.X, side=tk.BOTTOM, padx=3, pady=3)

    def create_log_tab(self, log_file_path):
        tab_frame = tk.Frame(self.notebook)
        tab_frame.log_file_path = log_file_path

        log_area = scrolledtext.ScrolledText(tab_frame, state='disabled', wrap=tk.WORD, font=self.current_font, bg="lightgray")
        log_area.pack(fill=tk.BOTH, expand=True)
        tab_frame.log_area = log_area

        tab_title = os.path.basename(log_file_path)
        self.notebook.add(tab_frame, text=tab_title)
        self.tab_frames[tab_frame] = tab_title

        self.add_close_button(tab_frame, tab_title)

        self.notebook.select(tab_frame)

        self.filepath_label.config(text=f"File: {log_file_path}")

        try:
            self.update_log_view(log_area, log_file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file {log_file_path}: {str(e)}")
            self.notebook.forget(tab_frame)
            del self.tab_frames[tab_frame]

    def add_close_button(self, tab_frame, tab_title):
        close_button = tk.Button(tab_frame, text="Close Tab", command=lambda: self.close_tab(tab_frame), bg="red", fg="white", font=("Arial", 8, "bold"))
        close_button.pack(side=tk.RIGHT, padx=5, pady=5)

    def close_tab(self, tab_frame):
        self.notebook.forget(tab_frame)
        del self.tab_frames[tab_frame]

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
                raise e
        if self.auto_refresh:
            self.after(1000, lambda: self.update_log_view(log_area, log_file_path))

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

    def restart_besclient_service(self):
        try:
            self.update_status_label("Stopping BESClient service...", "orange")
            subprocess.run(["net", "stop", "BESClient"], check=True, shell=True)
            self.update_status_label("Starting BESClient service...", "orange")
            subprocess.run(["net", "start", "BESClient"], check=True, shell=True)
            self.update_status_label("BESClient service restarted successfully!", "lightgreen")
        except subprocess.CalledProcessError as e:
            self.update_status_label(f"Failed to restart BESClient service: {str(e)}", "red")
        except Exception as e:
            self.update_status_label(f"An error occurred: {str(e)}", "red")
        self.update_besclient_status()

    def clear_latest_site_cache(self):
        try:
            latest_custom_site = self.get_latest_custom_site(r"C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData")
            if latest_custom_site == "No Custom Sites Found":
                messagebox.showwarning("Warning", "No Custom Sites found to clear cache.")
                return

            cache_directory = os.path.join(r"C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData", latest_custom_site)
            self.update_status_label(f"Clearing cache for {latest_custom_site}...", "orange")

            if os.path.exists(cache_directory):
                shutil.rmtree(cache_directory)
                os.makedirs(cache_directory)

            self.update_status_label(f"Cache for {latest_custom_site} cleared successfully!", "lightgreen")
        except Exception as e:
            self.update_status_label(f"Failed to clear cache: {str(e)}", "red")
        self.update_data_folder_status()

    def update_status_label(self, message, color):
        self.status_label.config(text=message, bg=color)
        self.update_idletasks()

    def check_port(self):
        host = self.host_entry.get()
        port_str = self.port_entry.get()
        try:
            port = int(port_str)
            if self.is_port_open(host, port):
                self.update_status_label(f"Port {port} on {host} is open.", "lightgreen")
            else:
                self.update_status_label(f"Port {port} on {host} is closed.", "red")
        except ValueError:
            self.update_status_label("Invalid port number.", "red")

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

    def open_latest_bigfix_log(self):
        bigfix_log_directory = r"C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs"
        log_files = glob.glob(os.path.join(bigfix_log_directory, "*.log"))
        if not log_files:
            messagebox.showwarning("Warning", "No BigFix log files found.")
            return

        latest_log_file = max(log_files, key=os.path.getmtime)
        self.create_log_tab(latest_log_file)

    def monitor_bes_data_folder(self):
        bes_data_directory = r"C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData"
        if not os.path.exists(bes_data_directory):
            messagebox.showwarning("Warning", "BES Data directory not found.")
            return

        self.create_data_folder_tab(bes_data_directory)

    def create_data_folder_tab(self, data_directory):
        tab_frame = tk.Frame(self.notebook)
        tab_frame.data_directory = data_directory

        log_area = scrolledtext.ScrolledText(tab_frame, state='normal', wrap=tk.WORD, font=self.current_font, bg="lightgray")
        log_area.pack(fill=tk.BOTH, expand=True)
        tab_frame.log_area = log_area

        tab_title = "BES Data Folder"
        self.notebook.add(tab_frame, text=tab_title)
        self.tab_frames[tab_frame] = tab_title

        self.add_close_button(tab_frame, tab_title)

        self.notebook.select(tab_frame)

        self.filepath_label.config(text=f"Directory: {data_directory}")

        self.update_data_folder_view(log_area, data_directory)

    def update_data_folder_view(self, log_area, data_directory):
        log_area.config(state='normal')
        log_area.delete('1.0', tk.END)

        try:
            folders = [(folder, os.path.getmtime(os.path.join(data_directory, folder)))
                       for folder in os.listdir(data_directory) if os.path.isdir(os.path.join(data_directory, folder))]
            folders.sort(key=lambda x: x[1], reverse=True)  # Sort by modification time, newest first

            # Display only the last 10 most recently updated folders
            for folder, mtime in folders[:10]:
                mod_time = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
                log_area.insert(tk.END, f"{folder} - Last Modified: {mod_time}\n")

        except Exception as e:
            messagebox.showerror("Error", str(e))

        log_area.config(state='disabled')
        if self.auto_refresh:
            self.after(10000, lambda: self.update_data_folder_view(log_area, data_directory))  # Refresh every 10 seconds

    def update_besclient_status(self):
        status = self.get_service_status("besclient")
        port_status = self.is_port_open("127.0.0.1", 52311)
        self.besclient_status_label.config(text=f"BESClient Status: {status}")
        self.port_status_label.config(text=f"Port 52311 Status: {'Open' if port_status else 'Closed'}")

        if status == "Running" and port_status:
            self.besclient_status_label.config(bg="lightgreen", fg="black")
            self.port_status_label.config(bg="lightgreen", fg="black")
        elif status == "Stopped" or not port_status:
            self.besclient_status_label.config(bg="red", fg="white")
            self.port_status_label.config(bg="red", fg="white")
        else:
            self.besclient_status_label.config(bg="yellow", fg="black")
            self.port_status_label.config(bg="yellow", fg="black")

        self.after(5000, self.update_besclient_status)  # Update status every 5 seconds

    def get_service_status(self, service_name):
        try:
            result = subprocess.run(["sc", "query", service_name], capture_output=True, text=True, shell=True)
            if "RUNNING" in result.stdout:
                return "Running"
            elif "STOPPED" in result.stdout:
                return "Stopped"
            else:
                return "Unknown"
        except Exception as e:
            return f"Error: {str(e)}"

    def update_workstation_info(self):
        hostname = socket.gethostname()
        self.workstation_name_label.config(text=f"Workstation Name: {hostname}")
        ip_addresses = self.get_ip_addresses()
        self.ip_address_label.config(text=f"IP Address(es): {', '.join(ip_addresses)}")

    def get_ip_addresses(self):
        ip_addresses = []
        hostname = socket.gethostname()
        for ip in socket.getaddrinfo(hostname, None):
            if ':' not in ip[4][0]:  # Filter out IPv6 addresses
                ip_addresses.append(ip[4][0])
        return ip_addresses

    def update_besclient_version(self):
        try:
            besclient_exe_path = r"C:\Program Files (x86)\BigFix Enterprise\BES Client\besclient.exe"
            if os.path.exists(besclient_exe_path):
                result = subprocess.run(['wmic', 'datafile', 'where', f'name="{besclient_exe_path.replace("\\", "\\\\")}"', 'get', 'Version', '/value'], capture_output=True, text=True)
                version_info = result.stdout.strip().split('=')[-1]
                self.besclient_version_label.config(text=f"BESClient Version: {version_info}")
            else:
                self.besclient_version_label.config(text="BESClient Version: Not Found")
        except Exception as e:
            self.besclient_version_label.config(text=f"BESClient Version: Error ({str(e)})")

    def update_relay_status(self):
        relay_info = self.get_relay_info()
        self.relay_status_label.config(text=f"Relay Status: {relay_info}")
        self.after(5000, self.update_relay_status)  # Update relay status every 5 seconds

    def get_relay_info(self):
        try:
            registry_key1 = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\BigFix\EnterpriseClient\Settings\Client\__RelayServer1")
            relay_server1, _ = winreg.QueryValueEx(registry_key1, "value")
            winreg.CloseKey(registry_key1)

            registry_key2 = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\BigFix\EnterpriseClient\Settings\Client\__RelayServer2")
            relay_server2, _ = winreg.QueryValueEx(registry_key2, "value")
            winreg.CloseKey(registry_key2)

            return f"{relay_server1}, {relay_server2}"
        except FileNotFoundError:
            return "Registry key not found"
        except Exception as e:
            return f"Error: {str(e)}"

    def update_throttling_status(self):
        throttling_status = self.get_throttling_status()
        self.toggle_throttling_button.config(text=f"Throttling Exempt: {throttling_status}")

    def get_throttling_status(self):
        try:
            registry_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\BigFix\EnterpriseClient\Settings\Client\ThrottlingExempt")
            throttling_status, _ = winreg.QueryValueEx(registry_key, "value")
            winreg.CloseKey(registry_key)
            return throttling_status
        except FileNotFoundError:
            return "Not Found"
        except Exception as e:
            return f"Error: {str(e)}"

    def toggle_throttling(self):
        current_status = self.get_throttling_status()
        new_status = "NO" if current_status == "YES" else "YES"
        try:
            registry_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\BigFix\EnterpriseClient\Settings\Client\ThrottlingExempt", 0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(registry_key, "value", 0, winreg.REG_SZ, new_status)
            winreg.CloseKey(registry_key)
            self.toggle_throttling_button.config(text=f"Throttling: {new_status}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to toggle throttling: {str(e)}")

    def update_subsidiary_status(self):
        try:
            registry_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\BigFix\EnterpriseClient\Settings\Client\Subsidiary")
            subsidiary_status, _ = winreg.QueryValueEx(registry_key, "value")
            winreg.CloseKey(registry_key)
            self.subsidiary_status_label.config(text=f"Subsidiary: {subsidiary_status}")
        except FileNotFoundError:
            self.subsidiary_status_label.config(text="Subsidiary: Not Found")
        except Exception as e:
            self.subsidiary_status_label.config(text=f"Subsidiary: Error ({str(e)})")

    def human_readable_size(self, size, decimal_places=2):
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size < 1024.0:
                return f"{size:.{decimal_places}f} {unit}"
            size /= 1024.0

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
        messagebox.showinfo("About", "BFXDIAGS\nVersion 1.29\nDeveloped by: https://github.com/burnoil/PythonProjects")

    def show_usage(self):
        usage_text = (
            "Usage Instructions:\n"
            "- To pause the log refresh, click 'Pause Refresh'.\n"
            "- To manually refresh the logs, click 'Refresh Now'.\n"
            "- To search within the logs, click 'Search' and enter the search term.\n"
            "- To restart the BESClient service, click 'Restart BESClient Service'.\n"
            "- To clear the latest site cache, click 'Clear Latest Site Cache'.\n"
            "- To check if a port is open, enter the host and port, then click 'Check Port'.\n"
            "- To change the display font, select from the dropdown menu.\n"
            "- To select a new log file, click 'Select Log File'.\n"
            "- To change highlight colors, go to 'Settings' > 'Change Colors'.\n"
            "- To clear the log display, click 'Clear Log'.\n"
            "- To save filtered logs, click 'Save Filtered Logs'.\n"
            "- To filter logs, enter a keyword and click 'Filter Logs'.\n"
            "- To monitor the BES Data folder, click 'Monitor BES Data Folder'.\n"
            "- To edit BigFix registry settings, click 'Edit BigFix Registry'."
        )
        messagebox.showinfo("Usage Instructions", usage_text)

    def open_registry_editor(self):
        self.registry_editor_window = tk.Toplevel(self)
        self.registry_editor_window.title("BigFix Registry Editor")
        self.registry_editor_window.geometry("1000x600")

        self.editor_frame = tk.Frame(self.registry_editor_window)
        self.editor_frame.pack(fill=tk.BOTH, expand=True)

        self.registry_tree = ttk.Treeview(self.editor_frame)
        self.registry_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.registry_tree.heading("#0", text="Registry Key", anchor='w')

        self.populate_registry_tree()

        self.registry_tree.bind("<Double-1>", self.on_registry_item_double_click)

        self.values_frame = tk.Frame(self.editor_frame)
        self.values_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.values_text = scrolledtext.ScrolledText(self.values_frame, wrap=tk.WORD, font=self.current_font)
        self.values_text.pack(fill=tk.BOTH, expand=True)

        self.values_text.bind("<Double-1>", self.edit_registry_value)

    def populate_registry_tree(self):
        root_key = r"SOFTWARE\WOW6432Node\BigFix"
        parent_node = ""
        self.insert_registry_items(parent_node, root_key)

    def insert_registry_items(self, parent_node, key):
        try:
            reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key)
            index = 0
            while True:
                try:
                    sub_key = winreg.EnumKey(reg_key, index)
                    sub_key_path = f"{key}\\{sub_key}"
                    node_id = self.registry_tree.insert(parent_node, 'end', text=sub_key, open=False)
                    self.insert_registry_items(node_id, sub_key_path)
                    index += 1
                except OSError:
                    break
            winreg.CloseKey(reg_key)
        except Exception as e:
            print(f"Failed to open registry key {key}: {e}")

    def on_registry_item_double_click(self, event):
        selected_item = self.registry_tree.selection()[0]
        item_path = self.get_item_path(selected_item)
        print(f"Opening registry key: {item_path}")  # Debugging line
        self.show_registry_key_values(item_path)

    def get_item_path(self, item_id):
        path_parts = []
        while item_id:
            item_text = self.registry_tree.item(item_id, "text")
            path_parts.insert(0, item_text)
            item_id = self.registry_tree.parent(item_id)
        return "SOFTWARE\\WOW6432Node\\BigFix\\" + "\\".join(path_parts)

    def show_registry_key_values(self, key):
        try:
            self.values_text.delete('1.0', tk.END)
            reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key)
            index = 0
            values = []
            while True:
                try:
                    value_name, value_data, value_type = winreg.EnumValue(reg_key, index)
                    value_type_str = self.get_registry_type_string(value_type)
                    values.append((value_name, value_data, value_type_str))
                    index += 1
                except OSError:
                    break
            winreg.CloseKey(reg_key)

            values_str = "\n".join([f"{name} ({type_str}): {data}" for name, data, type_str in values])
            self.values_text.insert(tk.END, values_str)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open registry key: {e}")

    @staticmethod
    def get_registry_type_string(value_type):
        type_map = {
            winreg.REG_SZ: "STRING",
            winreg.REG_DWORD: "DWORD",
            winreg.REG_BINARY: "BINARY",
            winreg.REG_MULTI_SZ: "MULTI_STRING",
            winreg.REG_EXPAND_SZ: "EXPAND_STRING",
            winreg.REG_QWORD: "QWORD",
            winreg.REG_NONE: "NONE",
        }
        return type_map.get(value_type, "UNKNOWN")

    def edit_registry_value(self, event):
        try:
            index = self.values_text.index("@%s,%s" % (event.x, event.y))
            line = self.values_text.get(f"{index} linestart", f"{index} lineend")
            name_type, data = line.split(": ", 1)
            name, value_type = name_type.split(" (", 1)
            value_type = value_type.strip(")")

            selected_item = self.registry_tree.selection()[0]
            item_path = self.get_item_path(selected_item)

            new_value = simpledialog.askstring("Edit Value", f"Enter new value for {name} ({value_type}):", initialvalue=data)
            if new_value is not None:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, item_path, 0, winreg.KEY_SET_VALUE) as reg_key:
                    if value_type == "STRING" or value_type == "EXPAND_STRING":
                        winreg.SetValueEx(reg_key, name, 0, winreg.REG_SZ, new_value)
                    elif value_type == "DWORD":
                        winreg.SetValueEx(reg_key, name, 0, winreg.REG_DWORD, int(new_value))
                    elif value_type == "QWORD":
                        winreg.SetValueEx(reg_key, name, 0, winreg.REG_QWORD, int(new_value))
                    elif value_type == "MULTI_STRING":
                        winreg.SetValueEx(reg_key, name, 0, winreg.REG_MULTI_SZ, new_value.splitlines())
                    elif value_type == "BINARY":
                        winreg.SetValueEx(reg_key, name, 0, winreg.REG_BINARY, bytes.fromhex(new_value))
                self.show_registry_key_values(item_path)
                messagebox.showinfo("Success", f"Value for {name} updated successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update value: {e}")

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

    def filter_logs(self):
        filter_term = self.filter_entry.get()
        if filter_term:
            current_tab = self.notebook.select()
            log_area = self.notebook.nametowidget(current_tab).log_area

            log_area.tag_remove('filter', '1.0', tk.END)
            start = '1.0'
            log_area.tag_config('filter', background='lightblue', foreground='black')

            while True:
                start = log_area.search(filter_term, start, nocase=1, stopindex=tk.END)
                if not start:
                    break
                end = f"{start}+{len(filter_term)}c"
                log_area.tag_add('filter', start, end)
                start = end
            log_area.see(tk.END)

    def clear_log_display(self):
        current_tab = self.notebook.select()
        log_area = self.notebook.nametowidget(current_tab).log_area
        log_area.config(state='normal')
        log_area.delete('1.0', tk.END)
        log_area.config(state='disabled')

    def save_filtered_logs(self):
        current_tab = self.notebook.select()
        log_area = self.notebook.nametowidget(current_tab).log_area
        log_content = log_area.get('1.0', tk.END).strip()

        save_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if save_path:
            try:
                with open(save_path, 'w') as file:
                    file.write(log_content)
                messagebox.showinfo("Save Logs", f"Filtered logs saved successfully to {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save logs: {e}")

    def change_font(self, font_name):
        self.current_font.config(family=font_name)
        for tab_frame in self.tab_frames.keys():
            tab_frame.log_area.config(font=self.current_font)
        self.values_text.config(font=self.current_font)

    def change_colors(self):
        colors = {}
        for keyword in self.highlight_words:
            color = colorchooser.askcolor(title=f"Choose color for {keyword}")[1]
            if color:
                colors[keyword] = color
        self.highlight_words.update(colors)
        self.update_current_log_view(force=True)

    def get_latest_custom_site(self, data_directory):
        try:
            folders = [(folder, os.path.getmtime(os.path.join(data_directory, folder)))
                       for folder in os.listdir(data_directory) if folder.startswith("CustomSite_")]
            if not folders:
                return "No Custom Sites Found"
            latest_folder = max(folders, key=lambda x: x[1])[0]
            return latest_folder
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return "Error"

    def update_data_folder_status(self):
        # This method can be used to update the status of the data folder
        pass

if __name__ == "__main__":
    app = LogViewer()
    app.mainloop()
