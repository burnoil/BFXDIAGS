import tkinter as tk
from tkinter import ttk, scrolledtext, simpledialog, filedialog, messagebox, colorchooser
from tkinter.font import Font
import subprocess
import sys
import socket
import os
import glob
import winreg
import shutil
from datetime import datetime
import requests
from requests.auth import HTTPBasicAuth
import threading
import webbrowser

class LogViewer(tk.Tk):
    ICON_PATH = 'log.ico'
    HIGHLIGHT_WORDS = {"ERROR": "red", "WARNING": "yellow", "SUCCESSFUL": "light green"}
    BESCLIENT_PATH = r"C:\Program Files (x86)\BigFix Enterprise\BES Client\besclient.exe"
    BES_DATA_DIRECTORY = r"C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData"
    AGENT_CACHE_DIRECTORY = r"C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\__Cache\Downloads"

    def __init__(self):
        super().__init__()
        self.title("BFXDIAGS 2.2")
        self.geometry("1200x650")
        self.set_icon()
        self.set_fonts()
        self.create_ui()
        self.initialize_variables()
        self.update_initial_status()

    def set_icon(self):
        base_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(__file__)
        self.iconbitmap(os.path.join(base_path, self.ICON_PATH))

    def set_fonts(self):
        self.button_font = Font(family="Tahoma", size=8, weight="bold")
        self.label_font = Font(family="Tahoma", size=8)
        self.current_font = Font(family="Tahoma", size=10)

    def create_ui(self):
        self.create_menu()
        self.create_status_labels()
        self.create_controls()
        self.create_main_layout()

    def initialize_variables(self):
        self.auto_refresh = True
        self.tab_frames = {}
        self.remote_host = None
        self.remote_username = None
        self.remote_password = None

    def update_initial_status(self):
        self.update_besclient_status()
        self.update_relay_status()
        self.update_workstation_info()
        self.update_besclient_version()
        self.update_throttling_status()
        self.update_subsidiary_status()

    def create_menu(self):
        menu_bar = tk.Menu(self)
        self.config(menu=menu_bar)

        help_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="Usage Instructions", command=self.show_usage)

        settings_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="Change Colors", command=self.change_colors)

        remote_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Remote", menu=remote_menu)
        remote_menu.add_command(label="Set Remote Host", command=self.set_remote_host)
        remote_menu.add_command(label="Clear Remote Host", command=self.clear_remote_host)

    def create_status_labels(self):
        self.info_frame = tk.Frame(self)
        self.info_frame.pack(fill=tk.X, pady=3)

        self.status_labels = {
            "besclient": tk.Label(self.info_frame, text="BESClient Status: Unknown", font=self.button_font, bg="lightgray"),
            "tcp_port": tk.Label(self.info_frame, text="TCP Port Status: Unknown", font=self.button_font, bg="lightgray"),
            "udp_port": tk.Label(self.info_frame, text="UDP Port Status: Unknown", font=self.button_font, bg="lightgray"),
            "workstation": tk.Label(self.info_frame, text="Workstation Name: Unknown", font=self.button_font, bg="lightgray"),
            "ip": tk.Label(self.info_frame, text="IP Address: Unknown", font=self.button_font, bg="lightgray"),
            "version": tk.Label(self.info_frame, text="BESClient Version: Unknown", font=self.button_font, bg="lightgray"),
            "subsidiary": tk.Label(self.info_frame, text="Subsidiary: Unknown", font=self.button_font, bg="lightgray"),
        }

        for label in self.status_labels.values():
            label.pack(side=tk.LEFT, padx=3, pady=3)

        self.status_frame = tk.Frame(self)
        self.status_frame.pack(fill=tk.X, pady=3)

        self.status_labels["relay"] = tk.Label(self.status_frame, text="Relay Status: Unknown", font=self.button_font, bg="lightgray")
        self.status_labels["relay"].pack(side=tk.LEFT, padx=3, pady=3)

        self.status_labels["remote"] = tk.Label(self.status_frame, text="Remote Status: Disconnected", font=self.button_font, bg="lightgray")
        self.status_labels["remote"].pack(side=tk.LEFT, padx=3, pady=3)

        self.status_labels["throttling"] = tk.Label(self.status_frame, text="Throttling: Unknown", font=self.button_font, bg="lightgray")
        self.status_labels["throttling"].pack(side=tk.LEFT, padx=3, pady=3)

        self.status_labels["status"] = tk.Label(self.status_frame, text="Status: Ready", font=self.button_font, bg="lightgray")
        self.status_labels["status"].pack(side=tk.LEFT, padx=3, pady=3)

    def create_controls(self):
        self.controls_frame = tk.Frame(self)
        self.controls_frame.pack(fill=tk.Y, side=tk.LEFT)

        self.create_control_buttons()
        self.create_filter_controls()

    def create_control_buttons(self):
        button_specs = [
            ("Pause Refresh", self.toggle_refresh, "lightblue"),
            ("Refresh Now", self.manual_refresh, "lightgreen"),
            ("Search", self.search_in_logs, "lightgrey"),
            ("Select Log File", self.select_log_file, "lightyellow"),
            ("Open Latest BigFix Log", self.open_latest_bigfix_log, "lightblue"),
            ("Monitor BES Data Folder", self.monitor_bes_data_folder, "lightgreen"),
            ("Restart BESClient Service", self.restart_besclient_service, "orange"),
            ("Clear Latest Site Cache", self.clear_latest_site_cache, "red"),
            ("Clear BigFix Agent Cache", self.clear_agent_cache, "red"),
            ("Toggle Throttling", self.toggle_throttling, "lightcoral"),
            ("Edit BigFix Registry", self.open_registry_editor, "lightcoral"),
        ]

        for text, command, bg in button_specs:
            btn = tk.Button(self.controls_frame, text=text, command=command, font=self.button_font, bg=bg)
            btn.pack(fill=tk.X, padx=3, pady=3)
            self.add_tooltip(btn, text)

        self.host_entry = tk.Entry(self.controls_frame, width=12, font=self.button_font)
        self.host_entry.pack(fill=tk.X, padx=3)
        self.host_entry.insert(0, "127.0.0.1")
        self.add_tooltip(self.host_entry, "Enter the host address")

        self.port_entry = tk.Entry(self.controls_frame, width=5, font=self.button_font)
        self.port_entry.pack(fill=tk.X, padx=3)
        self.port_entry.insert(0, "80")
        self.add_tooltip(self.port_entry, "Enter the port number")

        self.protocol_var = tk.StringVar(self)
        self.protocol_var.set("TCP")
        self.protocol_dropdown = tk.OptionMenu(self.controls_frame, self.protocol_var, "TCP", "UDP")
        self.protocol_dropdown.config(font=self.button_font)
        self.protocol_dropdown.pack(fill=tk.X, padx=3, pady=3)
        self.add_tooltip(self.protocol_dropdown, "Select the protocol")

        check_port_btn = tk.Button(self.controls_frame, text="Check Port", command=self.check_port, font=self.button_font, bg="lightcoral")
        check_port_btn.pack(fill=tk.X, padx=3, pady=3)
        self.add_tooltip(check_port_btn, "Check if the port is open")

    def create_filter_controls(self):
        self.filter_frame = tk.Frame(self.controls_frame)
        self.filter_frame.pack(fill=tk.X, pady=3)

        self.filter_entry = tk.Entry(self.filter_frame, width=20, font=self.button_font)
        self.filter_entry.pack(fill=tk.X, padx=3, pady=3)
        self.add_tooltip(self.filter_entry, "Enter filter term")

        self.filter_button = tk.Button(self.filter_frame, text="Filter Logs", command=self.filter_logs, font=self.button_font, bg="lightblue")
        self.filter_button.pack(fill=tk.X, padx=3, pady=3)
        self.add_tooltip(self.filter_button, "Filter the logs by the entered term")

        self.clear_button = tk.Button(self.filter_frame, text="Clear Log", command=self.clear_log_display, font=self.button_font, bg="lightpink")
        self.clear_button.pack(fill=tk.X, padx=3, pady=3)
        self.add_tooltip(self.clear_button, "Clear the log display")

        self.save_button = tk.Button(self.filter_frame, text="Save Filtered Logs", command=self.save_filtered_logs, font=self.button_font, bg="lightgray")
        self.save_button.pack(fill=tk.X, padx=3, pady=3)
        self.add_tooltip(self.save_button, "Save the filtered logs")

        self.font_var = tk.StringVar(self)
        self.font_var.set("Tahoma")
        font_options = ["Tahoma", "Segoe UI", "Arial", "Courier", "Times", "Helvetica"]
        self.font_dropdown = tk.OptionMenu(self.filter_frame, self.font_var, *font_options, command=self.change_font)
        self.font_dropdown.config(font=self.button_font)
        self.font_dropdown.pack(fill=tk.X, padx=3, pady=3)
        self.add_tooltip(self.font_dropdown, "Select the display font")

    def create_main_layout(self):
        self.main_frame = tk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=True, side=tk.RIGHT)

        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.filepath_label = tk.Label(self.main_frame, text="", bd=1, relief=tk.SUNKEN, anchor=tk.W, font=self.label_font)
        self.filepath_label.pack(fill=tk.X, side=tk.BOTTOM, padx=3, pady=3)

    def show_about(self):
        about_window = tk.Toplevel(self)
        about_window.title("About BFXDIAGS")
        about_window.geometry("300x150")
        about_text = tk.Text(about_window, height=10, width=50, bg="light gray", fg="black", wrap="word", borderwidth=2, relief="groove")
        about_text.pack(pady=10, padx=10)
        about_content = "BFXDIAGS\nVersion 2.2\nDeveloped by: "
        about_text.insert("end", about_content)
        about_text.insert("end", "https://github.com/burnoil/PythonProjects", "link")
        about_text.tag_configure("link", foreground="blue", underline=True)
        about_text.tag_bind("link", "<Button-1>", lambda e: webbrowser.open_new("https://github.com/burnoil/PythonProjects"))
        about_text.config(state="disabled", cursor="arrow")

    def show_usage(self):
        usage_text = (
            "Usage Instructions:\n"
            "- To pause the log refresh, click 'Pause Refresh'.\n"
            "- To manually refresh the logs, click 'Refresh Now'.\n"
            "- To search within the logs, click 'Search' and enter the search term.\n"
            "- To restart the BESClient service, click 'Restart BESClient Service'.\n"
            "- To clear the latest site cache, click 'Clear Latest Site Cache'.\n"
            "- To clear the BigFix agent cache, click 'Clear BigFix Agent Cache'.\n"
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

    def set_remote_host(self):
        remote_host = simpledialog.askstring("Remote Host", "Enter the remote host IP address or hostname:")
        remote_username = simpledialog.askstring("Remote Host", "Enter the remote host username:")
        remote_password = simpledialog.askstring("Remote Host", "Enter the remote host password:", show="*")
        if remote_host and remote_username and remote_password:
            self.remote_host = remote_host
            self.remote_username = remote_username
            self.remote_password = remote_password
            self.status_labels["remote"].config(text=f"Remote Status: Connected to {remote_host}")
            messagebox.showinfo("Remote Host", f"Remote host set to {remote_host}")

    def clear_remote_host(self):
        self.remote_host = None
        self.remote_username = None
        self.remote_password = None
        self.status_labels["remote"].config(text="Remote Status: Disconnected")
        messagebox.showinfo("Remote Host", "Remote host cleared")

    def execute_remote_command(self, command):
        if not self.remote_host or not self.remote_username or not self.remote_password:
            messagebox.showerror("Remote Command", "Remote host is not set")
            return None
        try:
            url = f"http://{self.remote_host}:5985/wsman"
            headers = {"Content-Type": "application/soap+xml"}
            body = f"""
            <s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:w="http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd">
                <s:Header>
                    <w:ResourceURI s:mustUnderstand="true">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd</w:ResourceURI>
                    <w:OptionSet>
                        <w:Option Name="WINRS_CONSOLEMODE_STDIN">TRUE</w:Option>
                    </w:OptionSet>
                    <w:OperationTimeout>PT60S</w:OperationTimeout>
                </s:Header>
                <s:Body>
                    <w:CommandLine>
                        <w:Command>{command}</w:Command>
                    </w:CommandLine>
                </s:Body>
            </s:Envelope>
            """
            response = requests.post(url, headers=headers, data=body, auth=HTTPBasicAuth(self.remote_username, self.remote_password), verify=False, timeout=10)
            response.raise_for_status()
            return response.text
        except requests.exceptions.RequestException as e:
            messagebox.showerror("Remote Command", f"Failed to execute remote command: {e}")
            return None

    def update_status_label(self, message, color="lightgray"):
        self.status_labels["status"].config(text=message, bg=color)

    def create_log_tab(self, log_file_path):
        if not os.path.exists(log_file_path):
            messagebox.showerror("Error", f"Log file not found: {log_file_path}")
            return

        tab_frame = tk.Frame(self.notebook)
        tab_frame.log_file_path = log_file_path

        log_area = scrolledtext.ScrolledText(tab_frame, state='disabled', wrap=tk.WORD, font=self.current_font, bg="lightgray")
        log_area.pack(fill=tk.BOTH, expand=True)
        tab_frame.log_area = log_area

        tab_title = os.path.basename(log_file_path)
        self.notebook.add(tab_frame, text=tab_title)
        self.tab_frames[tab_frame] = tab_frame

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
        for keyword, color in self.HIGHLIGHT_WORDS.items():
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
            tab_frame = self.notebook.nametowidget(current_tab)
            if not hasattr(tab_frame, 'log_area'):
                messagebox.showerror("Error", "No log area found in the selected tab.")
                return

            log_area = tab_frame.log_area
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
        if self.remote_host:
            command = "Restart-Service -Name BESClient"
            self.run_threaded_command(lambda: self.execute_remote_command(command), "Restarting BESClient service remotely...")
        else:
            self.run_threaded_command(self.local_restart_besclient, "Restarting BESClient service locally...")

    def local_restart_besclient(self):
        try:
            self.update_status_label("Stopping BESClient service...", "orange")
            stop_result = subprocess.run(["net", "stop", "BESClient"], capture_output=True, text=True, shell=True)
            if stop_result.returncode == 0:
                self.update_status_label("BESClient service stopped successfully.", "lightgreen")
            else:
                self.update_status_label(f"Failed to stop BESClient service: {stop_result.stderr}", "red")
                return

            self.update_status_label("Starting BESClient service...", "orange")
            start_result = subprocess.run(["net", "start", "BESClient"], capture_output=True, text=True, shell=True)
            if start_result.returncode == 0:
                self.update_status_label("BESClient service started successfully!", "lightgreen")
            else:
                self.update_status_label(f"Failed to start BESClient service: {start_result.stderr}", "red")
                return
        except subprocess.CalledProcessError as e:
            self.update_status_label(f"Failed to restart BESClient service: {str(e)}", "red")
        except Exception as e:
            self.update_status_label(f"An error occurred: {str(e)}", "red")
        self.update_besclient_status()

    def clear_latest_site_cache(self):
        command = """
        $latestSite = Get-ChildItem -Path 'C:\\Program Files (x86)\\BigFix Enterprise\\BES Client\\__BESData' | 
                      Where-Object { $_.PSIsContainer -and $_.Name -like 'CustomSite_*' } | 
                      Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($latestSite) {
            Remove-Item -Path $latestSite.FullName -Recurse -Force
            New-Item -ItemType Directory -Path $latestSite.FullName
            Write-Output 'Cache cleared successfully!'
        } else {
            Write-Output 'No Custom Sites found.'
        }
        """
        if self.remote_host:
            self.run_threaded_command(lambda: self.execute_remote_command(command), "Clearing latest site cache remotely...")
        else:
            self.run_threaded_command(self.local_clear_latest_site_cache, "Clearing latest site cache locally...")

    def local_clear_latest_site_cache(self):
        try:
            latest_custom_site = self.get_latest_custom_site(self.BES_DATA_DIRECTORY)
            if latest_custom_site == "No Custom Sites Found":
                messagebox.showwarning("Warning", "No Custom Sites found to clear cache.")
                return
            cache_directory = os.path.join(self.BES_DATA_DIRECTORY, latest_custom_site)
            self.update_status_label(f"Clearing cache for {latest_custom_site}...")
            if os.path.exists(cache_directory):
                shutil.rmtree(cache_directory)
                os.makedirs(cache_directory)
            self.update_status_label(f"Cache for {latest_custom_site} cleared successfully!")
        except Exception as e:
            self.update_status_label(f"Failed to clear cache: {str(e)}")
        self.update_data_folder_status()

    def clear_agent_cache(self):
        command = """
        $agentCache = 'C:\\Program Files (x86)\\BigFix Enterprise\\BES Client\\__BESData\\__Global\\__Cache\\Downloads'
        Remove-Item -Path $agentCache -Recurse -Force
        New-Item -ItemType Directory -Path $agentCache
        Write-Output 'Agent cache cleared successfully!'
        """
        if self.remote_host:
            self.run_threaded_command(lambda: self.execute_remote_command(command), "Clearing agent cache remotely...")
        else:
            self.run_threaded_command(self.local_clear_agent_cache, "Clearing agent cache locally...")

    def local_clear_agent_cache(self):
        try:
            self.update_status_label("Clearing BigFix agent cache...")
            if os.path.exists(self.AGENT_CACHE_DIRECTORY):
                shutil.rmtree(self.AGENT_CACHE_DIRECTORY)
                os.makedirs(self.AGENT_CACHE_DIRECTORY)
            self.update_status_label("BigFix agent cache cleared successfully!")
        except Exception as e:
            self.update_status_label(f"Failed to clear agent cache: {str(e)}")

    def check_port(self):
        host = self.host_entry.get()
        port_str = self.port_entry.get()
        protocol = self.protocol_var.get()
        try:
            port = int(port_str)
            self.run_threaded_command(lambda: self.check_port_status(host, port, protocol), f"Checking {protocol} port {port} on {host}...")
        except ValueError:
            self.update_status_label("Invalid port number.", "yellow")

    def check_port_status(self, host, port, protocol):
        is_open = self.is_port_open(host, port, protocol)
        label_key = "tcp_port" if protocol == "TCP" else "udp_port"
        if is_open:
            self.status_labels[label_key].config(text=f"{protocol} port {port} on {host} is open.", bg="lightgreen", fg="black")
        else:
            self.status_labels[label_key].config(text=f"{protocol} port {port} on {host} is closed.", bg="red", fg="white")
        return is_open

    def is_port_open(self, host, port, protocol):
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM if protocol == "UDP" else socket.SOCK_STREAM)
        sock.settimeout(1)
        try:
            if protocol == "TCP":
                sock.connect((host, port))
            else:
                sock.sendto(b"", (host, port))
                sock.recvfrom(1024)
            sock.close()
            return True
        except (socket.timeout, socket.error):
            return False

    def select_log_file(self):
        path = filedialog.askopenfilename(
            title="Select a log file",
            filetypes=[("All files", "*.*"), ("Text files", "*.txt"), ("Log files", "*.log")],
            defaultextension="*.*"
        )
        if path:
            self.create_log_tab(path)

    def open_latest_bigfix_log(self):
        if self.remote_host:
            self.run_threaded_command(self.find_and_open_latest_log_remote, "Opening latest BigFix log remotely...")
        else:
            self.run_threaded_command(self.find_and_open_latest_log, "Opening latest BigFix log...")

    def find_and_open_latest_log(self):
        bigfix_log_directory = os.path.join("C:\\", "Program Files (x86)", "BigFix Enterprise", "BES Client", "__BESData", "__Global", "Logs")
        log_files = glob.glob(os.path.join(bigfix_log_directory, "*.log"))

        if not log_files:
            self.update_status_label("No BigFix log files found.", "red")
            return

        latest_log_file = max(log_files, key=os.path.getmtime)
        self.create_log_tab(latest_log_file)

    def find_and_open_latest_log_remote(self):
        bigfix_log_directory = r"C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs"
        command = f"Get-ChildItem -Path '{bigfix_log_directory}' -Filter '*.log' | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | ForEach-Object {{ $_.FullName }}"
        latest_log_file = self.execute_remote_command(command).strip()

        if not latest_log_file:
            self.update_status_label("No BigFix log files found.", "red")
            return

        self.run_on_ui_thread(lambda: self.create_log_tab(latest_log_file))

    def run_on_ui_thread(self, func):
        self.after(0, func)

    def monitor_bes_data_folder(self):
        if not os.path.exists(self.BES_DATA_DIRECTORY):
            messagebox.showwarning("Warning", "BES Data directory not found.")
            return
        self.create_data_folder_tab(self.BES_DATA_DIRECTORY)

    def create_data_folder_tab(self, data_directory):
        tab_frame = tk.Frame(self.notebook)
        tab_frame.data_directory = data_directory

        log_area = scrolledtext.ScrolledText(tab_frame, state='normal', wrap=tk.WORD, font=self.current_font, bg="lightgray")
        log_area.pack(fill=tk.BOTH, expand=True)
        tab_frame.log_area = log_area

        tab_title = "BES Data Folder"
        self.notebook.add(tab_frame, text=tab_title)
        self.tab_frames[tab_frame] = tab_frame

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
            folders.sort(key=lambda x: x[1], reverse=True)
            for folder, mtime in folders[:10]:
                mod_time = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
                log_area.insert(tk.END, f"{folder} - Last Modified: {mod_time}\n")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        log_area.config(state='disabled')
        if self.auto_refresh:
            self.after(10000, lambda: self.update_data_folder_view(log_area, data_directory))

    def update_besclient_status(self):
        if self.remote_host:
            command = "Get-Service -Name BESClient"
            self.run_threaded_command(lambda: self.execute_remote_command(command), "Getting BESClient status remotely...")
        else:
            self.run_threaded_command(self.local_get_besclient_status, "Getting BESClient status locally...")

    def local_get_besclient_status(self):
        status = self.get_service_status("besclient")
        port_status_tcp = self.is_port_open("127.0.0.1", 52311, "TCP")
        port_status_udp = self.is_port_open("127.0.0.1", 52311, "UDP")
        self.status_labels["besclient"].config(text=f"BESClient Status: {status}")
        self.status_labels["tcp_port"].config(text=f"TCP Port 52311 Status: {'Open' if port_status_tcp else 'Closed'}")
        self.status_labels["udp_port"].config(text=f"UDP Port 52311 Status: {'Open' if port_status_udp else 'Closed'}")
        
        if status == "Running":
            self.status_labels["besclient"].config(bg="lightgreen", fg="black")
        else:
            self.status_labels["besclient"].config(bg="red", fg="white")

        self.update_port_status("tcp_port", port_status_tcp)
        self.update_port_status("udp_port", port_status_udp)

        self.after(5000, self.update_besclient_status)

    def update_port_status(self, label_key, is_open):
        if is_open:
            self.status_labels[label_key].config(bg="lightgreen", fg="black")
        else:
            self.status_labels[label_key].config(bg="red", fg="white")

    def get_service_status(self, service_name):
        try:
            result = subprocess.run(["sc", "query", service_name], capture_output=True, text=True, shell=True)
            return "Running" if "RUNNING" in result.stdout else "Stopped" if "STOPPED" in result.stdout else "Unknown"
        except Exception as e:
            return f"Error: {str(e)}"

    def update_workstation_info(self):
        if self.remote_host:
            command = "hostname"
            self.run_threaded_command(lambda: self.execute_remote_command(command), "Getting workstation info remotely...")
        else:
            self.run_threaded_command(self.local_get_workstation_info, "Getting workstation info locally...")

    def local_get_workstation_info(self):
        hostname = socket.gethostname()
        self.status_labels["workstation"].config(text=f"Workstation Name: {hostname}")
        ip_addresses = self.get_ip_addresses()
        self.status_labels["ip"].config(text=f"IP Address(es): {', '.join(ip_addresses)}")

    def get_ip_addresses(self):
        ip_addresses = []
        hostname = socket.gethostname()
        for ip in socket.getaddrinfo(hostname, None):
            if ':' not in ip[4][0]:
                ip_addresses.append(ip[4][0])
        return ip_addresses

    def update_besclient_version(self):
        if self.remote_host:
            command = """
            $path = 'C:\\Program Files (x86)\\BigFix Enterprise\\BES Client\\besclient.exe'
            (Get-Item $path).VersionInfo.ProductVersion
            """
            self.run_threaded_command(lambda: self.execute_remote_command(command), "Getting BESClient version remotely...")
        else:
            self.run_threaded_command(self.local_get_besclient_version, "Getting BESClient version locally...")

    def local_get_besclient_version(self):
        try:
            if os.path.exists(self.BESCLIENT_PATH):
                path = self.BESCLIENT_PATH.replace("\\", "\\\\")
                result = subprocess.run(['wmic', 'datafile', 'where', f'name="{path}"', 'get', 'Version', '/value'], capture_output=True, text=True)
                version_info = result.stdout.strip().split('=')[-1]
                self.status_labels["version"].config(text=f"BESClient Version: {version_info}")
            else:
                 self.status_labels["version"].config(text="BESClient Version: Not Found")
        except Exception as e:
            self.status_labels["version"].config(text=f"BESClient Version: Error ({str(e)})")

    def update_relay_status(self):
        if self.remote_host:
            command = """
            $regPath1 = 'HKLM:\\SOFTWARE\\WOW6432Node\\BigFix\\EnterpriseClient\\Settings\\Client\\__RelayServer1'
            $regPath2 = 'HKLM:\\SOFTWARE\\WOW6432Node\\BigFix\\EnterpriseClient\\Settings\\Client\\__RelayServer2'
            $relayServer1 = Get-ItemProperty -Path $regPath1 -Name 'value'
            $relayServer2 = Get-ItemProperty -Path $regPath2 -Name 'value'
            Write-Output "$($relayServer1.value), $($relayServer2.value)"
            """
            self.run_threaded_command(lambda: self.execute_remote_command(command), "Getting relay status remotely...")
        else:
            self.run_threaded_command(self.local_get_relay_status, "Getting relay status locally...")

    def local_get_relay_status(self):
        relay_info = self.get_relay_info()
        self.status_labels["relay"].config(text=f"Relay Status: {relay_info}")
        self.after(5000, self.update_relay_status)

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
        if self.remote_host:
            command = """
            $regPath = 'HKLM:\\SOFTWARE\\WOW6432Node\\BigFix\\EnterpriseClient\\Settings\\Client\\ThrottlingExempt'
            $throttlingStatus = Get-ItemProperty -Path $regPath -Name 'value'
            Write-Output $throttlingStatus.value
            """
            self.run_threaded_command(lambda: self.execute_remote_command(command), "Getting throttling status remotely...")
        else:
            self.run_threaded_command(self.local_get_throttling_status, "Getting throttling status locally...")

    def local_get_throttling_status(self):
        try:
            registry_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\\WOW6432Node\\BigFix\\EnterpriseClient\\Settings\\Client\\ThrottlingExempt")
            throttling_status, _ = winreg.QueryValueEx(registry_key, "value")
            winreg.CloseKey(registry_key)
            self.status_labels["throttling"].config(text=f"Throttling: {throttling_status}")
            return throttling_status
        except FileNotFoundError:
            self.status_labels["throttling"].config(text="Throttling: Not Found")
            return "Not Found"
        except Exception as e:
            self.status_labels["throttling"].config(text=f"Throttling: Error ({str(e)})")
            return "Error"

    def toggle_throttling(self):
        if self.remote_host:
            self.run_threaded_command(self.toggle_throttling_remote, "Toggling throttling remotely...")
        else:
            self.run_threaded_command(self.local_toggle_throttling, "Toggling throttling locally...")

    def toggle_throttling_remote(self):
        current_status_command = """
        $regPath = 'HKLM:\\SOFTWARE\\WOW6432Node\\BigFix\\EnterpriseClient\\Settings\\Client\\ThrottlingExempt'
        (Get-ItemProperty -Path $regPath -Name 'value').value
        """
        current_status = self.execute_remote_command(current_status_command).strip()
        new_status = "NO" if current_status == "YES" else "YES"
        command = f"""
        $regPath = 'HKLM:\\SOFTWARE\\WOW6432Node\\BigFix\\EnterpriseClient\\Settings\\Client\\ThrottlingExempt'
        Set-ItemProperty -Path $regPath -Name 'value' -Value '{new_status}'
        Write-Output '{new_status}'
        """
        output = self.execute_remote_command(command)
        self.run_on_ui_thread(lambda: self.status_labels["throttling"].config(text=f"Throttling: {output.strip()}"))

    def local_toggle_throttling(self):
        try:
            current_status = self.local_get_throttling_status()
            new_status = "NO" if current_status == "YES" else "YES"
            registry_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\\WOW6432Node\\BigFix\\EnterpriseClient\\Settings\\Client\\ThrottlingExempt", 0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(registry_key, "value", 0, winreg.REG_SZ, new_status)
            winreg.CloseKey(registry_key)
            self.status_labels["throttling"].config(text=f"Throttling: {new_status}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to toggle throttling: {str(e)}")

    def update_subsidiary_status(self):
        if self.remote_host:
            command = """
            $regPath = 'HKLM:\\SOFTWARE\\WOW6432Node\\BigFix\\EnterpriseClient\\Settings\\Client\\Subsidiary'
            $subsidiaryStatus = Get-ItemProperty -Path $regPath -Name 'value'
            Write-Output $subsidiaryStatus.value
            """
            self.run_threaded_command(lambda: self.execute_remote_command(command), "Getting subsidiary status remotely...")
        else:
            self.run_threaded_command(self.local_get_subsidiary_status, "Getting subsidiary status locally...")

    def local_get_subsidiary_status(self):
        try:
            registry_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\\WOW6432Node\\BigFix\\EnterpriseClient\\Settings\\Client\\Subsidiary")
            subsidiary_status, _ = winreg.QueryValueEx(registry_key, "value")
            winreg.CloseKey(registry_key)
            self.status_labels["subsidiary"].config(text=f"Subsidiary: {subsidiary_status}")
        except FileNotFoundError:
            self.status_labels["subsidiary"].config(text="Subsidiary: Not Found")
        except Exception as e:
            self.status_labels["subsidiary"].config(text=f"Subsidiary: Error ({str(e)})")

    def open_registry_editor(self):
        self.registry_editor_window = tk.Toplevel(self)
        self.registry_editor_window.title("BigFix Registry Editor")
        self.registry_editor_window.geometry("800x600")

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
        root_key = r"SOFTWARE\\WOW6432Node\\BigFix"
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
        self.update_current_log_view(force=True)
        if self.auto_refresh:
            self.update_status_label("Log refresh resumed.", "lightgreen")
            self.pause_button.config(text="Pause Refresh")
        else:
            self.update_status_label("Log refresh paused.", "orange")
            self.pause_button.config(text="Resume Refresh")

    def manual_refresh(self):
        self.update_current_log_view(force=True)

    def update_current_log_view(self, force=False):
        current_tab = self.notebook.select()
        tab_frame = self.notebook.nametowidget(current_tab)
        if not hasattr(tab_frame, 'log_area'):
            messagebox.showerror("Error", "No log area found in the selected tab.")
            return

        log_area = tab_frame.log_area
        log_file_path = tab_frame.log_file_path

        if (self.auto_refresh or force) and log_file_path:
            try:
                if self.remote_host:
                    self.update_log_view_remote(log_area, log_file_path)
                else:
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
            tab_frame = self.notebook.nametowidget(current_tab)
            if not hasattr(tab_frame, 'log_area'):
                messagebox.showerror("Error", "No log area found in the selected tab.")
                return

            log_area = tab_frame.log_area
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
        tab_frame = self.notebook.nametowidget(current_tab)
        if not hasattr(tab_frame, 'log_area'):
            messagebox.showerror("Error", "No log area found in the selected tab.")
            return

        log_area = tab_frame.log_area
        log_area.config(state='normal')
        log_area.delete('1.0', tk.END)
        log_area.config(state='disabled')

    def save_filtered_logs(self):
        current_tab = self.notebook.select()
        tab_frame = self.notebook.nametowidget(current_tab)
        if not hasattr(tab_frame, 'log_area'):
            messagebox.showerror("Error", "No log area found in the selected tab.")
            return

        log_area = tab_frame.log_area
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
        for tab_frame in self.tab_frames.values():
            tab_frame.log_area.config(font=self.current_font)
        self.values_text.config(font=self.current_font)

    def change_colors(self):
        colors = {}
        for keyword in self.HIGHLIGHT_WORDS:
            color = colorchooser.askcolor(title=f"Choose color for {keyword}")[1]
            if color:
                colors[keyword] = color
        self.HIGHLIGHT_WORDS.update(colors)
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

    def run_threaded_command(self, func, status_message):
        def wrapper():
            self.update_status_label(status_message)
            result = func()
            self.update_status_label("Operation completed.")
            return result

        threading.Thread(target=wrapper).start()

    def add_tooltip(self, widget, text):
        tooltip = tk.Toplevel(widget, bg='black', padx=1, pady=1)
        tooltip.withdraw()
        tooltip.overrideredirect(True)
        tk.Label(tooltip, text=text, bg='yellow', padx=2).pack()

        def enter(event):
            tooltip.deiconify()
            tooltip.geometry(f"+{event.x_root+10}+{event.y_root+10}")

        def leave(event):
            tooltip.withdraw()

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

if __name__ == "__main__":
    app = LogViewer()
    app.mainloop()