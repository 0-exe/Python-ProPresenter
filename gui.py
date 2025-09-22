import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import core
import threading
import json
import os

CONFIG_FILE = "config.json"

class ProPresenterGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ProPresenter Playlist Creator")
        self.geometry("800x750")

        # Config variables
        self.api_key = tk.StringVar()
        self.model_name = tk.StringVar(value="gemini-pro")

        # App variables
        self.file_path = tk.StringVar()
        self.playlist_name = tk.StringVar()
        self.translation = tk.StringVar(value='luther')
        self.service_items = []
        self.songs_to_import = []

        self.load_config()
        self.create_widgets()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                    self.api_key.set(config.get("api_key", ""))
                    self.model_name.set(config.get("model_name", "gemini-pro"))
            except (json.JSONDecodeError, TypeError):
                self.log(f"Warning: Could not read {CONFIG_FILE}. File might be corrupt.")

    def save_config(self):
        config = {
            "api_key": self.api_key.get(),
            "model_name": self.model_name.get()
        }
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=4)
        self.log("Settings saved.")

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Settings Frame ---
        settings_frame = ttk.LabelFrame(main_frame, text="Settings")
        settings_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(settings_frame, text="Gemini API Key:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.api_key_entry = ttk.Entry(settings_frame, textvariable=self.api_key, width=60, show="*")
        self.api_key_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)

        ttk.Label(settings_frame, text="Model Name:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Entry(settings_frame, textvariable=self.model_name, width=30).grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Button(settings_frame, text="Save Settings", command=self.save_config).grid(row=1, column=2, padx=5, pady=5)
        settings_frame.columnconfigure(1, weight=1)

        # --- Main App Frame ---
        app_frame = ttk.LabelFrame(main_frame, text="Create Playlist")
        app_frame.pack(fill=tk.X, padx=5, pady=5, margin=10)

        # File selection
        file_frame = ttk.Frame(app_frame)
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(file_frame, text="Word Document:").pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Entry(file_frame, textvariable=self.file_path, width=60).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse...", command=self.select_file).pack(side=tk.LEFT, padx=5, pady=5)

        # Playlist and Translation
        options_frame = ttk.Frame(app_frame)
        options_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(options_frame, text="Playlist Name:").pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Entry(options_frame, textvariable=self.playlist_name, width=40).pack(side=tk.LEFT, padx=5, pady=5)

        ttk.Label(options_frame, text="Psalm Translation:").pack(side=tk.LEFT, padx=5, pady=5)
        translations = ["luther", "kjv", "bbe", "web", "oeb-us"]
        ttk.OptionMenu(options_frame, self.translation, self.translation.get(), *translations).pack(side=tk.LEFT, padx=5, pady=5)

        # --- Log and Buttons ---
        log_frame = ttk.LabelFrame(main_frame, text="Log")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.log_text = tk.Text(log_frame, wrap=tk.WORD, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)

        self.start_button = ttk.Button(button_frame, text="Start", command=self.start_process)
        self.start_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.continue_button = ttk.Button(button_frame, text="Continue", command=self.continue_process, state=tk.DISABLED)
        self.continue_button.pack(side=tk.LEFT, padx=5, pady=5)

        ttk.Button(button_frame, text="Quit", command=self.quit).pack(side=tk.RIGHT, padx=5, pady=5)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if path:
            self.file_path.set(path)
            if not self.playlist_name.get():
                filename = os.path.basename(path).replace('.docx', '')
                self.playlist_name.set(filename)

    def start_process(self):
        if not self.api_key.get():
            messagebox.showerror("Error", "Please enter your Gemini API key in the Settings section and save it.")
            return
        if not self.file_path.get() or not self.playlist_name.get():
            messagebox.showerror("Error", "Please select a file and enter a playlist name.")
            return

        self.log("Starting process...")
        self.start_button.config(state=tk.DISABLED)
        threading.Thread(target=self.parse_docx_thread).start()

    def parse_docx_thread(self):
        self.service_items = core.parse_docx(self.file_path.get(), self.api_key.get(), self.model_name.get())
        if isinstance(self.service_items, str):
            self.log(f"Error: {self.service_items}")
            messagebox.showerror("Error", self.service_items)
            self.start_button.config(state=tk.NORMAL)
            return

        self.songs_to_import = [item["value"] for item in self.service_items if item["type"] == "song"]
        
        if self.songs_to_import:
            self.log("Please import the following songs into ProPresenter from CCLI SongSelect:")
            for song in self.songs_to_import:
                self.log(f"- {song}")
            self.log("\nPress 'Continue' after importing the songs.")
            self.continue_button.config(state=tk.NORMAL)
        else:
            self.log("No songs to import found in the document.")
            self.continue_process()

    def continue_process(self):
        self.continue_button.config(state=tk.DISABLED)
        self.log("Continuing process...")
        threading.Thread(target=self.create_playlist_thread).start()

    def create_playlist_thread(self):
        self.log("Fetching ProPresenter library...")
        library = core.get_library()
        if isinstance(library, str):
            self.log(f"Error: {library}")
            self.start_button.config(state=tk.NORMAL)
            return

        playlist_items = []
        for item in self.service_items:
            if item["type"] == "song":
                playlist_items.append({"type": "song", "name": item["value"]})
            elif item["type"] == "psalm":
                self.log(f"Fetching psalm: {item['value']}...")
                psalm_text = core.get_psalm(item["value"], self.translation.get())
                if "Error" in psalm_text:
                    self.log(psalm_text)
                    continue
                playlist_items.append({"type": "psalm", "name": item["value"], "content": psalm_text})
        
        self.log("Creating ProPresenter playlist...")
        result = core.create_propresenter_playlist(self.playlist_name.get(), playlist_items, library)
        self.log(result)
        self.log("\nProcess finished.")
        self.start_button.config(state=tk.NORMAL)

if __name__ == "__main__":
    app = ProPresenterGUI()
    app.mainloop()
