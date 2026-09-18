import os
import sys
import subprocess
import win32com.client
import customtkinter as ctk

# Set dark theme appearance
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
if not os.path.exists(CHROME_PATH):
    CHROME_PATH = os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe")

DESKTOP_PATH = os.path.expanduser(r"~\Desktop")

class ChromeTabCreatorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("🚀 Tvůrce Chrome Kombinací Karet")
        self.geometry("700x780")
        self.minsize(600, 650)
        icon_p = os.path.join(os.path.dirname(__file__), "app_icon.ico")
        if os.path.exists(icon_p):
            try:
                self.iconbitmap(icon_p)
            except Exception:
                pass

        self.urls = []

        # Header
        self.header_label = ctk.CTkLabel(
            self, 
            text="🌐 Tvůrce Zástupců pro Google Chrome", 
            font=ctk.CTkFont(size=22, weight="bold")
        )
        self.header_label.pack(pady=(20, 5))

        self.sub_label = ctk.CTkLabel(
            self, 
            text="Vytvoř si jednoduše ikonu na plochu, která na jedno kliknutí otevře zvolené weby.", 
            font=ctk.CTkFont(size=13),
            text_color="gray70"
        )
        self.sub_label.pack(pady=(0, 15))

        # Main Container
        self.main_frame = ctk.CTkFrame(self, corner_radius=12)
        self.main_frame.pack(fill="both", expand=True, padx=25, pady=(0, 15))

        # 1. Shortcut Name
        self.name_label = ctk.CTkLabel(self.main_frame, text="1. Název zástupce na ploše:", font=ctk.CTkFont(size=14, weight="bold"))
        self.name_label.pack(anchor="w", padx=20, pady=(15, 5))

        self.name_entry = ctk.CTkEntry(self.main_frame, placeholder_text="např. Ranní Přehled, AI Nástroje, Práce...", height=38, font=ctk.CTkFont(size=13))
        self.name_entry.pack(fill="x", padx=20, pady=(0, 12))

        # 2. Add URL input
        self.url_label = ctk.CTkLabel(self.main_frame, text="2. Přidat webovou adresu (URL):", font=ctk.CTkFont(size=14, weight="bold"))
        self.url_label.pack(anchor="w", padx=20, pady=(5, 5))

        self.url_input_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.url_input_frame.pack(fill="x", padx=20, pady=(0, 8))

        self.url_entry = ctk.CTkEntry(self.url_input_frame, placeholder_text="https://...", height=38, font=ctk.CTkFont(size=13))
        self.url_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.url_entry.bind("<Return>", lambda event: self.add_url())

        self.add_btn = ctk.CTkButton(self.url_input_frame, text="+ Přidat kartu", width=120, height=38, font=ctk.CTkFont(weight="bold"), command=self.add_url)
        self.add_btn.pack(side="right")

        # Quick Preset Chips
        self.presets_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.presets_frame.pack(fill="x", padx=20, pady=(0, 15))

        presets = [
            ("Gemini", "https://gemini.google.com"),
            ("Claude", "https://claude.ai"),
            ("ChatGPT", "https://chatgpt.com"),
            ("GitHub", "https://github.com"),
            ("YouTube", "https://www.youtube.com"),
            ("ČT24", "https://ct24.ceskatelevize.cz"),
            ("Seznam", "https://www.seznam.cz")
        ]
        for name, url in presets:
            btn = ctk.CTkButton(
                self.presets_frame, 
                text=f"+ {name}", 
                width=70, 
                height=26, 
                font=ctk.CTkFont(size=11),
                fg_color="#2b2b2b",
                hover_color="#3b3b3b",
                command=lambda u=url: self.add_specific_url(u)
            )
            btn.pack(side="left", padx=2)

        # 3. Tab List
        self.list_header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.list_header_frame.pack(fill="x", padx=20, pady=(0, 5))

        self.list_label = ctk.CTkLabel(self.list_header_frame, text="3. Vybrané karty (otevřou se najednou):", font=ctk.CTkFont(size=14, weight="bold"))
        self.list_label.pack(side="left")

        self.clear_btn = ctk.CTkButton(self.list_header_frame, text="Vymazat vše", width=85, height=24, font=ctk.CTkFont(size=11), fg_color="#c0392b", hover_color="#e74c3c", command=self.clear_urls)
        self.clear_btn.pack(side="right")

        self.scroll_frame = ctk.CTkScrollableFrame(self.main_frame, height=180, corner_radius=8)
        self.scroll_frame.pack(fill="both", expand=True, padx=20, pady=(0, 15))

        # 4. Options
        self.options_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.options_frame.pack(fill="x", padx=20, pady=(0, 10))

        self.dark_mode_var = ctk.BooleanVar(value=True)
        self.dark_checkbox = ctk.CTkCheckBox(self.options_frame, text="Vynutit Dark Mode (--force-dark-mode)", variable=self.dark_mode_var)
        self.dark_checkbox.pack(side="left", padx=(0, 20))

        self.new_window_var = ctk.BooleanVar(value=True)
        self.window_checkbox = ctk.CTkCheckBox(self.options_frame, text="Otevřít v novém okně (--new-window)", variable=self.new_window_var)
        self.window_checkbox.pack(side="left")

        # Status Label
        self.status_label = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=13, weight="bold"))
        self.status_label.pack(pady=(0, 8))

        # Action Buttons
        self.action_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.action_frame.pack(fill="x", padx=25, pady=(0, 25))

        self.test_btn = ctk.CTkButton(
            self.action_frame, 
            text="🚀 Otestovat ihned v Chrome", 
            height=44, 
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2980b9",
            hover_color="#3498db",
            command=self.test_launch
        )
        self.test_btn.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.create_btn = ctk.CTkButton(
            self.action_frame, 
            text="⭐ Vytvořit zástupce na Plochu", 
            height=44, 
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#27ae60",
            hover_color="#2ecc71",
            command=self.create_shortcut
        )
        self.create_btn.pack(side="right", fill="x", expand=True, padx=(10, 0))

        # Initial example
        self.add_specific_url("https://gemini.google.com")
        self.add_specific_url("https://github.com")

    def add_url(self):
        url = self.url_entry.get().strip()
        if url:
            if not url.startswith("http://") and not url.startswith("https://"):
                url = "https://" + url
            self.add_specific_url(url)
            self.url_entry.delete(0, "end")

    def add_specific_url(self, url):
        if url not in self.urls:
            self.urls.append(url)
            self.refresh_list()
            self.set_status("", "gray")

    def remove_url(self, url):
        if url in self.urls:
            self.urls.remove(url)
            self.refresh_list()

    def clear_urls(self):
        self.urls.clear()
        self.refresh_list()
        self.set_status("Seznam karet byl vymazán.", "gray")

    def refresh_list(self):
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()

        if not self.urls:
            empty_lbl = ctk.CTkLabel(self.scroll_frame, text="Zatím žádné karty. Přidej URL výše.", text_color="gray50")
            empty_lbl.pack(pady=30)
            return

        for idx, url in enumerate(self.urls, 1):
            row = ctk.CTkFrame(self.scroll_frame, height=36, fg_color="#1e1e1e", corner_radius=6)
            row.pack(fill="x", pady=2, padx=2)

            num_lbl = ctk.CTkLabel(row, text=f"{idx}.", width=25, font=ctk.CTkFont(weight="bold"), text_color="gray70")
            num_lbl.pack(side="left", padx=(10, 5))

            url_lbl = ctk.CTkLabel(row, text=url, anchor="w", font=ctk.CTkFont(size=13))
            url_lbl.pack(side="left", fill="x", expand=True, padx=5)

            del_btn = ctk.CTkButton(
                row, 
                text="✕", 
                width=30, 
                height=26, 
                fg_color="#c0392b", 
                hover_color="#e74c3c", 
                font=ctk.CTkFont(size=12, weight="bold"),
                command=lambda u=url: self.remove_url(u)
            )
            del_btn.pack(side="right", padx=8, pady=4)

    def build_arguments(self):
        args = []
        if self.dark_mode_var.get():
            args.append("--force-dark-mode")
        if self.new_window_var.get():
            args.append("--new-window")
        for u in self.urls:
            args.append(f'"{u}"')
        return " ".join(args)

    def test_launch(self):
        if not self.urls:
            self.set_status("Chyba: Přidej alespoň jednu URL adresu!", "#e74c3c")
            return
        
        args = [CHROME_PATH]
        if self.dark_mode_var.get():
            args.append("--force-dark-mode")
        if self.new_window_var.get():
            args.append("--new-window")
        args.extend(self.urls)

        subprocess.Popen(args)
        self.set_status(f"Otevírám {len(self.urls)} karet v Google Chrome...", "#3498db")

    def create_shortcut(self):
        name = self.name_entry.get().strip()
        if not name:
            self.set_status("Chyba: Zadej název zástupce!", "#e74c3c")
            return
        if not self.urls:
            self.set_status("Chyba: Přidej alespoň jednu URL adresu!", "#e74c3c")
            return

        # Sanitize filename
        safe_name = "".join(c for c in name if c not in r'\/:*?"<>|').strip()
        if not safe_name.lower().endswith(".lnk"):
            safe_name += ".lnk"

        shortcut_path = os.path.join(DESKTOP_PATH, safe_name)

        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(shortcut_path)
            shortcut.TargetPath = CHROME_PATH
            shortcut.Arguments = self.build_arguments()
            shortcut.IconLocation = f"{CHROME_PATH},0"
            shortcut.Description = f"Chrome kombinační zástupce: {name}"
            shortcut.Save()

            self.set_status(f"✅ Zástupce '{safe_name}' byl úspěšně vytvořen na Ploše!", "#2ecc71")
        except Exception as e:
            self.set_status(f"Chyba při vytváření: {str(e)}", "#e74c3c")

    def set_status(self, text, color):
        self.status_label.configure(text=text, text_color=color)

if __name__ == "__main__":
    try:
        app = ChromeTabCreatorApp()
        app.mainloop()
    except Exception as e:
        import traceback
        log_path = os.path.join(os.path.dirname(__file__), "crash.log")
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(traceback.format_exc())

