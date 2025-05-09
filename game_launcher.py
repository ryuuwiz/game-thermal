import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import subprocess
import os
import json
import time
import threading
import psutil
import sys

try:
    import win32com.client  # Requires: pip install pywin32
except ImportError:
    messagebox.showerror("Missing Dependency", "Please install 'pywin32' using pip:\npip install pywin32")
    sys.exit(1)

# === CONFIGURATION ===
CONFIG_FILE = "games_config.json"
THROTTLESTOP_PATH = r"C:\ThrottleStop\ThrottleStop.exe"  # <-- Replace this
SCRIPT_PATH = os.path.abspath(__file__)


# === UTILITY FUNCTIONS ===

def load_games():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {}

def save_games(games):
    with open(CONFIG_FILE, "w") as f:
        json.dump(games, f, indent=4)

def launch_game(game_name, exe_path):
    def runner():
        try:
            subprocess.Popen([THROTTLESTOP_PATH])
            time.sleep(2)
            game_proc = subprocess.Popen([exe_path])
            while True:
                if not psutil.pid_exists(game_proc.pid):
                    break
                time.sleep(5)
            for proc in psutil.process_iter(["name"]):
                if proc.info["name"] and "ThrottleStop" in proc.info["name"]:
                    proc.kill()
        except Exception as e:
            messagebox.showerror("Error", str(e))
    threading.Thread(target=runner).start()

def create_shortcut(game_name):
    desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
    shortcut_path = os.path.join(desktop, f"{game_name}.lnk")

    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = sys.executable
    shortcut.Arguments = f'"{SCRIPT_PATH}" --game="{game_name}"'
    shortcut.WorkingDirectory = os.path.dirname(SCRIPT_PATH)
    shortcut.IconLocation = SCRIPT_PATH
    shortcut.save()
    messagebox.showinfo("Shortcut Created", f"Shortcut for '{game_name}' created on Desktop.")


# === MAIN GUI APP ===

class GameLauncherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🎮 Game Launcher with ThrottleStop")
        self.root.geometry("480x500")
        self.root.resizable(False, False)

        self.games = load_games()

        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, font=("Segoe UI", 10))
        self.style.configure("Game.TButton", font=("Segoe UI", 9))

        title = ttk.Label(root, text="🎮 Game Launcher with ThrottleStop", font=("Segoe UI", 14, "bold"))
        title.pack(pady=(15, 10))

        self.frame = ttk.Frame(root)
        self.frame.pack(fill="both", expand=True, padx=20)

        canvas = tk.Canvas(self.frame, borderwidth=0)
        scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=canvas.yview)
        self.list_frame = ttk.Frame(canvas)

        self.list_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.list_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.refresh_game_list()

        add_btn = ttk.Button(root, text="➕ Add Game", command=self.add_game)
        add_btn.pack(pady=(10, 20))

    def refresh_game_list(self):
        for widget in self.list_frame.winfo_children():
            widget.destroy()

        if not self.games:
            label = ttk.Label(self.list_frame, text="No games added yet.", font=("Segoe UI", 10, "italic"))
            label.pack(pady=10)
            return

        for game_name, exe_path in self.games.items():
            row = ttk.Frame(self.list_frame)
            row.pack(fill="x", pady=5)

            name_label = ttk.Label(row, text=game_name, width=25)
            name_label.pack(side="left")

            launch_btn = ttk.Button(row, text="▶ Launch", style="Game.TButton",
                                    command=lambda g=game_name, p=exe_path: launch_game(g, p))
            launch_btn.pack(side="left", padx=5)

            shortcut_btn = ttk.Button(row, text="📎 Shortcut", style="Game.TButton",
                                      command=lambda g=game_name: create_shortcut(g))
            shortcut_btn.pack(side="left")

    def add_game(self):
        name = simpledialog.askstring("Game Name", "Enter game name:")
        if not name:
            return
        path = filedialog.askopenfilename(title="Select Game Executable")
        if path:
            self.games[name] = path
            save_games(self.games)
            self.refresh_game_list()


# === MAIN ENTRY POINT ===

def main():
    if "--game=" in " ".join(sys.argv):
        args = " ".join(sys.argv)
        game_name = args.split("--game=")[1].strip('"')
        games = load_games()
        if game_name in games:
            launch_game(game_name, games[game_name])
        else:
            print("Game not found in config.")
    else:
        root = tk.Tk()
        app = GameLauncherApp(root)
        root.mainloop()

if __name__ == "__main__":
    main()
