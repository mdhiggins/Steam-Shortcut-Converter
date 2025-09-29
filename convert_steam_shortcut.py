import os
import sys
import re
from win32com.client import Dispatch
import pythoncom

def read_url_file(url_path):
    """Read the Steam URL file and extract game ID and icon path."""
    try:
        with open(url_path, 'r') as file:
            content = file.read()
        
        # Extract game ID from URL
        game_id_match = re.search(r'steam://rungameid/(\d+)', content)
        game_id = game_id_match.group(1) if game_id_match else None
        
        # Extract icon path
        icon_match = re.search(r'IconFile=(.+)', content)
        icon_path = icon_match.group(1).strip() if icon_match else None
        
        return game_id, icon_path
    except Exception as e:
        print(f"Error reading URL file: {e}")
        return None, None

def get_steam_path(icon_path):
    """Infer Steam executable path from icon path."""
    if not icon_path:
        return None
    steam_dir = os.path.dirname(os.path.dirname(os.path.dirname(icon_path)))

    return os.path.join(steam_dir, "steam.exe")

def create_shortcut(url_path):
    """Create a Windows .lnk shortcut from a Steam URL shortcut."""
    # Get game ID and icon path
    game_id, icon_path = read_url_file(url_path)
    if not game_id or not icon_path:
        print("Failed to extract game ID or icon path.")
        return
    
    # Get Steam executable path
    steam_path = get_steam_path(icon_path)
    if not steam_path or not os.path.exists(steam_path):
        print("Could not determine Steam executable path.")
        return
    
    # Get the directory and name for the new shortcut
    url_dir = os.path.dirname(url_path)
    url_name = os.path.splitext(os.path.basename(url_path))[0]
    shortcut_path = os.path.join(url_dir, f"{url_name}.lnk")
    
    pythoncom.CoInitialize()
    
    try:
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.TargetPath = steam_path
        shortcut.Arguments = f"-applaunch {game_id}"
        shortcut.IconLocation = icon_path
        shortcut.WorkingDirectory = os.path.dirname(steam_path)
        shortcut.save()
        
        print(f"Shortcut created successfully: {shortcut_path}")
    except Exception as e:
        print(f"Error creating shortcut: {e}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Please provide a Steam URL shortcut file as an argument.")
        sys.exit(1)
    
    url_file = sys.argv[1]
    if not os.path.exists(url_file):
        print(f"File not found: {url_file}")
        sys.exit(1)
    
    create_shortcut(url_file)
