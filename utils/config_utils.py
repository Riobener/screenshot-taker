import os
from datetime import datetime

import yaml


def load_config():
    with open("config.yaml", 'r') as f:
        config = yaml.safe_load(f)
    return config


def get_save_path():
    config = load_config()
    directory_path = config.get("directory_path")
    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    filename = f"Screenshot_{current_time}.png"
    if directory_path == 'desktop':
        save_path = os.path.join(desktop_path, filename)
    else:
        save_path = os.path.join(directory_path, filename)
    return save_path
