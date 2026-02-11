import requests
import os

ICONS = {
    "list-task.png": "https://img.icons8.com/ios-filled/50/0d6efd/list.png", 
    "exclamation-triangle.png": "https://img.icons8.com/ios-filled/50/ffc107/error.png",
    "check-circle.png": "https://img.icons8.com/ios-filled/50/198754/checked-checkbox.png",
    "play-circle.png": "https://img.icons8.com/ios-filled/50/6610f2/play.png",
    "excel.png": "https://img.icons8.com/ios-filled/50/6c757d/microsoft-excel-2019.png",
    "search.png": "https://img.icons8.com/ios-filled/50/6c757d/search--v1.png",
    "logo.png": "https://img.icons8.com/ios-filled/50/0d6efd/futures.png"
}
# Using Icons8 for reliable PNGs as bootstrap raw SVGs need conversion. 
# Colored URLs for matching the theme.

def download_icons():
    if not os.path.exists("assets"):
        os.makedirs("assets")
        
    for name, url in ICONS.items():
        try:
            path = os.path.join("assets", name)
            if not os.path.exists(path):
                print(f"Downloading {name}...")
                r = requests.get(url)
                if r.status_code == 200:
                    with open(path, "wb") as f:
                        f.write(r.content)
                else:
                    print(f"Failed to download {name}: {r.status_code}")
            else:
                print(f"{name} already exists.")
        except Exception as e:
            print(f"Error {name}: {e}")

if __name__ == "__main__":
    download_icons()
