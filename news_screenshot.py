import os
import re
import sys
import json
import base64
import subprocess
import threading
import webbrowser
from urllib.parse import quote, quote_plus, urlparse
from tkinter import Tk, Label, Entry, Button, filedialog, StringVar, messagebox, simpledialog
from PIL import Image, ImageDraw, ImageFont
from docx import Document
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from io import BytesIO

# ---------------- Paths for executable-friendly ----------------
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

LOGO_FOLDER = os.path.join(BASE_DIR, "NewsLogos")
CUSTOM_LOGO_NAME = "custom"
NOTE_FOLDER = os.path.join(BASE_DIR, "Note")
os.makedirs(NOTE_FOLDER, exist_ok=True)
MISSING_LOG_FILE = os.path.join(NOTE_FOLDER, "missing_logos.txt")
SETTINGS_FILE = os.path.join(NOTE_FOLDER, "settings.json")
SUPPORTED_LOGO_EXTENSIONS = (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif", ".tif", ".tiff", ".ico", ".jfif")
APP_VERSION = "1.1.0"
UPDATE_INFO_URL = "https://raw.githubusercontent.com/deepndense-sketch/PrintNews/main/version.json"
GITHUB_REPO_OWNER = "deepndense-sketch"
GITHUB_REPO_NAME = "PrintNews"
GITHUB_BRANCH = "main"
GITHUB_LOGO_API_URL = f"https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/contents/NewsLogos?ref={GITHUB_BRANCH}"
GITHUB_CONTENTS_API_URL = f"https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/contents"
GITHUB_TOKEN_SETTINGS_KEY = "github_token"

# ---------------- Config ----------------
PADDING_TOP = 20
PADDING_BOTTOM = 20
MARGIN = 20
FONT_SIZE_HEADLINE = 48
FONT_SIZE_DATE = 24
FONT_SIZE_SOURCE = 20
LOGO_HEIGHT = 60
MAX_FILENAME_WORDS = 8
LINE_SPACING = 10
RIGHT_MARGIN = 25
MIN_WIDTH = 800
MAX_WIDTH = 1500
SUB_HEAD_COLOR = "#003366"  # dark blue for sub headline
GAP_BETWEEN_SEGMENTS = 20  # extra gap for // segments


def version_parts(version):
    parts = []
    for piece in re.findall(r"\d+", version or ""):
        parts.append(int(piece))
    return parts or [0]


def is_newer_version(remote_version, current_version):
    remote_parts = version_parts(remote_version)
    current_parts = version_parts(current_version)
    length = max(len(remote_parts), len(current_parts))
    remote_parts += [0] * (length - len(remote_parts))
    current_parts += [0] * (length - len(current_parts))
    return remote_parts > current_parts


def fetch_update_info():
    response = requests.get(UPDATE_INFO_URL, timeout=6)
    response.raise_for_status()
    info = json.loads(response.content.decode("utf-8-sig"))
    latest_version = str(info.get("version", "")).strip()
    return {
        "version": latest_version,
        "notes": str(info.get("notes", "")).strip(),
        "download_url": str(info.get("download_url", "")).strip(),
        "is_newer": bool(latest_version and is_newer_version(latest_version, APP_VERSION)),
    }


def check_for_updates(show_current=False):
    try:
        info = fetch_update_info()
        latest_version = info["version"]
        if info["is_newer"]:
            message = f"Update available to version {latest_version}.\n\nCurrent version: {APP_VERSION}\nLatest version: {latest_version}"
            if info["notes"]:
                message += f"\n\nWhat is updated:\n{info['notes']}"
            if info["download_url"]:
                message += f"\n\nDownload/update link:\n{info['download_url']}"
            messagebox.showinfo("Update Available", message)
        elif show_current:
            latest_label = latest_version or "unknown"
            messagebox.showinfo("No Update", f"No update available.\n\nCurrent version: {APP_VERSION}\nLatest version: {latest_label}")
    except Exception:
        if show_current:
            messagebox.showwarning("Update Check Failed", "Could not check for updates right now.")


update_info = None


def update_button_from_info(info=None, error=None):
    global update_info
    if error:
        update_button.config(text="Update check failed")
        return
    update_info = info
    latest_version = info.get("version") or APP_VERSION
    if info.get("is_newer"):
        update_button.config(text=f"Update to {latest_version}")
    else:
        update_button.config(text=f"Latest version {latest_version}")


def run_update_button_check():
    try:
        info = fetch_update_info()
        root.after(0, update_button_from_info, info, None)
    except Exception as e:
        root.after(0, update_button_from_info, None, e)


def start_update_button_check():
    update_button.config(text="Checking update...")
    threading.Thread(target=run_update_button_check, daemon=True).start()


def install_update(info):
    download_url = info.get("download_url")
    latest_version = info.get("version")
    if not download_url:
        root.after(0, messagebox.showwarning, "Update Failed", "The update file link is missing.")
        return
    if not getattr(sys, "frozen", False):
        root.after(0, messagebox.showinfo, "Update Available", f"Update available to version {latest_version}.\n\nDownload link:\n{download_url}")
        return

    try:
        root.after(0, update_button.config, {"text": f"Downloading {latest_version}..."})
        response = requests.get(download_url, timeout=60)
        response.raise_for_status()
        exe_path = sys.executable
        new_exe_path = exe_path + ".new"
        updater_path = os.path.join(BASE_DIR, "apply_update.bat")
        with open(new_exe_path, "wb") as f:
            f.write(response.content)
        bat = f"""@echo off
timeout /t 2 /nobreak > nul
move /y "{new_exe_path}" "{exe_path}" > nul
start "" "{exe_path}"
del "%~f0"
"""
        with open(updater_path, "w", encoding="ascii") as f:
            f.write(bat)
        subprocess.Popen(["cmd", "/c", updater_path], cwd=BASE_DIR)
        os._exit(0)
    except Exception as e:
        root.after(0, update_button.config, {"text": f"Update to {latest_version}"})
        root.after(0, messagebox.showwarning, "Update Failed", f"Could not update to version {latest_version}.\n\n{e}")


def get_github_token():
    return (
        os.environ.get("PRINTNEWS_GITHUB_TOKEN")
        or os.environ.get("GITHUB_TOKEN")
        or os.environ.get("GH_TOKEN")
        or str(load_settings().get(GITHUB_TOKEN_SETTINGS_KEY, "")).strip()
    )


def github_headers(token=None):
    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": "PrintNews",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    if token:
        headers["Authorization"] = f"Bearer {token}"
    return headers


def upload_logo_to_github(filename, token):
    path_for_api = quote(f"NewsLogos/{filename}", safe="/")
    api_url = f"{GITHUB_CONTENTS_API_URL}/{path_for_api}"
    local_path = os.path.join(LOGO_FOLDER, filename)
    with open(local_path, "rb") as f:
        encoded_content = base64.b64encode(f.read()).decode("ascii")

    payload = {
        "message": f"Add logo {filename}",
        "content": encoded_content,
        "branch": GITHUB_BRANCH,
    }
    response = requests.put(api_url, headers=github_headers(token), json=payload, timeout=20)
    response.raise_for_status()


def sync_logos_with_github():
    try:
        os.makedirs(LOGO_FOLDER, exist_ok=True)
        token = get_github_token()
        local_files = [
            name for name in os.listdir(LOGO_FOLDER)
            if name.lower().endswith(SUPPORTED_LOGO_EXTENSIONS)
            and os.path.isfile(os.path.join(LOGO_FOLDER, name))
        ]
        existing = {name.lower() for name in local_files}
        response = requests.get(GITHUB_LOGO_API_URL, headers=github_headers(token), timeout=10)
        response.raise_for_status()
        remote_files = response.json()
        remote_names = {
            item.get("name", "").lower()
            for item in remote_files
            if item.get("name", "").lower().endswith(SUPPORTED_LOGO_EXTENSIONS)
        }
        downloaded = []
        uploaded = []
        skipped = 0

        for item in remote_files:
            filename = item.get("name", "")
            if not filename.lower().endswith(SUPPORTED_LOGO_EXTENSIONS):
                continue
            if filename.lower() in existing:
                skipped += 1
                continue
            download_url = item.get("download_url")
            if not download_url:
                continue

            logo_response = requests.get(download_url, headers={"User-Agent": "PrintNews"}, timeout=20)
            logo_response.raise_for_status()
            save_path = os.path.join(LOGO_FOLDER, filename)
            if os.path.exists(save_path):
                skipped += 1
                existing.add(filename.lower())
                continue
            with open(save_path, "wb") as f:
                f.write(logo_response.content)
            existing.add(filename.lower())
            downloaded.append(filename)

        upload_skipped = 0
        upload_errors = []
        if token:
            for filename in local_files:
                if filename.lower() in remote_names:
                    upload_skipped += 1
                    continue
                try:
                    upload_logo_to_github(filename, token)
                    remote_names.add(filename.lower())
                    uploaded.append(filename)
                except requests.HTTPError as e:
                    response = getattr(e, "response", None)
                    status_code = response.status_code if response is not None else "unknown"
                    upload_errors.append(f"{filename}: GitHub returned {status_code}")
                except Exception as e:
                    upload_errors.append(f"{filename}: {e}")
        else:
            upload_errors.append("Click Set GitHub Token once on this computer to upload local logos to GitHub.")

        upload_error = None
        if upload_errors:
            upload_error = "\n".join(upload_errors[:10])
            if len(upload_errors) > 10:
                upload_error += f"\n...and {len(upload_errors) - 10} more"

        return downloaded, uploaded, skipped, upload_skipped, upload_error, None
    except Exception as e:
        return [], [], 0, 0, None, e


def show_logo_sync_result(downloaded, uploaded, skipped, upload_skipped, upload_error, error):
    if error:
        messagebox.showwarning("Logo Sync Failed", f"Could not sync logos from GitHub.\n\n{error}")
        return

    message = (
        "Logo sync complete.\n\n"
        f"Downloaded from GitHub: {len(downloaded)}\n"
        f"Uploaded to GitHub: {len(uploaded)}\n"
        f"Already had locally: {skipped}\n"
        f"Already on GitHub: {upload_skipped}"
    )
    if downloaded:
        message += "\n\nDownloaded logos:\n" + "\n".join(downloaded[:30])
        if len(downloaded) > 30:
            message += f"\n...and {len(downloaded) - 30} more"
    if uploaded:
        message += "\n\nUploaded logos:\n" + "\n".join(uploaded[:30])
        if len(uploaded) > 30:
            message += f"\n...and {len(uploaded) - 30} more"
    if upload_error:
        message += f"\n\nUpload skipped:\n{upload_error}"
    messagebox.showinfo("Logo Sync", message)


def run_logo_sync_thread():
    downloaded, uploaded, skipped, upload_skipped, upload_error, error = sync_logos_with_github()
    root.after(0, show_logo_sync_result, downloaded, uploaded, skipped, upload_skipped, upload_error, error)


def start_logo_sync():
    threading.Thread(target=run_logo_sync_thread, daemon=True).start()


def check_updates_clicked():
    if update_info and update_info.get("is_newer"):
        threading.Thread(target=install_update, args=(update_info,), daemon=True).start()
    else:
        start_update_button_check()
# ---------------- GUI ----------------
file_path = None
output_path = None


def load_settings():
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(word_file, output_folder):
    try:
        os.makedirs(NOTE_FOLDER, exist_ok=True)
        settings_data = load_settings()
        settings_data["last_word_file"] = word_file
        settings_data["last_output_folder"] = output_folder
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings_data, f, indent=2)
    except Exception as e:
        print("Could not save settings:", e)


def save_github_token(token):
    try:
        os.makedirs(NOTE_FOLDER, exist_ok=True)
        settings_data = load_settings()
        settings_data[GITHUB_TOKEN_SETTINGS_KEY] = token
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings_data, f, indent=2)
        return True
    except Exception as e:
        messagebox.showwarning("Token Not Saved", f"Could not save GitHub token.\n\n{e}")
        return False


def set_github_token():
    token = simpledialog.askstring(
        "Set GitHub Token",
        "Paste a GitHub token with Contents read/write permission for this repo.\n\n"
        "It will be saved on this computer only.",
        show="*",
        parent=root,
    )
    if token is None:
        return
    token = token.strip()
    if not token:
        messagebox.showwarning("Missing Token", "No token was entered.")
        return
    if save_github_token(token):
        messagebox.showinfo("GitHub Token Saved", "Done. Sync Logos can now upload from this computer.")


def browse_word():
    initial_file = word_var.get().strip()
    initial_dir = os.path.dirname(initial_file) if initial_file else BASE_DIR
    if not os.path.isdir(initial_dir):
        initial_dir = BASE_DIR
    path = filedialog.askopenfilename(initialdir=initial_dir, filetypes=[("Word Documents", "*.docx")])
    if path:
        word_var.set(path)
        if not output_var.get().strip():
            output_var.set(os.path.dirname(path))


def browse_output():
    initial_dir = output_var.get().strip() or os.path.dirname(word_var.get().strip()) or BASE_DIR
    if not os.path.isdir(initial_dir):
        initial_dir = BASE_DIR
    path = filedialog.askdirectory(initialdir=initial_dir, title="Choose PNG render folder")
    if path:
        output_var.set(path)


def run_app():
    global file_path, output_path
    file_path = word_var.get().strip()
    output_path = output_var.get().strip()
    if not file_path:
        messagebox.showwarning("Missing Word File", "Please select a Word file.")
        return
    if not output_path:
        output_path = os.path.dirname(file_path)
        output_var.set(output_path)
    if not os.path.isdir(output_path):
        messagebox.showwarning("Missing Render Folder", "Please choose a render folder for the PNG files.")
        return
    save_settings(file_path, output_path)
    root.quit()

settings = load_settings()

root = Tk()
root.title(f"News Image Generator v{APP_VERSION}")
root.geometry("620x245")
root.resizable(False, False)

word_var = StringVar(value=settings.get("last_word_file", ""))
output_var = StringVar(value=settings.get("last_output_folder", ""))

Label(root, text="Word File:").place(x=20, y=20)
Entry(root, textvariable=word_var, width=58).place(x=120, y=20)
Button(root, text="Browse", command=browse_word).place(x=520, y=16)

Label(root, text="Render Folder:").place(x=20, y=62)
Entry(root, textvariable=output_var, width=58).place(x=120, y=62)
Button(root, text="Browse", command=browse_output).place(x=520, y=58)

Button(root, text="Sync Logos", width=18, command=start_logo_sync).place(x=120, y=112)
update_button = Button(root, text="Checking update...", width=22, command=check_updates_clicked)
update_button.place(x=330, y=112)
Button(root, text="Set GitHub Token", width=18, command=set_github_token).place(x=120, y=155)
Button(root, text="Run", width=20, command=run_app).place(x=235, y=195)
root.after(500, start_update_button_check)

root.mainloop()

if not file_path:
    print("No Word file selected. Exiting.")
    sys.exit()

INPUT_FOLDER = os.path.dirname(file_path)
OUTPUT_FOLDER = output_path

# ---------------- Helpers ----------------
def resize_logo(img):
    ratio = LOGO_HEIGHT / img.height
    new_width = int(img.width * ratio)
    return img.resize((new_width, LOGO_HEIGHT), Image.LANCZOS)

def text_width(font, text):
    return font.getbbox(text)[2]

def wrap_headline(headline, main_font, sub_font, max_width):
    segments = headline.split("//")
    lines = []
    first_segment = True
    for seg in segments:
        seg = seg.strip()
        if not seg:
            continue
        font = main_font if first_segment else sub_font
        color = "black" if first_segment else SUB_HEAD_COLOR
        words = seg.split()
        line = ""
        for word in words:
            test_line = word if line == "" else line + " " + word
            if text_width(font, test_line) + 2*MARGIN > max_width:
                if line:
                    lines.append((line, font, color))
                line = word
            else:
                line = test_line
        if line:
            lines.append((line, font, color))
        # Add extra gap after each // segment except the first
        if not first_segment:
            lines.append(("", sub_font, SUB_HEAD_COLOR))
        first_segment = False
    return lines

def normalize_date(d):
    if re.fullmatch(r'\d{8}', d):
        return f"{d[:4]}-{d[4:6]}-{d[6:]}"
    return d

def base_logo_name(name):
    cleaned = (name or "Unknown").strip()
    return cleaned.split(".", 1)[0] if "." in cleaned else cleaned


def missing_logo_note_name(name):
    cleaned = (name or "Unknown").strip()
    parts = cleaned.split(".") if cleaned else []
    return ".".join(parts[:-1]) if len(parts) > 1 else cleaned


def preferred_logo_name(name):
    cleaned = (name or "Unknown").strip()
    preferred = missing_logo_note_name(cleaned)
    return preferred or cleaned


def logo_name_candidates(name):
    names = []
    cleaned = (name or "").strip()
    base_name = base_logo_name(cleaned)
    note_name = missing_logo_note_name(cleaned)
    parts = cleaned.split(".") if cleaned else []
    country_name = base_name + "-" + parts[-1] if len(parts) > 2 else ""
    for candidate in (cleaned, note_name, base_name, country_name, base_name + ".com"):
        if candidate:
            names.append(candidate)
    return list(dict.fromkeys(names))


def missing_logo_name(source):
    return missing_logo_note_name(source)


def missing_logo_search_name(source):
    return (source or "Unknown").strip()


def find_logo_path(name):
    if not os.path.isdir(LOGO_FOLDER):
        return None

    preferred_name = preferred_logo_name(name)
    expected = {
        f"{candidate}{ext}".lower()
        for candidate in logo_name_candidates(name)
        for ext in SUPPORTED_LOGO_EXTENSIONS
    }
    for filename in os.listdir(LOGO_FOLDER):
        if filename.lower() in expected:
            current_path = os.path.join(LOGO_FOLDER, filename)
            current_base, current_ext = os.path.splitext(filename)
            if preferred_name and current_base != preferred_name:
                preferred_filename = f"{preferred_name}{current_ext}"
                preferred_path = os.path.join(LOGO_FOLDER, preferred_filename)
                if not os.path.exists(preferred_path):
                    try:
                        os.replace(current_path, preferred_path)
                        print(f"Renamed logo: {filename} -> {preferred_filename}")
                        return preferred_path
                    except Exception:
                        pass
            return current_path
    return None


def open_logo(path):
    with Image.open(path) as img:
        return resize_logo(img.convert("RGBA"))


def get_logo(source, url=None):
    path = find_logo_path(source)
    if path:
        return open_logo(path), False

    if url:
        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--headless=new")
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            driver.get(url)
            driver.implicitly_wait(2)
            soup = BeautifulSoup(driver.page_source, "html.parser")
            for selector in ["logo", "site-logo", "header-logo"]:
                img_tag = soup.find("img", {"class": selector})
                if img_tag and img_tag.get("src"):
                    src = img_tag["src"]
                    if src.startswith("//"):
                        src = "https:" + src
                    elif src.startswith("/"):
                        parsed = urlparse(url)
                        src = f"{parsed.scheme}://{parsed.netloc}{src}"
                    resp = requests.get(src, timeout=5)
                    logo_img = Image.open(BytesIO(resp.content)).convert("RGBA")
                    logo_img.thumbnail((LOGO_HEIGHT*5, LOGO_HEIGHT*5))
                    driver.quit()
                    return resize_logo(logo_img), False
            driver.quit()
        except:
            try: driver.quit()
            except: pass

    custom_path = find_logo_path(CUSTOM_LOGO_NAME)
    if custom_path:
        return open_logo(custom_path), True

    img = Image.new("RGBA", (LOGO_HEIGHT, LOGO_HEIGHT), (255,255,255,255))
    return img, True

# ---------------- Fonts ----------------
HEADLINE_FONT_FILES = [
    "arialbd.ttf",
    "ARIALNB.TTF",
    "impact.ttf",
    "bahnschrift.ttf",
    "calibrib.ttf",
    "cambriab.ttf",
    "Candarab.ttf",
    "corbelb.ttf",
    "georgiab.ttf",
    "segoeuib.ttf",
    "tahomabd.ttf",
    "timesbd.ttf",
    "trebucbd.ttf",
    "verdanab.ttf",
    "SourceSansPro-Bold.otf",
    "SourceSansPro-Semibold.otf",
    "malgunbd.ttf",
    "msjhbd.ttc",
    "msyhbd.ttc",
    "comicbd.ttf",
]


def load_font(font_file, size):
    font_path = os.path.join("C:/Windows/Fonts", font_file)
    try:
        return ImageFont.truetype(font_path, size)
    except Exception:
        try:
            return ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", size)
        except Exception:
            return ImageFont.load_default()


headline_font_pairs = [
    (load_font(font_file, FONT_SIZE_HEADLINE), load_font(font_file, int(FONT_SIZE_HEADLINE * 0.8)))
    for font_file in HEADLINE_FONT_FILES
]


def headline_fonts_for_index(index):
    return headline_font_pairs[index % len(headline_font_pairs)]


def headline_fonts_for_source(source, source_font_indices, next_index):
    source_key = (source or "Unknown").strip().lower()
    if source_key not in source_font_indices:
        source_font_indices[source_key] = next_index
        next_index += 1
    return headline_fonts_for_index(source_font_indices[source_key]), next_index


font_head, font_sub_head = headline_fonts_for_index(0)
font_date = load_font("arial.ttf", FONT_SIZE_DATE)
font_source = load_font("arial.ttf", FONT_SIZE_SOURCE)


def open_missing_logo_searches(missing_search_sources):
    for source in sorted(missing_search_sources):
        query = quote_plus(f"{source} logo")
        webbrowser.open_new_tab(f"https://www.google.com/search?tbm=isch&q={query}")

# ---------------- Process Word ----------------
doc = Document(file_path)
missing_sources = set()
missing_search_sources = set()
headline_index = 0
source_font_indices = {}

for table in doc.tables:
    for row in table.rows:
        try:
            date_raw = row.cells[0].text.strip()
            number = row.cells[1].text.strip()
            headline = row.cells[2].text.strip()
            url = row.cells[3].text.strip()
            parsed = urlparse(url) if url else None
            source = parsed.netloc.replace("www.", "") if parsed else "Unknown"

            if not headline:
                continue
            (font_head, font_sub_head), headline_index = headline_fonts_for_source(
                source,
                source_font_indices,
                headline_index,
            )

            # filename
            words = headline.split()[:MAX_FILENAME_WORDS]
            name_base = re.sub(r'[\/:*?"<>|]', '', f"{number} {' '.join(words)}")[:120]
            name = name_base

            # source
            # get logo
            logo, used_fallback = get_logo(source, url)
            if used_fallback:
                missing_sources.add(missing_logo_name(source))
                missing_search_sources.add(missing_logo_search_name(source))

            save_path = os.path.join(OUTPUT_FOLDER, f"{name}.png")
            date = normalize_date(date_raw)

            # wrap headline with // handled
            width = MIN_WIDTH
            for _ in range(10):
                lines = wrap_headline(headline, font_head, font_sub_head, width)
                if len(lines) <= 2 or width >= MAX_WIDTH:
                    break
                width += 50

            text_h = sum(f.size + LINE_SPACING for _, f, _ in lines)
            logo_h = logo.height
            height = PADDING_TOP + logo_h + 20 + text_h + PADDING_BOTTOM

            # compute width
            max_line_width = max(text_width(f, l) for l, f, c in lines)
            date_source_width = text_width(font_date, date) + 10 + text_width(font_source, f" | {source}")
            logo_plus_spacing = logo.width + 20
            final_w = max(max_line_width + MARGIN, logo_plus_spacing + date_source_width + MARGIN)
            final_w += RIGHT_MARGIN

            # create image
            img = Image.new("RGB", (final_w, height), "white")
            draw = ImageDraw.Draw(img)
            y = PADDING_TOP

            # paste logo
            img.paste(logo, (MARGIN, y), logo)
            logo_bottom = y + logo.height

            # draw date + source
            dx = MARGIN + logo.width + 20
            dy = y + (logo.height - FONT_SIZE_DATE)
            draw.text((dx, dy), date, font=font_date, fill="black")
            dw = text_width(font_date, date)
            draw.text((dx + dw + 10, dy), f" | {source}", font=font_source, fill="black")

            # draw headline
            y_text = logo_bottom + 20
            prev_font = None
            for line, f, color in lines:
                if prev_font and f != prev_font:
                    y_text += GAP_BETWEEN_SEGMENTS
                draw.text((MARGIN, y_text), line, font=f, fill=color)
                y_text += f.size + LINE_SPACING
                prev_font = f

            img.save(save_path)
            print("Saved:", save_path)

        except Exception as e:
            print("Error:", e)

# ---------------- Save missing logos ----------------
with open(MISSING_LOG_FILE, "w", encoding="utf-8") as f:
    for s in sorted(missing_sources):
        f.write(s + "\n")
if missing_sources:
    print(f"Missing logos saved to: {MISSING_LOG_FILE}")
    open_missing_logo_searches(missing_search_sources)

print("Render is done.")
try:
    messagebox.showinfo("Render Complete", "Render is done.")
except Exception:
    pass
















