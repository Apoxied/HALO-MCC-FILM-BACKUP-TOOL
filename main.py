import os
import shutil
import string
import tkinter as tk
import tkinter.messagebox as messagebox
import tkinter.messagebox

import requests as requests
import win32api
import time

# Set up tkinter root
root = tk.Tk()
root.geometry("230x200")
root.title("MCC Films")
root.configure(bg='grey30')
root.resizable(False, False)

# Retrieve the image from a link online
image_url = "https://imgur.com/lyzvhPv.png"
response = requests.get(image_url)

# Open the image using tkinter
image = tk.PhotoImage(data=response.content)

# Display the image as an icon
root.tk.call('wm', 'iconphoto', root._w, image)


def copy_recent_h3_theater_file():
    search_dir = os.path.expanduser("~/AppData/LocalLow/MCC/Temporary/UserContent/Halo3/Movie")
    desktop_folder = os.path.expanduser("~/Desktop")
    try:
        most_recent_file = max(os.listdir(search_dir), key=lambda f: os.stat(os.path.join(search_dir, f)).st_mtime)
    except ValueError:
        tkinter.messagebox.showerror("Error", "You don't have any recent theater files for this game. Be sure to load the theater menu to populate your most recent films.")
        return
    if not os.path.exists(search_dir):
        print("Cannot locate theater files :(")
        return
    i = 1
    while True:
        backup_folder = os.path.join(desktop_folder, f'Film Backup ({i})')
        if not os.path.exists(backup_folder):
            os.makedirs(backup_folder)
            break
        i += 1
    shutil.copy(os.path.join(search_dir, most_recent_file), backup_folder)
    tkinter.messagebox.showinfo("Success!", "A folder has been created on your desktop.")
    base, extension = os.path.splitext(most_recent_file)
    base = base[4:-9]
    new_file_name = base + extension
    new_file_path = os.path.join(backup_folder, new_file_name)
    os.rename(os.path.join(backup_folder, most_recent_file), new_file_path)
    search_dirs = []
    for drive in string.ascii_uppercase:
        full_path = drive + ':\\'
        if os.path.exists(full_path):
            for path in [r'\Program Files (x86)\Steam\steamapps\common\Halo The Master Chief Collection\halo3\maps', r'\SteamLibrary\steamapps\common\Halo The Master Chief Collection\halo3\maps', r'\SteamLibary\steamapps\common\Halo The Master Chief Collection\halo3\maps']:
                full_path = os.path.join(drive + ':', path)
                if os.path.exists(full_path):
                    search_dirs.append(full_path)
                    break
    similar_file = None
    for dir in search_dirs:
        if not os.path.exists(dir):
            continue
        similar_files = [f for f in os.listdir(dir) if f.startswith(base)]
        if len(similar_files) > 0:
            similar_file = similar_files[0]
            similar_file_path = os.path.join(dir, similar_file)
            break
    if similar_file is not None:
        shutil.copy(similar_file_path, backup_folder)
    else:
        tkinter.messagebox.showerror("ERROR!", "Cannot locate map file.")

    dll_drive_letters = [d + ":" for d in string.ascii_uppercase]

    dll_dirs_to_search = []
    for drive_letter in dll_drive_letters:
        dll_dirs_to_search.extend([
            os.path.join(drive_letter, r'\Program Files (x86)\Steam\steamapps\common\Halo The Master Chief Collection\halo3'),
            os.path.join(drive_letter, r'SteamLibrary\steamapps\common\Halo The Master Chief Collection\halo3'),
            os.path.join(drive_letter, r'SteamLibary\steamapps\common\Halo The Master Chief Collection\halo3')
        ])

    dll_file = None
    for dll_dir in dll_dirs_to_search:
        if os.path.exists(dll_dir):
            dll_files = [f for f in os.listdir(dll_dir) if f.endswith('.dll')]
            if len(dll_files) > 0:
                dll_file = dll_files[0]
                dll_file_path = os.path.join(dll_dir, dll_file)
                break

    if dll_file is not None:
        shutil.copy(dll_file_path, backup_folder)
    else:
        tkinter.messagebox.showerror("ERROR!", "Cannot locate DLL file.")

    # Do not return here so that the script continues even if DLL is not found

    # Generate a list of drive letters from A to Z
    drive_letters = [drive + ':' for drive in string.ascii_uppercase if os.path.exists(drive + ':')]

    # Define the search directories
    paths_to_search = [
        r'\Program Files (x86)\Steam\steamapps\common\Halo The Master Chief Collection\mcc\binaries\win64',
        r'\SteamLibrary\steamapps\common\Halo The Master Chief Collection\mcc\binaries\win64',
        r'\SteamLibary\steamapps\common\Halo The Master Chief Collection\mcc\binaries\win64'
    ]

    # Find the file in the directories
    version = None
    for drive in drive_letters:
        for path in paths_to_search:
            full_path = os.path.join(drive, path)
            if os.path.exists(full_path):
                file_path = os.path.join(full_path, 'MCC-Win64-Shipping.exe')
                if os.path.exists(file_path):
                    info = win32api.GetFileVersionInfo(file_path, '\\')
                    ms = info['FileVersionMS']
                    ls = info['FileVersionLS']
                    version = '{}.{}.{}.{}'.format(win32api.HIWORD(ms), win32api.LOWORD(ms), win32api.HIWORD(ls),
                                                   win32api.LOWORD(ls))
                    break
        if version is not None:
            break

    if version is None:
        print("File not found")
    else:
        most_recent_file_path = os.path.join(search_dir, most_recent_file)

        # Get the file creation time
        file_time = os.path.getctime(most_recent_file_path)
        creation_time = time.ctime(file_time)

        # Save the file version and creation time to a text file in the backup folder
        file_version_file = os.path.join(backup_folder, 'version_information.txt')
        with open(file_version_file, 'w') as f:
            f.write(version + '\n\n')
            f.write('Film Created On: ' + creation_time)

def copy_recent_h3odst_theater_file():
    search_dir = os.path.expanduser("~/AppData/LocalLow/MCC/Temporary/UserContent/Halo3ODST/Movie")
    desktop_folder = os.path.expanduser("~/Desktop")
    try:
        most_recent_file = max(os.listdir(search_dir), key=lambda f: os.stat(os.path.join(search_dir, f)).st_mtime)
    except ValueError:
        tkinter.messagebox.showerror("Error", "You don't have any recent theater files for this game. Be sure to load the theater menu to populate your most recent films.")
        return
    if not os.path.exists(search_dir):
        print("Cannot locate theater files :(")
        return
    i = 1
    while True:
        backup_folder = os.path.join(desktop_folder, f'Film Backup ({i})')
        if not os.path.exists(backup_folder):
            os.makedirs(backup_folder)
            break
        i += 1
    shutil.copy(os.path.join(search_dir, most_recent_file), backup_folder)
    tkinter.messagebox.showinfo("Success!", "A folder has been created on your desktop.")
    base, extension = os.path.splitext(most_recent_file)
    base = base[4:-9]
    new_file_name = base + extension
    new_file_path = os.path.join(backup_folder, new_file_name)
    os.rename(os.path.join(backup_folder, most_recent_file), new_file_path)
    search_dirs = []
    for drive in string.ascii_uppercase:
        full_path = drive + ':\\'
        if os.path.exists(full_path):
            for path in [r'\Program Files (x86)\Steam\steamapps\common\Halo The Master Chief Collection\halo3odst\maps', r'\SteamLibrary\steamapps\common\Halo The Master Chief Collection\halo3odst\maps', r'\SteamLibary\steamapps\common\Halo The Master Chief Collection\halo3odst\maps']:
                full_path = os.path.join(drive + ':', path)
                if os.path.exists(full_path):
                    search_dirs.append(full_path)
                    break
    similar_file = None
    for dir in search_dirs:
        if not os.path.exists(dir):
            continue
        similar_files = [f for f in os.listdir(dir) if f.startswith(base)]
        if len(similar_files) > 0:
            similar_file = similar_files[0]
            similar_file_path = os.path.join(dir, similar_file)
            break
    if similar_file is not None:
        shutil.copy(similar_file_path, backup_folder)
    else:
        tkinter.messagebox.showerror("ERROR!", "Cannot locate map file.")

    dll_drive_letters = [d + ":" for d in string.ascii_uppercase]

    dll_dirs_to_search = []
    for drive_letter in dll_drive_letters:
        dll_dirs_to_search.extend([
            os.path.join(drive_letter, r'\Program Files (x86)\Steam\steamapps\common\Halo The Master Chief Collection\halo3odst'),
            os.path.join(drive_letter, r'SteamLibrary\steamapps\common\Halo The Master Chief Collection\halo3odst'),
            os.path.join(drive_letter, r'SteamLibary\steamapps\common\Halo The Master Chief Collection\halo3odst')
        ])

    dll_file = None
    for dll_dir in dll_dirs_to_search:
        if os.path.exists(dll_dir):
            dll_files = [f for f in os.listdir(dll_dir) if f.endswith('.dll')]
            if len(dll_files) > 0:
                dll_file = dll_files[0]
                dll_file_path = os.path.join(dll_dir, dll_file)
                break

    if dll_file is not None:
        shutil.copy(dll_file_path, backup_folder)
    else:
        tkinter.messagebox.showerror("ERROR!", "Cannot locate DLL file.")

    # Do not return here so that the script continues even if DLL is not found

    # Generate a list of drive letters from A to Z
    drive_letters = [drive + ':' for drive in string.ascii_uppercase if os.path.exists(drive + ':')]

    # Define the search directories
    paths_to_search = [
        r'\Program Files (x86)\Steam\steamapps\common\Halo The Master Chief Collection\mcc\binaries\win64',
        r'\SteamLibrary\steamapps\common\Halo The Master Chief Collection\mcc\binaries\win64',
        r'\SteamLibary\steamapps\common\Halo The Master Chief Collection\mcc\binaries\win64'
    ]

    # Find the file in the directories
    version = None
    for drive in drive_letters:
        for path in paths_to_search:
            full_path = os.path.join(drive, path)
            if os.path.exists(full_path):
                file_path = os.path.join(full_path, 'MCC-Win64-Shipping.exe')
                if os.path.exists(file_path):
                    info = win32api.GetFileVersionInfo(file_path, '\\')
                    ms = info['FileVersionMS']
                    ls = info['FileVersionLS']
                    version = '{}.{}.{}.{}'.format(win32api.HIWORD(ms), win32api.LOWORD(ms), win32api.HIWORD(ls),
                                                   win32api.LOWORD(ls))
                    break
        if version is not None:
            break

    if version is None:
        print("File not found")
    else:
        most_recent_file_path = os.path.join(search_dir, most_recent_file)

        # Get the file creation time
        file_time = os.path.getctime(most_recent_file_path)
        creation_time = time.ctime(file_time)

        # Save the file version and creation time to a text file in the backup folder
        file_version_file = os.path.join(backup_folder, 'version_information.txt')
        with open(file_version_file, 'w') as f:
            f.write(version + '\n\n')
            f.write('Film Created On: ' + creation_time)


def copy_recent_haloreach_theater_file():
    search_dir = os.path.expanduser("~/AppData/LocalLow/MCC/Temporary/UserContent/HaloReach/Movie")
    desktop_folder = os.path.expanduser("~/Desktop")
    try:
        most_recent_file = max(os.listdir(search_dir), key=lambda f: os.stat(os.path.join(search_dir, f)).st_mtime)
    except ValueError:
        tkinter.messagebox.showerror("Error", "You don't have any recent theater files for this game. Be sure to load the theater menu to populate your most recent films.")
        return
    if not os.path.exists(search_dir):
        print("Cannot locate theater files :(")
        return
    i = 1
    while True:
        backup_folder = os.path.join(desktop_folder, f'Film Backup ({i})')
        if not os.path.exists(backup_folder):
            os.makedirs(backup_folder)
            break
        i += 1
    shutil.copy(os.path.join(search_dir, most_recent_file), backup_folder)
    tkinter.messagebox.showinfo("Success!", "A folder has been created on your desktop.")
    base, extension = os.path.splitext(most_recent_file)
    base = base[12:-9]
    new_file_name = base + extension
    new_file_path = os.path.join(backup_folder, new_file_name)
    os.rename(os.path.join(backup_folder, most_recent_file), new_file_path)
    search_dirs = []
    for drive in string.ascii_uppercase:
        full_path = drive + ':\\'
        if os.path.exists(full_path):
            for path in [r'\Program Files (x86)\Steam\steamapps\common\Halo The Master Chief Collection\haloreach\maps', r'\SteamLibrary\steamapps\common\Halo The Master Chief Collection\haloreach\maps', r'\SteamLibary\steamapps\common\Halo The Master Chief Collection\haloreach\maps']:
                full_path = os.path.join(drive + ':', path)
                if os.path.exists(full_path):
                    search_dirs.append(full_path)
                    break
    similar_file = None
    for dir in search_dirs:
        if not os.path.exists(dir):
            continue
        similar_files = [f for f in os.listdir(dir) if f.startswith(base)]
        if len(similar_files) > 0:
            similar_file = similar_files[0]
            similar_file_path = os.path.join(dir, similar_file)
            break
    if similar_file is not None:
        shutil.copy(similar_file_path, backup_folder)
    else:
        tkinter.messagebox.showerror("ERROR!", "Cannot locate map file.")

    dll_drive_letters = [d + ":" for d in string.ascii_uppercase]

    dll_dirs_to_search = []
    for drive_letter in dll_drive_letters:
        dll_dirs_to_search.extend([
            os.path.join(drive_letter, r'\Program Files (x86)\Steam\steamapps\common\Halo The Master Chief Collection\halo3reach'),
            os.path.join(drive_letter, r'SteamLibrary\steamapps\common\Halo The Master Chief Collection\halo3reach'),
            os.path.join(drive_letter, r'SteamLibary\steamapps\common\Halo The Master Chief Collection\halo3reach')
        ])

    dll_file = None
    for dll_dir in dll_dirs_to_search:
        if os.path.exists(dll_dir):
            dll_files = [f for f in os.listdir(dll_dir) if f.endswith('.dll')]
            if len(dll_files) > 0:
                dll_file = dll_files[0]
                dll_file_path = os.path.join(dll_dir, dll_file)
                break

    if dll_file is not None:
        shutil.copy(dll_file_path, backup_folder)
    else:
        tkinter.messagebox.showerror("ERROR!", "Cannot locate DLL file.")

    # Do not return here so that the script continues even if DLL is not found

    # Generate a list of drive letters from A to Z
    drive_letters = [drive + ':' for drive in string.ascii_uppercase if os.path.exists(drive + ':')]

    # Define the search directories
    paths_to_search = [
        r'\Program Files (x86)\Steam\steamapps\common\Halo The Master Chief Collection\mcc\binaries\win64',
        r'\SteamLibrary\steamapps\common\Halo The Master Chief Collection\mcc\binaries\win64',
        r'\SteamLibary\steamapps\common\Halo The Master Chief Collection\mcc\binaries\win64'
    ]

    # Find the file in the directories
    version = None
    for drive in drive_letters:
        for path in paths_to_search:
            full_path = os.path.join(drive, path)
            if os.path.exists(full_path):
                file_path = os.path.join(full_path, 'MCC-Win64-Shipping.exe')
                if os.path.exists(file_path):
                    info = win32api.GetFileVersionInfo(file_path, '\\')
                    ms = info['FileVersionMS']
                    ls = info['FileVersionLS']
                    version = '{}.{}.{}.{}'.format(win32api.HIWORD(ms), win32api.LOWORD(ms), win32api.HIWORD(ls),
                                                   win32api.LOWORD(ls))
                    break
        if version is not None:
            break

    if version is None:
        print("File not found")
    else:
        most_recent_file_path = os.path.join(search_dir, most_recent_file)

        # Get the file creation time
        file_time = os.path.getctime(most_recent_file_path)
        creation_time = time.ctime(file_time)

        # Save the file version and creation time to a text file in the backup folder
        file_version_file = os.path.join(backup_folder, 'version_information.txt')
        with open(file_version_file, 'w') as f:
            f.write(version + '\n\n')
            f.write('Film Created On: ' + creation_time)

def copy_recent_h4_theater_file():
    search_dir = os.path.expanduser("~/AppData/LocalLow/MCC/Temporary/UserContent/Halo4/Movie")
    desktop_folder = os.path.expanduser("~/Desktop")
    try:
        most_recent_file = max(os.listdir(search_dir), key=lambda f: os.stat(os.path.join(search_dir, f)).st_mtime)
    except ValueError:
        tkinter.messagebox.showerror("Error", "You don't have any recent theater files for this game. Be sure to load the theater menu to populate your most recent films.")
        return
    if not os.path.exists(search_dir):
        print("Cannot locate theater files :(")
        return
    i = 1
    while True:
        backup_folder = os.path.join(desktop_folder, f'Film Backup ({i})')
        if not os.path.exists(backup_folder):
            os.makedirs(backup_folder)
            break
        i += 1
    shutil.copy(os.path.join(search_dir, most_recent_file), backup_folder)
    tkinter.messagebox.showinfo("Success!", "A folder has been created on your desktop.")
    base, extension = os.path.splitext(most_recent_file)
    base = base[12:-9]
    new_file_name = base + extension
    new_file_path = os.path.join(backup_folder, new_file_name)
    os.rename(os.path.join(backup_folder, most_recent_file), new_file_path)
    search_dirs = []
    for drive in string.ascii_uppercase:
        full_path = drive + ':\\'
        if os.path.exists(full_path):
            for path in [r'\Program Files (x86)\Steam\steamapps\common\Halo The Master Chief Collection\halo4\maps', r'\SteamLibrary\steamapps\common\Halo The Master Chief Collection\halo4\maps', r'\SteamLibary\steamapps\common\Halo The Master Chief Collection\halo4\maps']:
                full_path = os.path.join(drive + ':', path)
                if os.path.exists(full_path):
                    search_dirs.append(full_path)
                    break
    similar_file = None
    for dir in search_dirs:
        if not os.path.exists(dir):
            continue
        similar_files = [f for f in os.listdir(dir) if f.startswith(base)]
        if len(similar_files) > 0:
            similar_file = similar_files[0]
            similar_file_path = os.path.join(dir, similar_file)
            break
    if similar_file is not None:
        shutil.copy(similar_file_path, backup_folder)
    else:
        tkinter.messagebox.showerror("ERROR!", "Cannot locate map file.")

    dll_drive_letters = [d + ":" for d in string.ascii_uppercase]

    dll_dirs_to_search = []
    for drive_letter in dll_drive_letters:
        dll_dirs_to_search.extend([
            os.path.join(drive_letter, r'\Program Files (x86)\Steam\steamapps\common\Halo The Master Chief Collection\halo4'),
            os.path.join(drive_letter, r'SteamLibrary\steamapps\common\Halo The Master Chief Collection\halo4'),
            os.path.join(drive_letter, r'SteamLibary\steamapps\common\Halo The Master Chief Collection\halo4')
        ])

    dll_file = None
    for dll_dir in dll_dirs_to_search:
        if os.path.exists(dll_dir):
            dll_files = [f for f in os.listdir(dll_dir) if f.endswith('.dll')]
            if len(dll_files) > 0:
                dll_file = dll_files[0]
                dll_file_path = os.path.join(dll_dir, dll_file)
                break

    if dll_file is not None:
        shutil.copy(dll_file_path, backup_folder)
    else:
        tkinter.messagebox.showerror("ERROR!", "Cannot locate DLL file.")

    # Do not return here so that the script continues even if DLL is not found

    # Generate a list of drive letters from A to Z
    drive_letters = [drive + ':' for drive in string.ascii_uppercase if os.path.exists(drive + ':')]

    # Define the search directories
    paths_to_search = [
        r'\Program Files (x86)\Steam\steamapps\common\Halo The Master Chief Collection\mcc\binaries\win64',
        r'\SteamLibrary\steamapps\common\Halo The Master Chief Collection\mcc\binaries\win64',
        r'\SteamLibary\steamapps\common\Halo The Master Chief Collection\mcc\binaries\win64'
    ]

    # Find the file in the directories
    version = None
    for drive in drive_letters:
        for path in paths_to_search:
            full_path = os.path.join(drive, path)
            if os.path.exists(full_path):
                file_path = os.path.join(full_path, 'MCC-Win64-Shipping.exe')
                if os.path.exists(file_path):
                    info = win32api.GetFileVersionInfo(file_path, '\\')
                    ms = info['FileVersionMS']
                    ls = info['FileVersionLS']
                    version = '{}.{}.{}.{}'.format(win32api.HIWORD(ms), win32api.LOWORD(ms), win32api.HIWORD(ls),
                                                   win32api.LOWORD(ls))
                    break
        if version is not None:
            break

    if version is None:
        print("File not found")
    else:
        most_recent_file_path = os.path.join(search_dir, most_recent_file)

        # Get the file creation time
        file_time = os.path.getctime(most_recent_file_path)
        creation_time = time.ctime(file_time)

        # Save the file version and creation time to a text file in the backup folder
        file_version_file = os.path.join(backup_folder, 'version_information.txt')
        with open(file_version_file, 'w') as f:
            f.write(version + '\n\n')
            f.write('Film Created On: ' + creation_time)

# noinspection PyShadowingNames
def create_open_close_buttons(root, x, y, open_command, close_command, open_text, close_text):
    open_button = tk.Button(root, text=open_text, command=open_command, bg="purple3", fg="white", cursor="hand2",
                            font=("arcadia", 12, "bold"))
    open_button.pack()
    open_button.place(x=x, y=y)

    # Calculate the width of the open_button widget
    open_27 = open_button.winfo_reqwidth()

    close_button = tk.Button(root, text=close_text, command=close_command, bg="white", fg="black", cursor="hand2",
                             font=("arcadia", 12, "bold"))
    close_button.pack()
    # Position the close_button widget to the right of the open_button widget
    close_button.place(x=x + open_27, y=y)


# noinspection PyShadowingNames
def yellow_button(root, x, y, open_command, close_command, open_text, close_text):
    # Calculate the width of the yellow_button widgets
    yellow_button1_width = tk.Button(root, text=open_text, command=open_command, bg="yellow", fg="black",
                                     cursor="hand2", font=("arcadia", 12, "bold")).winfo_reqwidth()
    yellow_button2_width = tk.Button(root, text=close_text, command=close_command, bg="yellow", fg="black",
                                     cursor="hand2", font=("arcadia", 12, "bold")).winfo_reqwidth()

    # Determine the maximum width of the yellow_button widgets
    max_width = max(yellow_button1_width, yellow_button2_width)

    yellow_button1 = tk.Button(root, text=open_text, command=open_command, bg="gold", fg="black", cursor="hand2",
                               font=("arcadia", 12, "bold"))
    yellow_button1.pack()
    yellow_button1.place(x=x, y=y, width=max_width)

    yellow_button2 = tk.Button(root, text=close_text, command=close_command, bg="gold", fg="black", cursor="hand2",
                               font=("arcadia", 12, "bold"))
    yellow_button2.pack()
    yellow_button2.place(x=x + max_width + 5, y=y, width=max_width)


# Button 13b: Backup Theater Stuff
B1 = tk.Button(root, text='Halo 3 Film Backup', command=copy_recent_h3_theater_file, bg="gold", fg="black",
                 cursor="hand2", font=("arcadia", 12, "bold"))
B1.pack()
B1.place(x=10, y=10)

B2 = tk.Button(root, text='Halo 3 ODST Film Backup', command=copy_recent_h3odst_theater_file, bg="gold", fg="black",
                 cursor="hand2", font=("arcadia", 12, "bold"))
B2.pack()
B2.place(x=10, y=60)

B3 = tk.Button(root, text='Halo Reach Film Backup', command=copy_recent_haloreach_theater_file, bg="gold", fg="black",
                 cursor="hand2", font=("arcadia", 12, "bold"))
B3.pack()
B3.place(x=10, y=110)

B4 = tk.Button(root, text='Halo 4 Film Backup', command=copy_recent_h4_theater_file, bg="gold", fg="black",
                 cursor="hand2", font=("arcadia", 12, "bold"))
B4.pack()
B4.place(x=10, y=160)


root.mainloop()