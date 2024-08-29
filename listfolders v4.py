import os
import csv
from tkinter import Tk, Button, Label, filedialog, messagebox, Listbox, Scrollbar, END, RIGHT, LEFT, Y, Frame, Toplevel, Entry, Radiobutton, StringVar
from tkinterdnd2 import TkinterDnD, DND_FILES
from openpyxl import Workbook

# Global constants and variables
initial_dir = os.path.expanduser("~")  # Start in the user's home directory by default
selected_folders = []
undo_stack = []
redo_stack = []

# Folder Handling Functions
def add_folder():
    global initial_dir
    folder = filedialog.askdirectory(title="Select Folder", initialdir=initial_dir)
    if folder:
        save_state_for_undo()  # Save state before modification
        selected_folders.append(folder)
        update_listbox()
        initial_dir = os.path.dirname(folder)

def update_listbox():
    listbox.delete(0, END)
    for folder in selected_folders:
        listbox.insert(END, folder)
    label.config(text=f"Selected Folders: {len(selected_folders)}")
    listbox.yview(END)  # Automatically scroll to the bottom

def on_drop(event):
    files = root.tk.splitlist(event.data)
    save_state_for_undo()  # Save state before modification
    for file in files:
        if os.path.isdir(file):  # Only add directories
            selected_folders.append(file)
    update_listbox()

def save_state_for_undo():
    undo_stack.append(list(selected_folders))
    redo_stack.clear()  # Clear the redo stack after a new action

def remove_folder():
    selected = listbox.curselection()
    if selected:
        save_state_for_undo()  # Save state before modification
        selected_folder = selected_folders[selected[0]]
        selected_folders.remove(selected_folder)
        update_listbox()
    else:
        messagebox.showwarning("No Selection", "Please select a folder to remove.")

def clear_all():
    if selected_folders:
        save_state_for_undo()  # Save state before modification
        selected_folders.clear()
        update_listbox()

def undo_action():
    if undo_stack:
        redo_stack.append(list(selected_folders))  # Save current state for redo
        last_state = undo_stack.pop()
        selected_folders.clear()
        selected_folders.extend(last_state)
        update_listbox()
    else:
        messagebox.showinfo("Undo", "No actions to undo.")

def redo_action():
    if redo_stack:
        undo_stack.append(list(selected_folders))  # Save current state for undo
        next_state = redo_stack.pop()
        selected_folders.clear()
        selected_folders.extend(next_state)
        update_listbox()
    else:
        messagebox.showinfo("Redo", "No actions to redo.")

def generate_file():
    if not selected_folders:
        messagebox.showwarning("No Folders Selected", "Please select at least one folder.")
        return

    folder_names = collect_folder_names()

    # Prompt for file type and name
    file_type = filedialog.asksaveasfilename(defaultextension=".txt",
                                             filetypes=[("Text Files", "*.txt"), ("CSV Files", "*.csv"), ("Excel Files", "*.xlsx")])

    if file_type.endswith('.txt'):
        save_as_txt(file_type, folder_names)
    elif file_type.endswith('.csv'):
        save_as_csv(file_type, folder_names)
    elif file_type.endswith('.xlsx'):
        save_as_excel(file_type, folder_names)

def save_as_txt(filename, folder_names):
    with open(filename, "w") as f:
        f.write("Folder Names\n")
        f.write(f"Number of Folders: {len(folder_names)}\n\n")
        for name in folder_names:
            f.write(f"{name}\n")
    messagebox.showinfo("Success", f"Folder names saved to {filename}")

def save_as_csv(filename, folder_names):
    with open(filename, "w", newline='') as f:
        writer = csv.writer(f)
        writer.writerow(["Folder Names"])
        writer.writerow([f"Number of Folders: {len(folder_names)}"])
        writer.writerow([])
        for name in folder_names:
            writer.writerow([name])
    messagebox.showinfo("Success", f"Folder names saved to {filename}")

def save_as_excel(filename, folder_names):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Folder Names"

    # Write the header
    sheet["A1"] = "Folder Names"
    sheet["A2"] = f"Number of Folders: {len(folder_names)}"

    # Write the folder names
    for idx, name in enumerate(folder_names, start=4):
        sheet[f"A{idx}"] = name

    workbook.save(filename)
    messagebox.showinfo("Success", f"Folder names saved to {filename}")

def collect_folder_names():
    folder_names = []
    for folder in selected_folders:
        for root_dir, dirs, files in os.walk(folder):
            folder_names.extend(dirs)
            break  # We only want the immediate subdirectories, not a recursive listing
    folder_names.sort()
    return folder_names

def search_folders():
    search_window = Toplevel(root)
    search_window.title("Folder Search")
    search_window.geometry("350x250")

    Label(search_window, text="Type folder name to search for:").pack(pady=10)
    folder_name_entry = Entry(search_window)
    folder_name_entry.pack(pady=10)

    search_depth = StringVar(value="top")

    top_level_rb = Radiobutton(search_window, text="Top Level Directories Only", variable=search_depth, value="top")
    top_level_rb.pack(anchor="w", padx=20)

    deep_search_rb = Radiobutton(search_window, text="Up to 3 Directories Down", variable=search_depth, value="deep")
    deep_search_rb.pack(anchor="w", padx=20)

    all_dirs_rb = Radiobutton(search_window, text="All Directories", variable=search_depth, value="all")
    all_dirs_rb.pack(anchor="w", padx=20)

    def perform_search():
        folder_name = folder_name_entry.get().strip()
        save_state_for_undo()  # Save state before modification
        if folder_name.startswith('"') and folder_name.endswith('"'):
            folder_name = folder_name[1:-1]  # Remove the quotes
            match_function = lambda name: name == folder_name.lower()
        else:
            match_function = lambda name: folder_name.lower() in name

        if folder_name:
            drives = [f"{d}:\\" for d in "ABCDEFGHIJKLMNOPQRSTUVWXYZ" if os.path.exists(f"{d}:\\")]
            for drive in drives:
                if search_depth.get() == "top":
                    for folder in os.listdir(drive):
                        folder_path = os.path.join(drive, folder)
                        if os.path.isdir(folder_path) and match_function(folder.lower()):
                            selected_folders.append(folder_path)
                elif search_depth.get() == "deep":
                    for root, dirs, _ in os.walk(drive):
                        depth = root[len(drive):].count(os.sep)
                        if depth <= 3:
                            for dir_name in dirs:
                                if match_function(dir_name.lower()):
                                    selected_folders.append(os.path.join(root, dir_name))
                else:  # All directories
                    for root, dirs, _ in os.walk(drive):
                        for dir_name in dirs:
                            if match_function(dir_name.lower()):
                                selected_folders.append(os.path.join(root, dir_name))
            update_listbox()
            search_window.destroy()

    search_button = Button(search_window, text="Search and Add", command=perform_search)
    search_button.pack(pady=20)

# UI Setup
root = TkinterDnD.Tk()  # Use TkinterDnD for drag-and-drop functionality
root.title("Select Folders")
root.geometry("900x600")  # Increase window size by 50%

label = Label(root, text="Selected Folders: 0")
label.pack(pady=10)

frame = Frame(root)
frame.pack(fill='both', expand=True, padx=10, pady=10)

scrollbar = Scrollbar(frame)
scrollbar.pack(side=RIGHT, fill=Y)

listbox = Listbox(frame)
listbox.pack(side=RIGHT, fill='both', expand=True)
scrollbar.config(command=listbox.yview)

listbox.drop_target_register(DND_FILES)
listbox.dnd_bind('<<Drop>>', on_drop)

info_label = Label(root, text="Drag and drop folders into window above, or use the buttons below.")
info_label.pack(pady=10)

button_row_frame = Frame(root)
button_row_frame.pack(pady=10, fill='x')

left_button_frame = Frame(button_row_frame)
left_button_frame.pack(side=LEFT)

add_button = Button(left_button_frame, text="Add Folder", command=add_folder)
add_button.pack(side=LEFT, padx=10)

remove_button = Button(left_button_frame, text="Remove Folder", command=remove_folder)
remove_button.pack(side=LEFT, padx=10)

search_button = Button(left_button_frame, text="Search", command=search_folders)
search_button.pack(side=LEFT, padx=10)

center_button_frame = Frame(button_row_frame)
center_button_frame.pack(side=LEFT, expand=True)

generate_button = Button(center_button_frame, text="Generate File", command=generate_file)
generate_button.pack(side=LEFT, padx=10)

right_button_frame = Frame(button_row_frame)
right_button_frame.pack(side=RIGHT)

undo_button = Button(right_button_frame, text="Undo", command=undo_action)
undo_button.pack(side=LEFT, padx=10)

redo_button = Button(right_button_frame, text="Redo", command=redo_action)
redo_button.pack(side=LEFT, padx=10)

clear_button = Button(right_button_frame, text="Clear All", command=clear_all)
clear_button.pack(side=LEFT, padx=10)

# Start the main loop
root.mainloop()
