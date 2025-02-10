import tkinter as tk
from tkinter import ttk, simpledialog, messagebox
from datetime import datetime
import webbrowser
import threading
import time
import os  # For shutting down the system
from pynput.keyboard import Key, Controller

# Meeting schedule storage
TimeTable = []  # Start with an empty timetable

# Automation control flag
is_running = False
keyboard = Controller()


# Function to generate a Zoom URL from a Meeting ID and password
def generate_zoom_url(meeting_id, password=None):
    meeting_id = meeting_id.replace(" ", "")  # Remove spaces from Meeting ID
    url = f"https://us04web.zoom.us/j/{meeting_id}"
    if password:
        url += f"?pwd={password}"
    return url


# Function to check and handle meetings
def check_meetings():
    global is_running
    while is_running:
        now = datetime.now()
        current_hour, current_minute = now.hour, now.minute

        for meeting in TimeTable[:]:  # Iterate over a copy to allow safe deletion
            try:
                url_or_code, start_time, end_time, *password = meeting
                password = password[0] if password else None
                start_hour, start_minute = map(int, start_time.split(":"))
                end_hour, end_minute = map(int, end_time.split(":"))

                # Start meeting
                if current_hour == start_hour and current_minute == start_minute:
                    print(f"Joining meeting: {url_or_code}")
                    if "zoom.us" in url_or_code:  # If it's a URL
                        webbrowser.open(url_or_code)
                    else:  # If it's a Meeting ID
                        meeting_url = generate_zoom_url(url_or_code, password)
                        webbrowser.open(meeting_url)

                # End meeting
                if current_hour == end_hour and current_minute == end_minute:
                    print("Ending meeting...")
                    keyboard.press("x")
                    time.sleep(1)
                    keyboard.press(Key.enter)
                    time.sleep(1)

                    # Remove the completed meeting
                    TimeTable.remove(meeting)
                    refresh_table()
            except Exception as e:
                print(f"Error processing meeting: {meeting}. Details: {e}")

        # If no meetings remain, stop automation and shut down
        if not TimeTable:
            print("All meetings completed. Stopping automation...")
            stop_automation()
            shutdown_computer()
            break

        time.sleep(30)  # Check every 30 seconds


# Start the automation
def start_automation():
    global is_running
    if not is_running:
        is_running = True
        threading.Thread(target=check_meetings, daemon=True).start()
        print("Automation started.")


# Stop the automation
def stop_automation():
    global is_running
    is_running = False
    print("Automation stopped.")


# Shutdown the computer
def shutdown_computer():
    print("Shutting down the computer...")
    os.system("shutdown /s /t 1")  # Windows command to shut down the system


# GUI functionality
def add_entry():
    url_or_code = simpledialog.askstring("Add Entry", "Enter Zoom URL or Meeting ID:")
    start_time = simpledialog.askstring("Add Entry", "Enter Start Time (HH:MM):")
    end_time = simpledialog.askstring("Add Entry", "Enter End Time (HH:MM):")
    password = simpledialog.askstring("Add Entry", "Enter Meeting Password (optional):")

    if url_or_code and start_time and end_time:
        try:
            entry = [url_or_code, start_time, end_time]
            if password:
                entry.append(password)
            TimeTable.append(entry)
            refresh_table()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add entry. Details: {e}")
    else:
        messagebox.showerror("Error", "URL/ID, Start Time, and End Time are required!")


def edit_entry():
    selected_item = listbox.selection()
    if not selected_item:
        messagebox.showwarning("No Selection", "Please select an entry to edit.")
        return

    # Get the index of the selected item
    item = selected_item[0]
    values = listbox.item(item, "values")
    index = int(values[0]) - 1  # Adjust for zero-based indexing

    # Get the current values
    current_entry = TimeTable[index]
    url_or_code = simpledialog.askstring("Edit Entry", "Enter Zoom URL or Meeting ID:", initialvalue=current_entry[0])
    start_time = simpledialog.askstring("Edit Entry", "Enter Start Time (HH:MM):", initialvalue=current_entry[1])
    end_time = simpledialog.askstring("Edit Entry", "Enter End Time (HH:MM):", initialvalue=current_entry[2])
    password = simpledialog.askstring("Edit Entry", "Enter Meeting Password (optional):",
                                       initialvalue=current_entry[3] if len(current_entry) > 3 else "")

    if url_or_code and start_time and end_time:
        try:
            # Update the entry
            new_entry = [url_or_code, start_time, end_time]
            if password:
                new_entry.append(password)
            TimeTable[index] = new_entry  # Replace the old entry
            refresh_table()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to edit entry. Details: {e}")
    else:
        messagebox.showerror("Error", "URL/ID, Start Time, and End Time are required!")


def delete_entry():
    selected_item = listbox.selection()
    if not selected_item:
        messagebox.showwarning("No Selection", "Please select an entry to delete.")
        return

    # Get the index of the selected item
    item = selected_item[0]
    values = listbox.item(item, "values")
    index = int(values[0]) - 1  # Adjust for zero-based indexing
    del TimeTable[index]
    refresh_table()


def refresh_table():
    for i in listbox.get_children():
        listbox.delete(i)
    for idx, entry in enumerate(TimeTable):
        display_entry = (idx + 1, entry[0], entry[1], entry[2], entry[3] if len(entry) > 3 else "N/A")
        listbox.insert("", "end", values=display_entry)


# GUI setup
app = tk.Tk()
app.title("Zoom Scheduler")

# Set window size
app.geometry("800x400")

frame = ttk.Frame(app, padding="10")
frame.pack(fill="both", expand=True)

# Table for displaying the timetable
columns = ("#", "URL/ID", "Start Time", "End Time", "Password")
listbox = ttk.Treeview(frame, columns=columns, show="headings", height=10)
for col in columns:
    listbox.heading(col, text=col)
    listbox.column(col, width=150, anchor="center")
listbox.pack(fill="both", expand=True)

# Buttons for controls
btn_frame = ttk.Frame(app, padding="10")
btn_frame.pack()

btn_add = ttk.Button(btn_frame, text="Add Entry", command=add_entry)
btn_add.grid(row=0, column=0, padx=10)

btn_edit = ttk.Button(btn_frame, text="Edit Entry", command=edit_entry)
btn_edit.grid(row=0, column=1, padx=10)

btn_delete = ttk.Button(btn_frame, text="Delete Entry", command=delete_entry)
btn_delete.grid(row=0, column=2, padx=10)

btn_start = ttk.Button(btn_frame, text="Start Automation", command=start_automation)
btn_start.grid(row=0, column=3, padx=10)

btn_stop = ttk.Button(btn_frame, text="Stop Automation", command=stop_automation)
btn_stop.grid(row=0, column=4, padx=10)

refresh_table()

# Stop automation on GUI close
def on_close():
    stop_automation()
    app.destroy()


app.protocol("WM_DELETE_WINDOW", on_close)
app.mainloop()
