import win32gui
import win32process
import time
import datetime
import csv
import atexit
import psutil
import os

# Constants
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_FILENAME = os.path.join(SCRIPT_DIR, 'temps_passé_FH.csv')
CSV_DATETIME_FORMAT = '%Y-%m-%d %H:%M:%S'
SLEEP_INTERVAL = 2

# Initialize the window tracking list
window_list = []

# Variable to track the last time the file was saved
last_save_time = datetime.datetime.now()

def read_csv():
    try:
        with open(CSV_FILENAME, 'r', newline='') as csvfile:
            csvreader = csv.reader(csvfile)
            next(csvreader)  # skip the header
            for row in csvreader:
                window_list.append([datetime.datetime.strptime(row[0], CSV_DATETIME_FORMAT), row[1], row[2], float(row[3]), row[4]])
    except FileNotFoundError:
        pass

def save_window_times():
    with open(CSV_FILENAME, 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(['Date', 'Windows Infos', 'Fenêtre', 'Temps (minutes)', 'Commentaires'])
        for row in window_list:
            csvwriter.writerow([row[0].strftime(CSV_DATETIME_FORMAT), row[1], row[2], row[3], row[4]])

def get_process_name(handle):
    _, pid = win32process.GetWindowThreadProcessId(handle)
    if not isinstance(pid, int) or pid <= 0:
        return "Invalid PID"
    try:
        process = psutil.Process(pid)
        return process.name()
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        return "N/A"

def main():
    global last_save_time
    atexit.register(save_window_times)
    read_csv()

    active_window = None

    while True:
        handle = win32gui.GetForegroundWindow()
        window_title = win32gui.GetWindowText(handle)
        process_name = get_process_name(handle)
        window_info = f"{process_name} - {window_title}"

        current_time = datetime.datetime.now()

        if window_info != active_window:
            if active_window is not None:
                time_difference = (current_time - last_active_time).total_seconds() / 60
                # Finding the last entry for the active window
                last_entries = [row for row in window_list if row[1] == active_window]
                if last_entries:
                    last_entry = last_entries[-1]
                    last_entry[3] += time_difference

            window_list.append([current_time, window_info, window_title, 0, ""])

            active_window = window_info
            last_active_time = current_time

        # Check if a minute has passed since the last save
        if (current_time - last_save_time).total_seconds() >= 60:
            save_window_times()
            last_save_time = current_time

        time.sleep(SLEEP_INTERVAL)

if __name__ == '__main__':
    main()
