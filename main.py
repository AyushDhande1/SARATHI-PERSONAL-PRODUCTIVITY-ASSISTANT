import time
import csv
import os
import psutil
import threading
import pandas as pd
import win32gui
import win32process
from calendar_utils import add_event
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.auth.exceptions import GoogleAuthError
from datetime import datetime, timedelta
from calendar_utils import schedule_goal_in_calendar, get_calendar_service
import pytz
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QMessageBox,
    QTableWidget, QTableWidgetItem, QLineEdit, QHBoxLayout, QListWidget
)
import sys
import json
from sentence_transformers import SentenceTransformer, util

# File paths
LOG_FILE = "screen_activity_log.csv"
GOALS_FILE = "user_goals.json"
ANALYZED_LOG_FILE = "analyzed_screen_activity_log.csv"

# Activity tracking globals
current_window = None
current_app = None
current_url = None
start_timestamp = None
activity_data = []
tracking = False

# Load NLP model
model = SentenceTransformer("all-MiniLM-L6-v2")
SIMILARITY_THRESHOLD = 0.25

def init_csv():
    if not os.path.exists(LOG_FILE):
        headers = ["Start Timestamp", "End Timestamp", "Application", "Window Title", "URL", "Time Spent (s)"]
        with open(LOG_FILE, mode='w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, quoting=csv.QUOTE_ALL)
            writer.writerow(headers)

def get_active_process_name():
    try:
        hwnd = win32gui.GetForegroundWindow()
        if hwnd:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            process = psutil.Process(pid)
            return process.name()
        return "Unknown"
    except Exception:
        return "Unknown"

def get_active_window_title():
    try:
        hwnd = win32gui.GetForegroundWindow()
        if hwnd:
            return win32gui.GetWindowText(hwnd)
        return "Unknown"
    except Exception:
        return "Unknown"

def get_calendar_service():
    try:
        creds = Credentials.from_authorized_user_file('token.json', ['https://www.googleapis.com/auth/calendar'])
        service = build('calendar', 'v3', credentials=creds)
        return service
    except GoogleAuthError as e:
        print(f"Authentication failed: {e}")
        return None

def fetch_events(start_time, end_time):
    service = get_calendar_service()
    if not service:
        return []

    try:
        events_result = service.events().list(
            calendarId='primary',
            timeMin=start_time.isoformat(),
            timeMax=end_time.isoformat(),
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        return events_result.get('items', [])
    except Exception as e:
        print(f"Error fetching events: {e}")
        return []

def find_free_slots(events, start_time, end_time, goal_duration):
    free_slots = []
    current_time = start_time

    for event in events:
        event_start = datetime.fromisoformat(event['start']['dateTime'])
        event_end = datetime.fromisoformat(event['end']['dateTime'])

        if (event_start - current_time).total_seconds() >= goal_duration:
            free_slots.append((current_time, event_start))

        current_time = max(current_time, event_end)

    if (end_time - current_time).total_seconds() >= goal_duration:
        free_slots.append((current_time, end_time))

    return free_slots

def add_event_to_calendar(event_name, start_time, end_time):
    service = get_calendar_service()
    if not service:
        return False

    event = {
        'summary': event_name,
        'start': {'dateTime': start_time.isoformat(), 'timeZone': 'Asia/Kolkata'},
        'end': {'dateTime': end_time.isoformat(), 'timeZone': 'Asia/Kolkata'},
    }

    try:
        service.events().insert(calendarId='primary', body=event).execute()
        print(f"Event '{event_name}' scheduled successfully.")
        return True
    except Exception as e:
        print(f"Error scheduling event: {e}")
        return False


def track_activity():
    global current_window, current_app, current_url, start_timestamp, activity_data, tracking
    app_mapping = {"Code.exe": "VS Code", "chrome.exe": "Google Chrome", "msedge.exe": "Microsoft Edge"}

    while tracking:
        try:
            window_title = get_active_window_title()
            exe_name = get_active_process_name()
            app_name = app_mapping.get(exe_name, exe_name)
            url = "N/A"
            end_timestamp = datetime.now()

            if window_title != current_window or app_name != current_app:
                if current_window and current_app and start_timestamp:
                    time_spent = (end_timestamp - start_timestamp).total_seconds()
                    activity_data.append([
                        start_timestamp.strftime("%Y-%m-%d %H:%M:%S"),
                        end_timestamp.strftime("%Y-%m-%d %H:%M:%S"),
                        current_app, current_window, current_url, round(time_spent, 2)
                    ])
                current_window, current_app, current_url, start_timestamp = window_title, app_name, url, end_timestamp

            time.sleep(1)
        except Exception as e:
            print(f"Error tracking activity: {e}")
            time.sleep(1)

def save_logs():
    global activity_data
    while tracking:
        if activity_data:
            try:
                df = pd.DataFrame(activity_data, columns=["Start Timestamp", "End Timestamp", "Application", "Window Title", "URL", "Time Spent (s)"])
                df.to_csv(LOG_FILE, mode='a', header=False, index=False, quoting=csv.QUOTE_ALL)
                activity_data = []
            except Exception as e:
                print(f"Error writing to CSV: {e}")
        time.sleep(1)

def run_semantic_matching():
    if not os.path.exists(LOG_FILE) or not os.path.exists(GOALS_FILE):
        return

    log_df = pd.read_csv(LOG_FILE)
    with open(GOALS_FILE, "r", encoding='utf-8') as f:
        goals_data = json.load(f)

    log_df["Matched Goal"] = ""
    log_df["Similarity Score"] = 0.0

    for index, row in log_df.iterrows():
        activity_text = f"{row['Application']} {row['Window Title']}"
        activity_embedding = model.encode(activity_text, convert_to_tensor=True)

        max_score = 0
        matched_goal = "No Match"

        for goal_entry in goals_data:
            goal = goal_entry["goal"]
            keywords = goal_entry["keywords"]
            goal_text = goal + " " + " ".join(keywords)
            goal_embedding = model.encode(goal_text, convert_to_tensor=True)
            similarity = util.cos_sim(activity_embedding, goal_embedding).item()

            if similarity > max_score:
                max_score = similarity
                matched_goal = goal

        log_df.at[index, "Matched Goal"] = matched_goal if max_score >= SIMILARITY_THRESHOLD else "No Match"
        log_df.at[index, "Similarity Score"] = round(max_score, 3)

    log_df["User Corrected Goal"] = log_df["Matched Goal"]
    log_df.to_csv(ANALYZED_LOG_FILE, index=False)
    print("Semantic matching complete.")
    
    latest_score = calculate_daily_reward_scores_with_streak(log_df)
    window.reward_label.setText(f"Daily Reward: {round(latest_score)} points")
    schedule_goal_events(log_df)
    
def schedule_goal_events(log_df):
    from datetime import datetime

    service = get_calendar_service()

    for index, row in log_df.iterrows():
        if row["Matched Goal"] != "No Match":
            try:
                start = datetime.strptime(row["Start Timestamp"], "%Y-%m-%d %H:%M:%S")
                end = datetime.strptime(row["End Timestamp"], "%Y-%m-%d %H:%M:%S")
                duration_seconds = int((end - start).total_seconds())
                hours = duration_seconds // 3600
                minutes = (duration_seconds % 3600) // 60
                schedule_goal_in_calendar(row["Matched Goal"], hours, minutes, service)
            except Exception as e:
                print(f"Calendar scheduling failed for row {index}: {e}")

                
def save_goals_to_file(goals):
    with open(GOALS_FILE, "w", encoding="utf-8") as f:
        json.dump(goals, f, indent=4)

def load_goals_from_file():
    if os.path.exists(GOALS_FILE):
        with open(GOALS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def calculate_daily_reward_scores_with_streak(log_df):
    reward_scores = {}
    streak = {}
    current_streak = 0

    for index, row in log_df.iterrows():
        if row["Matched Goal"] != "No Match":
            date = row["Start Timestamp"].split(" ")[0]
            time_spent = row["Time Spent (s)"]
            
            # Convert time spent into points (e.g., 1 point per 60 seconds)
            points = time_spent / 60  # Points per minute (adjust the ratio as needed)
            
            # Add points to the daily reward score
            reward_scores[date] = reward_scores.get(date, 0) + points
            
            # Track streaks
            if date not in streak:
                # If this is the first time we see the date, it's the start of a streak
                streak[date] = current_streak + 1
                current_streak += 1
            else:
                # If the user has already earned points on this day, increase the streak
                streak[date] = streak.get(date, 0)

    # Assuming you want to return the latest day's reward score to update the GUI
    latest_date = max(reward_scores.keys())
    latest_score = reward_scores[latest_date]

    # Return the latest reward score
    return latest_score




class ScreenMonitorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Smart Screen Monitor")
        self.setGeometry(200, 200, 500, 500)

        layout = QVBoxLayout()
        self.label = QLabel("Status: Monitoring Not Started")
        layout.addWidget(self.label)

        self.goal_input = QLineEdit()
        self.goal_input.setPlaceholderText("Enter study goal")
        layout.addWidget(self.goal_input)

        self.keyword_input = QLineEdit()
        self.keyword_input.setPlaceholderText("Enter keywords (comma separated)")
        layout.addWidget(self.keyword_input)

        self.duration_hours_input = QLineEdit()
        self.duration_hours_input.setPlaceholderText("Duration Hours")
        layout.addWidget(self.duration_hours_input)

        self.duration_minutes_input = QLineEdit()
        self.duration_minutes_input.setPlaceholderText("Duration Minutes")
        layout.addWidget(self.duration_minutes_input)

        btn_layout = QHBoxLayout()
        self.save_button = QPushButton("Save Goal")
        self.save_button.clicked.connect(self.save_goal)
        btn_layout.addWidget(self.save_button)

        self.delete_button = QPushButton("Delete Goal")
        self.delete_button.clicked.connect(self.delete_goal)
        btn_layout.addWidget(self.delete_button)
        layout.addLayout(btn_layout)

        self.goal_list = QListWidget()
        layout.addWidget(self.goal_list)
        self.load_goals()
     
        self.reward_label = QLabel("Daily Reward: 0 points")
        layout.addWidget(self.reward_label)        
        
        self.streak_label = QLabel("Current Streak: 0 days")
        layout.addWidget(self.streak_label)
        
        self.start_button = QPushButton("Start Monitoring")
        self.start_button.clicked.connect(self.start_tracking)
        layout.addWidget(self.start_button)

        self.stop_button = QPushButton("Stop Monitoring")
        self.stop_button.clicked.connect(self.stop_tracking)
        layout.addWidget(self.stop_button)

        self.view_button = QPushButton("View Logs")
        self.view_button.clicked.connect(self.view_logs)
        layout.addWidget(self.view_button)

        self.view_analyzed_button = QPushButton("View/Edit Analyzed Data")
        self.view_analyzed_button.clicked.connect(self.view_analyzed_logs)
        layout.addWidget(self.view_analyzed_button)

        self.setLayout(layout)
        self.goals = load_goals_from_file()

    def update_reward(self, points):
        self.reward_label.setText(f"Daily Reward: {points} points")

    def update_streak(self, streak_days):
        self.streak_label.setText(f"Current Streak: {streak_days} days")

    def increase_reward_and_streak(self):
        self.reward_points += 10  # Increase by 10 points for each goal-aligned task
        self.update_reward(self.reward_points)
        
        # Increase streak if user is actively monitoring (example condition)
        if self.is_monitoring_active():  # You can define this check based on your tracking logic
            self.streak += 1
            self.update_streak(self.streak)
        else:
            self.streak = 0  # Reset streak if monitoring stops or gap is detected
            self.update_streak(self.streak)
            
    def save_goal(self):
        goal = self.goal_input.text().strip()
        keywords = [kw.strip() for kw in self.keyword_input.text().strip().split(",") if kw.strip()]
    
    # Get duration from user input
        try:
            hours = int(self.duration_hours_input.text())
        except ValueError:
            hours = 0
        try:
            minutes = int(self.duration_minutes_input.text())
        except ValueError:
            minutes = 0

        if goal and keywords:
            self.goals.append({"goal": goal, "keywords": keywords})
            self.goal_list.addItem(f"{goal} - {keywords}")
            save_goals_to_file(self.goals)
            self.goal_input.clear()
            self.keyword_input.clear()
            self.duration_hours_input.clear()
            self.duration_minutes_input.clear()

            try:
                calendar_service = get_calendar_service()
                schedule_goal_in_calendar(goal, hours, minutes, calendar_service)
                print(f"[INFO] Scheduled '{goal}' for {hours}h {minutes}m in calendar.")
            except Exception as e:
                print(f"[ERROR] Could not schedule goal: {e}")
                QMessageBox.warning(self, "Calendar Error", f"Could not schedule in calendar.\n{e}")

            QMessageBox.information(self, "Saved", f"Goal '{goal}' saved and scheduled.")
        else:
            QMessageBox.warning(self, "Invalid", "Please enter goal, keywords, and duration.")



    def delete_goal(self):
        
        
        selected_items = self.goal_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "No Selection", "Please select a goal to delete.")
            return
        for item in selected_items:
            index = self.goal_list.row(item)
            self.goal_list.takeItem(index)
            del self.goals[index]
        save_goals_to_file(self.goals)
        QMessageBox.information(self, "Deleted", "Goal(s) deleted.")

    def load_goals(self):
        self.goal_list.clear()
        for goal in load_goals_from_file():
            self.goal_list.addItem(f"{goal['goal']} - {goal['keywords']}")

    def start_tracking(self):
        global tracking
        if not tracking:
            tracking = True
            threading.Thread(target=track_activity, daemon=True).start()
            threading.Thread(target=save_logs, daemon=True).start()
            self.label.setText("Status: Monitoring Active")
            QMessageBox.information(self, "Started", "Screen monitoring started.")

    def stop_tracking(self):
        global tracking
        if tracking:
            tracking = False
            self.label.setText("Status: Monitoring Stopped")
            run_semantic_matching()
            QMessageBox.information(self, "Stopped", "Monitoring stopped.\nSemantic analysis complete.")

    def view_logs(self):
        self.log_viewer = LogViewer()
        self.log_viewer.show()

    def view_analyzed_logs(self):
        if not os.path.exists(ANALYZED_LOG_FILE):
            QMessageBox.warning(self, "Not Found", "Analyzed logs not found. Please stop monitoring first.")
            return
        self.analyzed_log_viewer = AnalyzedLogEditor()
        self.analyzed_log_viewer.show()

class LogViewer(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Activity Log Viewer")
        self.setGeometry(250, 250, 800, 400)
        layout = QVBoxLayout()
        self.table = QTableWidget()
        layout.addWidget(self.table)
        self.setLayout(layout)
        self.load_logs()

    def load_logs(self):
        if not os.path.exists(LOG_FILE):
            QMessageBox.critical(self, "Error", "Log file not found!")
            return
        df = pd.read_csv(LOG_FILE)
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns)
        for row_idx, row_data in df.iterrows():
            for col_idx, col_data in enumerate(row_data):
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(col_data)))

class AnalyzedLogEditor(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Edit Analyzed Log")
        self.setGeometry(300, 300, 1000, 500)
        layout = QVBoxLayout()
        self.table = QTableWidget()
        layout.addWidget(self.table)

        self.save_button = QPushButton("Save Corrections")
        self.save_button.clicked.connect(self.save_corrections)
        layout.addWidget(self.save_button)

        self.setLayout(layout)
        self.load_logs()

    def load_logs(self):
        if not os.path.exists(ANALYZED_LOG_FILE):
            QMessageBox.critical(self, "Error", "Analyzed log file not found!")
            return
        self.df = pd.read_csv(ANALYZED_LOG_FILE)
        self.table.setRowCount(len(self.df))
        self.table.setColumnCount(len(self.df.columns))
        self.table.setHorizontalHeaderLabels(self.df.columns)

        for row_idx, row_data in self.df.iterrows():
            for col_idx, col_data in enumerate(row_data):
                item = QTableWidgetItem(str(col_data))
                if self.df.columns[col_idx] == "User Corrected Goal":
                    item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
                else:
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row_idx, col_idx, item)

    def save_corrections(self):
        for row in range(self.table.rowCount()):
            self.df.at[row, "User Corrected Goal"] = self.table.item(row, self.df.columns.get_loc("User Corrected Goal")).text()
        self.df.to_csv(ANALYZED_LOG_FILE, index=False)
        QMessageBox.information(self, "Saved", "Corrections saved successfully!")

if __name__ == "__main__":
    init_csv()
    app = QApplication(sys.argv)
    window = ScreenMonitorApp()
    window.show()
    sys.exit(app.exec())