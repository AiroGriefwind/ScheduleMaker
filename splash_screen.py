import sys
import time
from PySide6.QtWidgets import QApplication, QSplashScreen, QProgressBar, QLabel, QVBoxLayout, QWidget
from PySide6.QtCore import Qt, QThread, Signal, QTimer
from PySide6.QtGui import QPixmap, QFont, QColor

class LoadingThread(QThread):
    progress_update = Signal(int, str)
    
    def __init__(self, tasks):
        super().__init__()
        self.tasks = tasks
        self.total_tasks = len(tasks)
    
    def run(self):
        # Task 1: Initialize logging
        self.progress_update.emit(10, "Initializing logging system...")
        from logger_utils import setup_logging
        setup_logging()
        
        # Task 2: Load employee data
        self.progress_update.emit(30, "Loading employee data...")
        from scheduling_logic import load_employees
        employees = load_employees()
        
        # Task 3: Load availability data
        self.progress_update.emit(50, "Loading availability data...")
        from scheduling_logic import load_data
        availability = load_data()
        
        # Task 4: Initialize UI components
        self.progress_update.emit(70, "Initializing UI components...")
        # (UI initialization code)
        
        # Task 5: Check for updates
        self.progress_update.emit(90, "Checking for updates...")
        from updater import Updater
        updater = Updater()
        updater.check_for_updates()
        
        # Finished
        self.progress_update.emit(100, "Ready!")


class SplashScreen(QSplashScreen):
    def __init__(self):
        super().__init__()
        self.setWindowFlag(Qt.WindowStaysOnTopHint)
        self.setWindowFlag(Qt.FramelessWindowHint)
        
        # Create a widget to hold our layout
        self.content = QWidget(self)
        layout = QVBoxLayout(self.content)
        
        # Add title
        self.title = QLabel("Schedule Maker")
        self.title.setAlignment(Qt.AlignCenter)
        self.title.setStyleSheet("font-size: 24px; font-weight: bold; color: #333;")
        layout.addWidget(self.title)
        
        # Add progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid grey;
                border-radius: 5px;
                text-align: center;
                height: 25px;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                width: 10px;
                margin: 0.5px;
            }
        """)
        layout.addWidget(self.progress_bar)
        
        # Add status label
        self.status_label = QLabel("Initializing...")
        self.status_label.setAlignment(Qt.AlignRight)
        self.status_label.setStyleSheet("font-size: 11px; color: #555;")
        layout.addWidget(self.status_label)
        
        # Set the content widget as the central widget
        self.setFixedSize(400, 200)
        self.content.setGeometry(0, 0, 400, 200)
        self.content.setStyleSheet("background-color: white;")
        
    def update_progress(self, value, status):
        self.progress_bar.setValue(value)
        self.status_label.setText(status)
        self.repaint()  # Force repaint to update the UI

def initialize_app(app, main_window_class):
    # Define initialization tasks
    tasks = [
        "Loading configuration...",
        "Initializing logging system...",
        "Loading employee data...",
        "Loading role rules...",
        "Loading availability data...",
        "Initializing UI components...",
        "Setting up event handlers...",
        "Checking for updates...",
        "Finalizing initialization..."
    ]
    
    # Create and show splash screen
    splash = SplashScreen()
    splash.show()
    
    # Create loading thread
    loading_thread = LoadingThread(tasks)
    loading_thread.progress_update.connect(splash.update_progress)
    
    # Create a reference to store the main window
    main_window = [None]
    
    def on_finished():
        # Create and show main window
        main_window[0] = main_window_class()
        main_window[0].show()
        splash.finish(main_window[0])
    
    loading_thread.finished.connect(on_finished)
    loading_thread.start()
    
    return app.exec(), main_window[0]
