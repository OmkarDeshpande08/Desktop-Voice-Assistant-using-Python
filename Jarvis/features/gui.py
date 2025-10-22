"""
GUI module for Jarvis Assistant
"""

from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1200, 800)
        MainWindow.setMinimumSize(QtCore.QSize(1200, 800))
        MainWindow.setMaximumSize(QtCore.QSize(1200, 800))
        MainWindow.setStyleSheet("""
            QMainWindow {
                background-color: #1e1e1e;
                color: #ffffff;
            }
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
            QTextBrowser {
                background-color: #2d2d2d;
                color: #ffffff;
                border: 1px solid #555555;
                border-radius: 5px;
                padding: 5px;
            }
            QLabel {
                color: #ffffff;
            }
        """)
        
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        
        # Main layout
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        
        # Title label
        self.title_label = QtWidgets.QLabel(self.centralwidget)
        self.title_label.setAlignment(QtCore.Qt.AlignCenter)
        self.title_label.setObjectName("title_label")
        font = QtGui.QFont()
        font.setPointSize(24)
        font.setBold(True)
        self.title_label.setFont(font)
        self.title_label.setText("JARVIS - Voice Assistant")
        self.gridLayout.addWidget(self.title_label, 0, 0, 1, 4)
        
        # Main animation label (live wallpaper)
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setMinimumSize(QtCore.QSize(400, 300))
        self.label.setMaximumSize(QtCore.QSize(400, 300))
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setObjectName("label")
        self.label.setStyleSheet("border: 2px solid #0078d4; border-radius: 10px;")
        self.label.setText("Main Animation Area")
        self.gridLayout.addWidget(self.label, 1, 0, 2, 2)
        
        # Status animation label
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setMinimumSize(QtCore.QSize(300, 200))
        self.label_2.setMaximumSize(QtCore.QSize(300, 200))
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setObjectName("label_2")
        self.label_2.setStyleSheet("border: 2px solid #0078d4; border-radius: 10px;")
        self.label_2.setText("Status Animation")
        self.gridLayout.addWidget(self.label_2, 1, 2, 1, 2)
        
        # Date display
        self.textBrowser = QtWidgets.QTextBrowser(self.centralwidget)
        self.textBrowser.setMaximumSize(QtCore.QSize(150, 50))
        self.textBrowser.setObjectName("textBrowser")
        font = QtGui.QFont()
        font.setPointSize(12)
        self.textBrowser.setFont(font)
        self.gridLayout.addWidget(self.textBrowser, 2, 2, 1, 1)
        
        # Time display
        self.textBrowser_2 = QtWidgets.QTextBrowser(self.centralwidget)
        self.textBrowser_2.setMaximumSize(QtCore.QSize(150, 50))
        self.textBrowser_2.setObjectName("textBrowser_2")
        font = QtGui.QFont()
        font.setPointSize(12)
        self.textBrowser_2.setFont(font)
        self.gridLayout.addWidget(self.textBrowser_2, 2, 3, 1, 1)
        
        # Control buttons
        self.button_frame = QtWidgets.QFrame(self.centralwidget)
        self.button_layout = QtWidgets.QHBoxLayout(self.button_frame)
        
        # Start button
        self.pushButton = QtWidgets.QPushButton(self.button_frame)
        self.pushButton.setMinimumSize(QtCore.QSize(120, 40))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.setText("Start Jarvis")
        self.button_layout.addWidget(self.pushButton)
        
        # Exit button
        self.pushButton_2 = QtWidgets.QPushButton(self.button_frame)
        self.pushButton_2.setMinimumSize(QtCore.QSize(120, 40))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.setText("Exit")
        self.pushButton_2.setStyleSheet("""
            QPushButton {
                background-color: #d13438;
            }
            QPushButton:hover {
                background-color: #b71c1c;
            }
            QPushButton:pressed {
                background-color: #8e0000;
            }
        """)
        self.button_layout.addWidget(self.pushButton_2)
        
        self.gridLayout.addWidget(self.button_frame, 3, 0, 1, 4)
        
        # Status bar
        self.status_label = QtWidgets.QLabel(self.centralwidget)
        self.status_label.setObjectName("status_label")
        self.status_label.setText("Status: Ready")
        self.status_label.setStyleSheet("color: #00ff00; font-size: 12px;")
        self.gridLayout.addWidget(self.status_label, 4, 0, 1, 4)
        
        MainWindow.setCentralwidget(self.centralwidget)
        
        # Create placeholder images for animations
        self.create_placeholder_images()
        
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def create_placeholder_images(self):
        """Create placeholder images if the actual GIF files don't exist"""
        import os
        
        # Create utils/images directory if it doesn't exist
        os.makedirs("Jarvis/utils/images", exist_ok=True)
        
        # Create placeholder files if they don't exist
        placeholder_files = [
            "Jarvis/utils/images/live_wallpaper.gif",
            "Jarvis/utils/images/initiating.gif"
        ]
        
        for file_path in placeholder_files:
            if not os.path.exists(file_path):
                # Create a simple placeholder file
                with open(file_path, 'w') as f:
                    f.write("# Placeholder for animation file")

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "JARVIS Voice Assistant"))