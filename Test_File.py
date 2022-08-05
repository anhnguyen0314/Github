import PyQt5.QtWidgets as qtwidget
import PyQt5.QtCore as qtcore
import time

app = qtwidget.QApplication([])

class MainWindow(qtwidget.QWidget):
    def __init__(self):
        super().__init__()
        
        # Set window title
        self.setWindowTitle('Python')
        
        height = 100
        width = 500
        self.status = "stop"
        
        # Set fixed window size
        self.setFixedHeight(height)
        self.setFixedWidth(width)
        self.display = qtwidget.QLabel("Label")
        self.display.setStyleSheet("background-color: #e3e1da;\
                                    border: 1px solid black;\
                                    padding-left: 5px")
        
        self.btn1 = qtwidget.QPushButton("Button", self)
        self.btn1.clicked.connect(self.button_action)
        
        # Set progam main layout 
        main_layout = qtwidget.QVBoxLayout()
        
        # Create horizontal box for buttons
        sub_layout = qtwidget.QHBoxLayout()
        
        # Add buttons to horizontal box
        sub_layout.addWidget(self.btn1)
        
        # Add horizontal layout to vertical box layout
        main_layout.addLayout(sub_layout)
        main_layout.addWidget(self.display)
        
        
        self.setLayout(main_layout)
        self.show()

    def button_action(self):
        self.display.setText("First")
        qtcore.QTimer.singleShot(100, lambda: self.test_function())
        
    def test_function(self):
        self.display.setText("Second")
        
mw = MainWindow()

app.exec_()
