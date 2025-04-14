import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QVBoxLayout, QWidget

class SimpleWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Simple Test')
        self.setGeometry(100, 100, 400, 200)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout()
        central_widget.setLayout(layout)
        
        label = QLabel('テストウィンドウ')
        layout.addWidget(label)
        
        button = QPushButton('クリック')
        button.clicked.connect(lambda: label.setText('ボタンがクリックされました！'))
        layout.addWidget(button)

if __name__ == '__main__':
    import os
    # macOS向けの設定
    os.environ['QT_MAC_WANTS_LAYER'] = '1'
    
    # シンプルに起動
    app = QApplication(sys.argv)
    window = SimpleWindow()
    window.show()
    sys.exit(app.exec_()) 