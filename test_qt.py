import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel

class TestWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("テストウィンドウ")
        self.setGeometry(100, 100, 400, 300)
        
        # メインウィジェットとレイアウトの設定
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()
        main_widget.setLayout(layout)
        
        # ラベルの追加
        label = QLabel("これはテストです")
        layout.addWidget(label)
        
        # ボタンの追加
        button = QPushButton("クリック")
        button.clicked.connect(lambda: print("ボタンがクリックされました"))
        layout.addWidget(button)

def main():
    app = QApplication(sys.argv)
    window = TestWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main() 