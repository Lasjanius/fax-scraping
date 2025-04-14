import tkinter as tk
from tkinter import ttk

def main():
    root = tk.Tk()
    root.title("テストウィンドウ")
    root.geometry("400x300")
    
    label = ttk.Label(root, text="これはテストです")
    label.pack(pady=20)
    
    button = ttk.Button(root, text="クリック", command=lambda: print("ボタンがクリックされました"))
    button.pack(pady=20)
    
    root.mainloop()

if __name__ == "__main__":
    main() 