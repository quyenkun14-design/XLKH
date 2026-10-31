import tkinter as tk
from gui import LoginWindow

if __name__ == "__main__":
    root = tk.Tk()
    app = LoginWindow(root)
    root.mainloop()