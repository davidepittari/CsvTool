from tkinter import Tk
from gui_utils import create_interface

def main():
    root = Tk()
    root.title("CSV to Excel Application")
    create_interface(root)
    root.mainloop()

if __name__ == "__main__":
    main()
