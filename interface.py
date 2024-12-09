from tkinter import *
from tkinter import ttk

def create_tkinter_app():
    # Create the main application window
    root = Tk()
    root.title("PwC Extractor")
    root.geometry("800x400")

    # Create a frame for padding
    frm = ttk.Frame(root, padding=10)
    frm.grid()

    # Add a label and a button to the frame
    ttk.Label(frm, text="Hello World!").grid(column=0, row=0)
    ttk.Button(frm, text="Quit", command=root.destroy).grid(column=1, row=0)

    # Start the main event loop
    root.mainloop()

if __name__ == "__main__":
    create_tkinter_app()
    pass