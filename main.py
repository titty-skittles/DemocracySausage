import tkinter as tk

from gui import ElectionApp


def main() -> None:
    root = tk.Tk()
    ElectionApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()