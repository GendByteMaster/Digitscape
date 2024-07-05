import customtkinter as ctk
from DigitscapeAnalyzer import DigitscapeAnalyzer
from DigitscapeGUI import DigitscapeGUI
import matplotlib

matplotlib.use("TkAgg")

if __name__ == "__main__":
    ctk.set_appearance_mode("dark")  # Установите режим "dark" или "light"
    ctk.set_default_color_theme("blue")  # Установите тему "blue" или другую доступную тему

    root = ctk.CTk()
    analyzer = DigitscapeAnalyzer()
    app = DigitscapeGUI(root, analyzer)
    root.mainloop()
