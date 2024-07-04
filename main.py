import ttkbootstrap as ttk
from DigitscapeAnalyzer import DigitscapeAnalyzer
from DigitscapeGUI import DigitscapeGUI

if __name__ == "__main__":
    try:
        root = ttk.Window(title="Digitscape Analyzer", size=(800, 600), resizable=(False, False))
        analyzer = DigitscapeAnalyzer()
        gui = DigitscapeGUI(root, analyzer)
        root.mainloop()
    except Exception as e:
        print(f"Произошла ошибка: {e}")