import ctypes
print("admin?", ctypes.windll.shell32.IsUserAnAdmin() != 0)