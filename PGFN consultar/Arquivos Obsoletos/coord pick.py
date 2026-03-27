import pyautogui
import time

print("Mova o mouse sobre o checkbox e aguarde 3 segundos...")
time.sleep(3)
x, y = pyautogui.position()
print(f"Coordenadas do mouse: x={x}, y={y}")