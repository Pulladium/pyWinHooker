import asyncio
import threading
import pygetwindow as gw
import pyautogui
import pyperclip
import time
import win32api
import win32con
import pyWinhook
import queue
import pythoncom
import pyautogui
import pynput
from pynput import mouse

from functools import partial

from pywinauto import Application
# Инициализация переменных позиции мыши
from win32gui import PumpMessages

MIN_MOUSE_DRAG_DISTANCE = 10
async def print_selected_text(signal_queue_got):
    """Asynchronously print the selected text."""
    last_text = None
    selection_time = 0
    printed = False
    while not printed:
        signal = await signal_queue_got.get()
        if signal == "text_selected":
            # Получить активное окно
            active_window = gw.getActiveWindow()
            if active_window is not None:
                try:
                    # Получить приложение, связанное с активным окном
                    app = Application().connect(handle=active_window._hWnd)
                    # Попытаться получить текстовое поле активного окна
                    text_field = app.window(handle=active_window._hWnd).children(control_type="Edit")[0]
                    # Получить выделенный текст
                    selected_text = text_field.get_selection_indices()
                    print(f"Selected text: {selected_text}")
                except IndexError:
                    print("Text field not found in the active window.")
                except Exception as e:
                    print(f"An error occurred: {e}")
        await asyncio.sleep(1)

async def main():
    """Main asynchronous function."""
    signal_queue = asyncio.Queue()
    mouse_tracker = MouseTracker(signal_queue)
    asyncio.create_task(print_selected_text(signal_queue))
    asyncio.create_task(mouse_tracker.track_mouse())
    # Start the asynchronous task to print selected text
    # asyncio.create_task(print_selected_text(signal_queue))
    # await asyncio.sleep(1)
    # Keep the main function running
    while True:
        await asyncio.sleep(1)
class MouseTracker:
    def __init__(self, signal_queue):
        self.signal_queue = signal_queue
        self.start_mouse_pos = None
        self.mouse_pressed = False

    def on_click(self, x, y, button, pressed):
        if button == mouse.Button.left:
            self.mouse_pressed = pressed

    async def track_mouse(self):
        # Listen to mouse events in a separate thread
        mouse_listener = mouse.Listener(on_click=self.on_click)
        mouse_listener.start()

        while True:
            if self.mouse_pressed:
                if self.start_mouse_pos is None:
                    self.start_mouse_pos = pyautogui.position()
            elif self.start_mouse_pos:
                end_mouse_pos = pyautogui.position()
                distance = ((end_mouse_pos[0] - self.start_mouse_pos[0]) ** 2 + (end_mouse_pos[1] - self.start_mouse_pos[1]) ** 2) ** 0.5
                if distance > MIN_MOUSE_DRAG_DISTANCE:
                    self.signal_queue.put_nowait('text_selected')
                    print(f"Distance: {distance:.2f}")
                self.start_mouse_pos = None
            await asyncio.sleep(0.1)

    # def __init__(self):
    #     self.signal_queue = asyncio.Queue()
    #     self.start_mouse_pos = None
    #     self.end_mouse_pos = None
    #
    # def on_mouse_left_down(self, event):
    #     try:
    #         print("Mouse left button pressed")
    #         # Запомнить начальную позицию мыши
    #         self.start_mouse_pos = event.Position
    #         return True
    #     except Exception as e:
    #         print(f"Exception in on_mouse_left_down: {e}")
    #         return False
    #
    # def on_mouse_left_up(self, event):
    #     try:
    #         print("Mouse left button released")
    #         # Запомнить конечную позицию мыши
    #         self.end_mouse_pos = event.Position
    #         # Проверить, что позиции мыши были захвачены
    #         if self.start_mouse_pos and self.end_mouse_pos:
    #             # Вычислить расстояние между начальной и конечной позициями
    #             distance = ((self.end_mouse_pos[0] - self.start_mouse_pos[0]) ** 2 + (self.end_mouse_pos[1] - self.start_mouse_pos[1]) ** 2) ** 0.5
    #             print(f"Distance: {distance:.2f}")
    #             if distance > MIN_MOUSE_DRAG_DISTANCE:
    #                 self.signal_queue.put_nowait('text_selected')
    #                 print("Text selected")
    #                 return True
    #         else:
    #             print("Mouse positions not captured properly")
    #             return False
    #     except Exception as e:
    #         print(f"Exception in on_mouse_left_up: {e}")
    #         return False
    #
    # def start_mouse_tracking(self):
    #     hook_manager = pyWinhook.HookManager()
    #     hook_manager.HookMouse()
    #     try:
    #         if callable(self.on_mouse_left_down) and callable(self.on_mouse_left_up):
    #             # Subscribe to left mouse down and up events
    #             hook_manager.MouseLeftDown = self.on_mouse_left_down
    #             hook_manager.MouseLeftUp = self.on_mouse_left_up
    #
    #             print("Mouse tracking started")
    #         else:
    #             print("on_mouse_left_down or on_mouse_left_up not callable")
    #
    #     except Exception as e:
    #         print(f"Error in start_mouse_tracking: {e}")
    #     PumpMessages()
    #     # Unhook the mouse when done
    #     hook_manager.UnhookMouse()
    #     print("Mouse tracking stopped")

# def on_mouse_left_down(event):
#     try:
#         print("Mouse left button pressed")
#         global start_mouse_pos
#         # Запомнить начальную позицию мыши
#         start_mouse_pos = (event.Position[0], event.Position[1])
#         return True
#     except Exception as e:
#         print(f"Exception in on_mouse_left_down: {e}")
#         return False
#
#
# def on_mouse_left_up(event):
#     try:
#         print("Mouse left button released")
#         #+startr_mouse_pos)
#         global end_mouse_pos
#         # Запомнить конечную позицию мыши
#
#         end_mouse_pos = (event.Position[0], event.Position[1])
#         # Проверить, что позиции мыши были захвачены
#         if start_mouse_pos and end_mouse_pos:
#             #Вычислить расстояние между начальной и конечной позициями
#             distance = ((end_mouse_pos[0] - start_mouse_pos[0]) ** 2 + (
#                         end_mouse_pos[1] - start_mouse_pos[1]) ** 2) ** 0.5
#             print(f"Distance: {distance:.2f}")
#             if distance > MIN_MOUSE_DRAWN_DISTANCE:
#                 signal_queue_got.put("text_selected")
#                 print("Text selected")
#
#         else:
#             print("Mouse positions not captured properly")
#         return True
#     except Exception as e:
#         print(f"Exception in on_mouse_left_up: {e}")
#         return False
#
#
# def start_mouse_tracking():
#     hook_manager = pyWinhook.HookManager()
#     hook_manager.HookMouse()
#
#     try:
#         if callable(on_mouse_left_down) and callable(on_mouse_left_up):
#             # Subscribe to left mouse down and up events
#             hook_manager.MouseLeftDown = on_mouse_left_down
#             hook_manager.MouseLeftUp = on_mouse_left_up
#
#         else:
#             print("on_mouse_left_down or on_mouse_left_up not callable")
#
#     except Exception as e:
#         print(f"Error in start_mouse_tracking: {e}")
#
#     import pythoncom
#     PumpMessages()
#
#     # Unhook the mouse when done
#     hook_manager.UnhookMouse()








# Run the main asynchronous function
if __name__ == "__main__":
    print("Program started")
    asyncio.run(main())
