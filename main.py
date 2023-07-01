import asyncio
import pygetwindow as gw
import pyautogui
import win32api
import win32clipboard
import pyWinhook
import threading
import time

# Глобальный флаг для отслеживания выделения текста с помощью мыши
import win32con

# Глобальные переменные для отслеживания позиции мыши
start_mouse_pos = None
mouse_selecting = False


def on_mouse_event(event):
    global mouse_selecting
    if event.MessageName == "mouse left down":
        mouse_selecting = True
    elif event.MessageName == "mouse left up":
        mouse_selecting = False
    return True



def on_mouse_left_down(event):
    global start_mouse_pos
    print("Mouse left button pressed")  # Add this line
    start_mouse_pos = (event.Position[0], event.Position[1])
    return True

def on_mouse_left_up(event):
    global mouse_selecting
    # Записать конечную позицию мыши
    end_mouse_pos = (event.Position[0], event.Position[1])

    # Вычислить расстояние между начальной и конечной позициями
    distance = ((end_mouse_pos[0] - start_mouse_pos[0]) ** 2 + (end_mouse_pos[1] - start_mouse_pos[1]) ** 2) ** 0.5
    print(f"Mouse left button released, distance: {distance}")


    # Получить горизонтальное разрешение экрана
    screen_width = win32api.GetSystemMetrics(0)

    # Вычислить среднюю ширину символа на экране (предположим, что средний символ занимает 1/80 ширины экрана)
    char_width = screen_width / 80

    # Если расстояние больше, чем ширина 3 символов, установить флаг
    if distance > char_width * 3:
        mouse_selecting = True
    else:
        mouse_selecting = False
    return True

def start_mouse_tracking():
    hook_manager = pyWinhook.HookManager()

    # Subscribe to left mouse down and up events
    hook_manager.SubscribeMouseLeftDown(on_mouse_left_down)
    hook_manager.SubscribeMouseLeftUp(on_mouse_left_up)

    import pythoncom
    pythoncom.PumpMessages()

    # Unhook the mouse when done
    hook_manager.UnhookMouse()


mouse_tracking_thread = threading.Thread(target=start_mouse_tracking)
mouse_tracking_thread.daemon = True
mouse_tracking_thread.start()


async def print_selected_text():
    global mouse_selecting
    last_text = None
    selection_time = 0
    last_print_time = 0

    while True:
        await asyncio.sleep(1)
        active_window = gw.getActiveWindow()
        if active_window and mouse_selecting:
            # Save the current clipboard content
            win32clipboard.OpenClipboard()
            try:
                original_clipboard = win32clipboard.GetClipboardData()
            except TypeError:
                original_clipboard = None
            win32clipboard.CloseClipboard()

            # Press CTRL+C to copy the selected text to clipboard
            # Simulate CTRL+C (VK_CONTROL=0x11, VK_C=0x43)
            win32api.keybd_event(0x11, 0, 0, 0)  # Press CTRL
            win32api.keybd_event(0x43, 0, 0, 0)  # Press C
            win32api.keybd_event(0x43, 0, win32con.KEYEVENTF_KEYUP, 0)  # Release C
            win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)  # Release CTRL
            await asyncio.sleep(0.1)

            # Get the text from the clipboard
            try:
                win32clipboard.OpenClipboard()
                try:
                    text = win32clipboard.GetClipboardData()
                except TypeError:
                    text = None
                win32clipboard.CloseClipboard()
            except Exception as e:
                print(f"Error: {e}")


            # Restore the original clipboard content
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            if original_clipboard is not None:
                win32clipboard.SetClipboardText(original_clipboard)
            win32clipboard.CloseClipboard()

            if text == last_text:
                selection_time += 1
            else:
                selection_time = 0

            last_text = text
            current_time = time.time()

            if selection_time >= 2 and current_time - last_print_time > 2:
                print(f"Application: {active_window.title}")
                print(f"Selected Text: {text}\n")
                last_print_time = current_time


async def main():
    print("Program started")
    asyncio.create_task(print_selected_text())
    while True:
        await asyncio.sleep(1)


asyncio.run(main())
