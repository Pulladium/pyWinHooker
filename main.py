import asyncio
import pygetwindow as gw
import pyautogui
import pyperclip
import time
import win32api
import win32con
import pyWinhook

# Инициализация переменных позиции мыши
from win32gui import PumpMessages

start_mouse_pos = ()
end_mouse_pos = ()


def on_mouse_left_down(event):
    try:
        print("Mouse left button pressed")
        global start_mouse_pos
        # Запомнить начальную позицию мыши
        start_mouse_pos = (event.Position[0], event.Position[1])
        return True
    except Exception as e:
        print(f"Exception in on_mouse_left_down: {e}")
        return False



def on_mouse_left_up(event):
    try:
        print("Mouse left button released")
        global end_mouse_pos
        # Запомнить конечную позицию мыши
        end_mouse_pos = (event.Position[0], event.Position[1])
        # Проверить, что позиции мыши были захвачены
        if start_mouse_pos and end_mouse_pos:
            # Вычислить расстояние между начальной и конечной позициями
            distance = ((end_mouse_pos[0] - start_mouse_pos[0]) ** 2 + (
                        end_mouse_pos[1] - start_mouse_pos[1]) ** 2) ** 0.5
            print(f"Distance: {distance:.2f}")
        else:
            print("Mouse positions not captured properly")
        return True
    except Exception as e:
        print(f"Exception in on_mouse_left_up: {e}")
        return False


def start_mouse_tracking():
    hook_manager = pyWinhook.HookManager()
    hook_manager.HookMouse()

    try:
        if callable(on_mouse_left_down) and callable(on_mouse_left_up):
            # Subscribe to left mouse down and up events
            hook_manager.MouseLeftDown = on_mouse_left_down
            hook_manager.MouseLeftUp = on_mouse_left_up
        else:
            print("on_mouse_left_down or on_mouse_left_up not callable")

    except Exception as e:
        print(f"Error in start_mouse_tracking: {e}")

    import pythoncom
    PumpMessages()
    # Unhook the mouse when done
    hook_manager.UnhookMouse()


async def print_selected_text():
    """Asynchronously print the selected text."""
    last_text = None
    selection_time = 0
    printed = False

    while True:
        # Wait for 1 second
        await asyncio.sleep(1)
        # Get the active window
        active_window = gw.getActiveWindow()
        # If there is an active window
        if active_window:
            # Press CTRL+C to copy the selected text to clipboard
            win32api.keybd_event(0x11, 0, 0, 0)  # Press CTRL
            win32api.keybd_event(0x43, 0, 0, 0)  # Press C
            win32api.keybd_event(0x43, 0, win32con.KEYEVENTF_KEYUP, 0)  # Release C
            win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)  # Release CTRL

            # Wait for the clipboard to update
            await asyncio.sleep(0.1)

            # Get the text from the clipboard
            text = pyperclip.paste()

            # Check if the text is the same as before
            if text == last_text:
                # Increment the selection time
                selection_time += 1
            else:
                # Reset the selection time and printed flag
                selection_time = 0
                printed = False

            # Update the last text
            last_text = text

            # If thetext has been selected for more than 2 seconds and not printed yet, print it
            if selection_time >= 2 and not printed:
                # Print the application title
                print(f"Application: {active_window.title}")

                # Print the text
                print(f"Selected Text: {text}\n")

                # Set the printed flag to True
                printed = True


async def main():
    """Main asynchronous function."""
    # Start mouse tracking
    start_mouse_tracking()

    # Start the asynchronous task to print selected text
    asyncio.create_task(print_selected_text())

    # Keep the main function running
    while True:
        await asyncio.sleep(1)


# Run the main asynchronous function
if __name__ == "__main__":
    print("Program started")
    asyncio.run(main())