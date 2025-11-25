import keyboard
import time
import win32clipboard as wc
import tkinter as tk

original_clipboard = {"dtype": str, "data": int}
artificial_clipboard = {"dtype": str, "data": int}


def set_data(clipboard) -> None:
    """"
    This sets the data of a given clipboard (original or artificial) to be what is on the computer's clipboard
    :param clipboard: a clipboard (artificial_clipboard or original_clipboard)
    """
    if wc.IsClipboardFormatAvailable(wc.CF_DIB):
        clipboard["dtype"] = wc.CF_DIB
        clipboard["data"] = wc.GetClipboardData(wc.CF_DIB)
    elif wc.IsClipboardFormatAvailable(wc.CF_UNICODETEXT):
        clipboard["dtype"] = wc.CF_UNICODETEXT
        clipboard["data"] = wc.GetClipboardData(wc.CF_UNICODETEXT)
    elif wc.IsClipboardFormatAvailable(wc.CF_OEMTEXT):
        clipboard["dtype"] = wc.CF_OEMTEXT
        clipboard["data"] = wc.GetClipboardData(wc.CF_OEMTEXT)
    elif wc.IsClipboardFormatAvailable(wc.CF_BITMAP):
        clipboard["dtype"] = wc.CF_BITMAP
        clipboard["data"] = wc.GetClipboardData(wc.CF_BITMAP)
    elif wc.IsClipboardFormatAvailable(wc.CF_ENHMETAFILE):
        clipboard["dtype"] = wc.CF_ENHMETAFILE
        clipboard["data"] = wc.GetClipboardData(wc.CF_ENHMETAFILE)
    elif wc.IsClipboardFormatAvailable(wc.CF_TEXT):
        clipboard["dtype"] = wc.CF_TEXT
        clipboard["data"] = wc.GetClipboardData(wc.CF_TEXT)
    elif wc.IsClipboardFormatAvailable(wc.CF_METAFILEPICT):
        clipboard["dtype"] = wc.CF_METAFILEPICT
        clipboard["data"] = wc.GetClipboardData(wc.CF_METAFILEPICT)
    else:
        return


def load_data(clipboard) -> None:
    """
    This takes data from a given clipboard and places it into the user's windows clipboard
    :param clipboard: one of the clipboards (see above)
    """
    try:
        wc.EmptyClipboard()
        wc.SetClipboardData(clipboard["dtype"], clipboard["data"])
    finally:
        return


def copy_function() -> None:
    """
    This is called by the hotkeys from the keyboard module once all the copying values (see binds dict) are pressed
    """

    # Variables are made global to allow for them to keep their data after the function concludes
    global original_clipboard
    global artificial_clipboard
    wc.OpenClipboard()
    set_data(original_clipboard)
    wc.EmptyClipboard()

    # The original clipboard sensitive code was placed into a try finally statement in case things break
    try:
        wc.CloseClipboard()
        # Waits until all the keys are released so that the user's fingers on the keyboard doesn't disrupt the future ctrl c/ctrl v presses
        while keyboard.is_pressed(binds['cv1']) or keyboard.is_pressed(binds['cv2']) or keyboard.is_pressed(binds['cv3']):
            time.sleep(0.005)
        keyboard.press('ctrl')
        keyboard.press('c')
        # Short pause is necessary, else it presses and releases it all on practically the same frame, before the computer can register
        time.sleep(0.1)
        keyboard.release('c')
        keyboard.release('ctrl')
        wc.OpenClipboard()
        set_data(artificial_clipboard)

    finally:
        load_data(original_clipboard)
        wc.CloseClipboard()


def paste_function() -> None:
    """
    The same situation as copy_function(), however this one does paste and some different operations
    """
    global original_clipboard
    global artificial_clipboard
    wc.OpenClipboard()
    set_data(original_clipboard)
    # This is also wrapped in a try, finally statement just for extra security
    try:
        load_data(artificial_clipboard)
        wc.CloseClipboard()
        while keyboard.is_pressed(binds['pv1']) or keyboard.is_pressed(binds['pv2']) or keyboard.is_pressed(binds['pv3']):
            time.sleep(0.01)
        keyboard.press('ctrl')
        keyboard.press('v')
        time.sleep(0.1)
        keyboard.release('v')
        keyboard.release('ctrl')
        wc.OpenClipboard()
    finally:
        load_data(original_clipboard)
        wc.CloseClipboard()

"""
This is where all of the keybind data is stored. This is probably the most important dictionary in this program
It comes with a set of default values as well.
"""
binds: dict = {"cv1": "ctrl",
               "cv2": "shift",
               "cv3": "f1",
               "pv1": "ctrl",
               "pv2": "shift",
               "pv3": "f2"}

# A little variable to indicate if the hotkeys need to be turned on or off
run = True


""""
From here on out, most of it is UI work. 
"""
window = tk.Tk() # New window
window.geometry("333x233") # Dimensions
window.title("Secondary Copy and Paste") # Title


# buttons are placed into a dictionary for the sake of iteration
button_storage = {"cv1": tk.Button(window, text='Set key 1', font=("Helvetica", 9), bg="#E0E0E0", command=lambda: set_val("cv1")),
                 "cv2": tk.Button(window, text='Set key 2', font=("Helvetica", 9), bg="#E0E0E0", command=lambda: set_val("cv2")),
                 "cv3": tk.Button(window, text='Set key 3', font=("Helvetica", 9), bg="#E0E0E0", command=lambda: set_val("cv3")),
                 "pv1": tk.Button(window, text='Set key 1', font=("Helvetica", 9), bg="#E0E0E0", command=lambda: set_val("pv1")),
                 "pv2": tk.Button(window, text='Set key 2', font=("Helvetica", 9), bg="#E0E0E0", command=lambda: set_val("pv2")),
                 "pv3": tk.Button(window, text='Set key 3', font=("Helvetica", 9), bg="#E0E0E0", command=lambda: set_val("pv3"))}
button_storage["cv1"].place(x=73, y=83, height=33, width=67)
button_storage["cv2"].place(x=157, y=83, height=33, width=67)
button_storage["cv3"].place(x=240, y=83, height=33, width=67)
button_storage["pv1"].place(x=73, y=150, height=33, width=67)
button_storage["pv2"].place(x=157, y=150, height=33, width=67)
button_storage["pv3"].place(x=240, y=150, height=33, width=67)

start_button = tk.Button(window, text="Enable alternate copy/paste", bg="#E0E0E0", font=("Helvetica",9), command=lambda: stop_start())
start_button.place(x=73, y=33, height=33, width=233)

copy_label = tk.Label(window, text="Copying\n keybind:", font=("Helvetica", 9))
copy_label.place(x=13, y=90)
paste_label = tk.Label(window, text="Pasting\n keybind:", font=("Helvetica", 9))
paste_label.place(x=13, y=157)
locked_label = tk.Label(window, text="Keybinds\nunlocked", font=("Helvetica", 9))
locked_label.place(x=13, y=37)


# Updates font size for the rest of the bits
# Labels are stored in a dictionary for the sake of iteration

label_storage = {"cv1": tk.Label(window, text=binds["cv1"], anchor="n", font=("Helvetica", 11)),
                 "cv2": tk.Label(window, text=binds["cv2"], anchor="n", font=("Helvetica", 11)),
                 "cv3": tk.Label(window, text=binds["cv3"], anchor="n", font=("Helvetica", 11)),
                 "pv1": tk.Label(window, text=binds["pv1"], anchor="n", font=("Helvetica", 11)),
                 "pv2": tk.Label(window, text=binds["pv2"], anchor="n", font=("Helvetica", 11)),
                 "pv3": tk.Label(window, text=binds["pv3"], anchor="n", font=("Helvetica", 11))}
label_storage["cv1"].place(x=73, y=117)
label_storage["cv2"].place(x=157, y=117)
label_storage["cv3"].place(x=240, y=117)
label_storage["pv1"].place(x=73, y=183)
label_storage["pv2"].place(x=157, y=183)
label_storage["pv3"].place(x=240, y=183)


def set_val(val) -> None:
    """
    This sets the value of a label, ones that are listed directly above.
    Unless the program is running, in which case the program can't change the keybinds
    :param val: every index in binds. Everything from cv1 to pv3
    """
    global binds
    binds[val] = keyboard.read_key()
    label_storage[val].config(text=binds[val])


def cv_duplicate_check() -> str:
    """
    As the name describes, this checks to see if the copy value list has a duplicate.
    This is solely used for making it so that the user can put multiple of the same keybind into the program
    without it breaking since, due to the keyboard's preferred format, having multiple of the same makes it do nothing.

    """
    temp_dict = {"cv1": binds["cv1"], "cv2": binds["cv2"], "cv3": binds["cv3"]}
    new_binding = []  # empty list
    for key, data in temp_dict.items():
        if data not in new_binding:
            new_binding.append(data)
    new_text: str = new_binding[0]
    new_binding.pop(0)
    for item in new_binding:
        new_text = new_text + f"+{item}"
    return new_text


def pv_duplicate_check() -> str:
    """
    This is effectively the same as cv_duplicate_check(), but with the temp_dict swapped.
    I could set this up to be one method for both, but then I would have to separate copying values and paste values
    which would've added more lines of code than saved.
    """
    temp_dict = {"pv1": binds["pv1"], "pv2": binds["pv2"], "pv3": binds["pv3"]}
    new_binding = []  # empty list
    for key, data in temp_dict.items():
        if data not in new_binding:
            new_binding.append(data)

    # This little portion is to avoid having an awkward "+" at the start of new_text
    new_text: str = new_binding[0]
    new_binding.pop(0)
    for item in new_binding:
        new_text = new_text + f"+{item}"
    return new_text


def stop_start() -> None:
    """
    This changes what the effective keybind is and the text on the start/stop button
    """
    global run
    if run:
        start_button.config(text="Pause alternate copy/paste")
        locked_label.config(text="Keybinds\nlocked")
        keyboard.add_hotkey(cv_duplicate_check(), copy_function)
        keyboard.add_hotkey(pv_duplicate_check(), paste_function)
        for key in button_storage:
            button_storage[key].config(state="disabled")
        for key in label_storage:
            label_storage[key].config(fg = "#948678")
        run = not run
    else:
        start_button.config(text="Enable alternate copy/paste")
        locked_label.config(text="Keybinds\nunlocked")
        keyboard.clear_all_hotkeys()
        for key in button_storage:
            button_storage[key].config(state="normal")
        for key in label_storage:
            label_storage[key].config(fg = "#000000")
        run = not run


window.mainloop()

# Note: Ran on computer startup b/c that requires transporting the python file across the user's directories and can mess with sensitive data
