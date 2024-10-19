""" GUI Functions 
    src/projectutils/guifunc.py 
    This file contains the functions with graphical user interfaces
    Author: vipulasrilanka@yahoo.com 
    (c) 2024 """
import tkinter as tk
from tkinter import filedialog, simpledialog
from constants.errorcodes import ERROR_SUCCESS

MESSAGE_NEW = 1
MESSAGE_ADD = 2
MESSAGE_CLEAR = 3
WINDOW_QUIT = 0
GET_PASSWORD = 4
WAIT_FOR_PASSWORD = 5
RETURN_PASSWORD = 6

def getPdfFileName(dialogTitle, initDir):
    """Select a PDF File"""
    # Create a root window (but hide it)
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Open file dialog and allow user to select a file
    filePath = filedialog.askopenfilename(
        title=dialogTitle, 
        filetypes=[("PDF file", "*.pdf")],
        initialdir=initDir 
        )
    #destroy the window. (quit does not work here)
    root.destroy()
    # Return the file path selected by the user
    return filePath

def getExcelFileName(dialogTitle, initDir):
    """Select an Excel file"""
    # Create a root window (but hide it)
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Open file dialog and allow user to select a file
    filePath = filedialog.askopenfilename(
        title=dialogTitle, 
        filetypes=[("Microsoft Excel file", "*.xlsx")],
        initialdir=initDir 
        )
    #destroy the window. (quit does not work here)
    root.destroy()
    # Return the file path selected by the user
    return filePath

def getPassword(fileName):
    """Get password to open file with fileName"""
    # Create the root window (but hide it)
    root = tk.Tk()
    root.withdraw()  # Hide the main Tkinter window

    # Prompt the user to enter a password, mask the input with '*'
    password = simpledialog.askstring(
        title="Password Required", 
        prompt=f"Enter password for {fileName}:", 
        show="*"
    )
    #destroy the window
    root.destroy()
    # Return the password entered by the user
    return password

def showStatus(messageHolder, windowName):
    """Pop up a GUI with a read-only text box displaying messageHolder."""
    print("+Fn showStatus")
    lastAction = str(None)
    def messageProcesser():
        """Update the text box with the latest value of the message."""
        nonlocal root, text_box, lastAction
        messageId = messageHolder["id"]
        action = messageHolder["action"]
        message = messageHolder["message"]  # Access the first element of the list
        if(lastAction == str(messageId) + str(message) + str(action)):
            #no update, just loop.
            root.after(500, messageProcesser) 
            return
        lastAction = str(messageId) + str(message) + str(action)
        if action == WINDOW_QUIT:
            # Close the window 
            root.destroy() 
            # root.quit()  
            return
        if(action == GET_PASSWORD):
            #open the password dialog box
            messageHolder["action"] = WAIT_FOR_PASSWORD
            messageHolder["message"] = simpledialog.askstring(
                title="Password Required", 
                prompt=f"{message}", 
                show="*")
            #return the password
            messageHolder["action"] = RETURN_PASSWORD
            root.after(500, messageProcesser) 
            return
        if (action == RETURN_PASSWORD or action == WAIT_FOR_PASSWORD):
            #no action it we get here.
            root.after(500, messageProcesser) 
            return
        #It is a message process
        text_box.config(state=tk.NORMAL)  # Make text box editable to update content
        if (action == MESSAGE_NEW or action == MESSAGE_CLEAR):
            text_box.delete(1.0, tk.END)      # Clear the current content
        if action == MESSAGE_NEW:
            text_box.insert(tk.END, str(messageId) + " " + message)  # Insert the latest message
        elif action == MESSAGE_ADD:
            text_box.insert(tk.END, "\n"+ str(messageId) + " " + message)
        else:
            print("Message box = Clear")
        text_box.config(state=tk.DISABLED)  # Make text box read-only again
        root.after(500, messageProcesser)  # Call this function again after 1 second (1000 ms)
    # Create the main window
    root = tk.Tk()
    root.title(windowName)
    # Get the screen width and height
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    # Set window width to half the screen width and height to one-quarter of the screen height
    window_width = int(screen_width / 2)
    window_height = int(screen_height / 4)
    # Set the geometry for the window to open at the top-left corner (0, 0)
    root.geometry(f"{window_width}x{window_height}+0+0")
    # Create a Text widget for showing the message
    text_box = tk.Text(root, height=10, width=50)
    text_box.pack(padx=10, pady=10)
    # Make the text box read-only
    text_box.config(state=tk.DISABLED)
    # Start updating the text box with the current value of the message
    messageProcesser()
    # Start the Tkinter event loop
    root.mainloop()
    #return
    print("-Fn showStatus")
    return ERROR_SUCCESS