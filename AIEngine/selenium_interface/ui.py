from appJar import gui
from selenium_interface.settings import *
import time


class ui(gui):
    '''
    Handles behind-the-scenes operations of the user interface.
    '''
    def __init__(self, begin_function=(lambda : exit())):
        self.begin = begin_function
        gui.__init__(self)
        self.rpa = None
        self.log_path = ''

        self.setIcon("./icon.ico")
        buttonStyle = {"highlightthickness": 0, "bd": 0, "fg": "#f1f2f6", "bg": "#2f3542",
                       "activebackground": "#1f2532", "activeforeground": "#e1e2e6"}
        boxStyle = {"highlightthickness": 0, "bd": 0, "fg": "#f1f2f6", "bg": "#57606f"}
        bgStyle = {"fg": "#f1f2f6", "bg": "#2f3542", "activebackground": "#1f2532", "activeforeground": "#e1e2e6"}

        self.startSubWindow("Login")
        self.setBg("#2f3542", override=True)
        self.setFg("#f1f2f6", override=True)

        box = self.addLabelEntry("Username")
        self.setEntryDefault("Username", UID)
        box.config(**boxStyle)

        box = self.addLabelSecretEntry("Password")
        self.setEntryDefault("Password", ''.join(["*" for ch in PW]))
        box.config(**boxStyle)

        frame = self.addLabelDirectoryEntry("PM directory")
        button = frame.theButton
        button.config(**buttonStyle)
        box = frame.theWidget
        box.config(**boxStyle)
        self.setEntryDefault("PM directory", RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_POLICY_MANIPULATION)

        frame = self.addLabelDirectoryEntry("MS directory")
        button = frame.theButton
        button.config(**buttonStyle)
        box = frame.theWidget
        box.config(**boxStyle)
        self.setEntryDefault("MS directory", RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_MEMBER_SETUP)

        frame = self.addLabelDirectoryEntry("Upload directory")
        button = frame.theButton
        button.config(**buttonStyle)
        box = frame.theWidget
        box.config(**boxStyle)
        self.setEntryDefault("Upload directory", RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_BULK_UPLOAD)

        button = self.addImageButton("Begin", self.press, "start.png")
        button.config(highlightthickness=0, bd=0, **bgStyle)

        self.enableEnter(self.press)
        self.stopSubWindow()

        self.startSubWindow("Running")
        self.setFg("#FDFDFD", override=True)
        self.setBg("#050505", override=True)
        self.addMessage("run", "Logging in")
        self.addEmptyMessage("working")
        tmp = self.addScrolledTextArea("log")
        tmp.config(width=90, height=20, **boxStyle)
        self.addEmptyMessage("elapsed")
        self.setMessageWidth("run", 390)
        self.setMessageWidth("working", 390)
        self.setMessageWidth("elapsed", 300)
        self.stopSubWindow()

        self.go(startWindow="Login")

    def run(self):
        '''
        Invokes recurring functions within ui
        '''
        self.update()
        self.track_window()

    def press(self, button):
        '''
        Event that starts main program operation after user enters credentials.

        :param button:                          object pointer for button that was pressed, unused, mandatory for appjar event.
        '''
        f = open("./rpa/settings.py", "r")
        entries = {}

        if self.getEntry("Username") != "":
            entries["UID"] = self.getEntry("Username")
        else:
            entries["UID"] = UID
        if self.getEntry("Password") != "":
            entries["PW"] = self.getEntry("Password")
        else:
            entries["PW"] = PW
        if self.getEntry("MS directory") != "":
            entries["RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_MEMBER_SETUP"] = self.getEntry("MS directory")
        else:
            entries[
                "RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_MEMBER_SETUP"] = RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_MEMBER_SETUP
        if self.getEntry("PM directory") != "":
            entries["RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_POLICY_MANIPULATION"] = self.getEntry("PM directory")
        else:
            entries[
                "RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_POLICY_MANIPULATION"] = RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_POLICY_MANIPULATION
        if self.getEntry("Upload directory") != "":
            entries["RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_BULK_UPLOAD"] = self.getEntry("Upload directory")
        else:
            entries[
                "RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_BULK_UPLOAD"] = RELATIVE_PATH_TO_FOLDER_OF_INPUT_XLSX_FOR_BULK_UPLOAD

        new = self.settings_replace(entries, f)
        f.close()
        f = open("./rpa/settings.py", "w")
        f.write(new)

        self.hideSubWindow("Login")
        self.showSubWindow("Running")
        self.setMessage("run", "Logging in")
        self.begin(entries, self)

    def settings_replace(self, entries, text):
        '''
        Method to handle modifying strings from the settings file that were different from current user session input

        :param list entries:                                       The entries in the settings file to be changed
        :param list text:                                           The complete contents of the original settings file
        :rtype:                                                         list
        '''
        out = ""
        for line in text:
            n = 0
            for field, value in entries.items():
                if line.find(field + " = '") == -1:
                    continue
                else:
                    out += field + " = '" + value + "'\n"
                    n = 1
                    break
            if n == 0:
                out += line
        return out

    def add_rpa(self, rpa):
        '''
        Set internal values from the RPA module

        :param RPA rpa:                             RPA module in use
        '''
        self.rpa = rpa
        self.log_path = rpa.logger._Logger__full_path_to_log_file

    def update(self):
        '''
        Writes the contents of a file to the text area in the window.

        :param str file_path:                         Path to file to be used as display
        '''
        self.setMessage("elapsed", str(int(time.clock())) + " seconds elapsed")
        f = open(self.log_path, 'r')
        self.clearTextArea('log')
        self.setTextArea('log', f.read())
        self.after(500, self.update)

    def track_window(self):
        '''
        Checks periodically to ensure that browser window is still open.
        '''
        try:
            self.rpa.driver.window_handles
        except:
            self.rpa.logger.info(
                "EXIT",
                "window closed"
            )
            self.stop()
            exit()
        self.after(3000, self.track_window)
