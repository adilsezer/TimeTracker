import win32process
import win32gui
import win32api
import psutil
import time
import tkinter as tk
import tkinter.ttk as ttk
from ttkthemes import ThemedTk
from tkinter import messagebox
import re
import threading
import datetime


class TimeTrackerGUI:
    def __init__(self, master):
        self.master = master
        master.set_theme("yaru")
        master.title("App Time Tracker")
        master.minsize(300, 200)
        self.RedStyle = ttk.Style()

        menu = tk.Menu(master)
        master.option_add("*tearOff", False)
        master.config(menu=menu)
        fileMenu = tk.Menu(menu)
        menu.add_cascade(label="File", menu=fileMenu)
        editMenu = tk.Menu(menu)
        menu.add_cascade(label="Help", menu=editMenu)

        fileMenu.add_command(label="Quit", command=self.QuitApp)
        editMenu.add_command(label="About", command=self.AboutMenu)

        self.TopFrame = ttk.LabelFrame(master, text="Application's Run Time",)
        self.TopFrame.pack(fill="both", expand="yes")

        self.BottomFrame = ttk.Frame(master)
        self.BottomFrame.pack(side="bottom")

        self.StatusString = tk.StringVar()
        self.StatusWidget = ttk.Entry(
            self.BottomFrame,
            textvariable=self.StatusString,
            justify=tk.CENTER,
            font=("Arial", 9, "bold"),
        ).pack()

        self.TopLabel = ttk.Label(self.TopFrame, text="")
        self.TopLabel.pack()

        self.MidLabel = ttk.Label(master, text="")
        self.MidLabel.pack(side="top", fill=tk.X)

        self.StopResumeBtn = ttk.Button(
            text="Stop/Resume", command=self.StopResumeTimer
        )
        self.StopResumeBtn.pack(side="left")

        self.Restartbtn = ttk.Button(text="Restart Timer", command=self.RestartTimer)
        self.Restartbtn.pack(side="left")

        self.Addbtn = ttk.Button(text="Add Item", command=self.AddAppTime)
        self.Addbtn.pack(side="left")

        self.Deletebtn = ttk.Button(text="Delete Item", command=self.DeleteAppTime)
        self.Deletebtn.pack(side="left")

        self.ModifyBtn = ttk.Button(text="Modify Item", command=self.ModifyAppTime)
        self.ModifyBtn.pack(side="left")

        self.Refresher()

    def Refresher(self):
        global status
        if status == "Stopped":
            self.StatusString.set("Status: " + status)
            self.RedStyle.configure("TEntry", foreground="red")
        else:
            self.TopLabel.configure(text=maincode(quote))
            self.MidLabel.configure(
                text="Total Time: " + str(convertSeconds(totaltime))
            )
            self.RedStyle.configure("TEntry", foreground="green")
        self.StatusString.set("Status: " + status)
        self.TopLabel.after(1000, self.Refresher)

    def RestartTimer(self):
        global process_time
        global timestamp
        global totaltime

        process_time = {}
        timestamp = {}
        totaltime = 0

    def StopResumeTimer(self):
        global status
        if not status == "Stopped":
            status = "Stopped"
        else:
            status = "Active"

    def ModifyAppTime(self):
        self.ModifyWindow = tk.Toplevel(self.master)
        self.ModifyApp = ModifyAppWindow(self.ModifyWindow, self.TopLabel)

    def DeleteAppTime(self):
        self.DeleteWindow = tk.Toplevel(self.master)
        self.DeleteApp = DeleteAppWindow(self.DeleteWindow, self.TopLabel)

    def AddAppTime(self):
        self.AddWindow = tk.Toplevel(self.master)
        self.AddApp = AddAppWindow(self.AddWindow)

    def AboutMenu(self):
        messagebox.showinfo(
            title="App Time Tracker",
            message="This application was developed to calculate your active application times",
        )

    def QuitApp(self):
        self.master.destroy()


class ModifyAppWindow:
    def __init__(self, master, TopLabel):
        self.master = master
        self.TopLabel = TopLabel

        self.AppTimes = self.TopLabel.cget("text").split("\n")[:-1]

        self.AppLabel = ttk.Label(self.master, text="Select an Application").grid(
            row=0, column=0, padx=3, pady=3
        )
        self.AppName = ttk.Combobox(self.master, values=(self.AppTimes), width=10)
        self.AppName.grid(
            row=0, column=1, padx=3, pady=3, sticky=tk.W + tk.E + tk.N + tk.S
        )
        self.TimeLabel = ttk.Label(self.master, text="Enter a Time\n(hh:mm:ss)").grid(
            row=1, column=0, padx=3, pady=3, sticky=tk.W + tk.E + tk.N + tk.S
        )

        self.AppTimeEntry = tk.StringVar()
        self.AppTime = ttk.Entry(self.master, textvariable=self.AppTimeEntry).grid(
            row=1, column=1, padx=3, pady=3
        )

        self.AppTimeEntry.set("00:00:00")

        self.ModifyButton = ttk.Button(
            self.master, text="Apply", command=self.ModifyDict
        ).grid(columnspan=2, padx=3, pady=3)

    def ModifyDict(self):
        global process_time
        global totaltime
        AppName = self.AppName.get()
        AppTime = self.AppTimeEntry.get()
        pattern = r"\d\d:\d\d:\d\d"
        result = re.match(pattern, AppTime)
        if not result or AppName == "":
            messagebox.showinfo(
                parent=self.master,
                title="Error",
                message="Please select an application and enter the time in hh:mm:ss format!",
            )

        else:
            AppName = AppName.replace(":", "").split(" ")
            date_time = datetime.datetime.strptime(AppTime, "%H:%M:%S")
            a_timedelta = date_time - datetime.datetime(1900, 1, 1)
            timedif = a_timedelta.total_seconds() - process_time[AppName[0]]
            totaltime += timedif
            del process_time[AppName[0]]
            process_time[AppName[0]] = a_timedelta.total_seconds()
            messagebox.showinfo(
                parent=self.master,
                title="App Modified",
                message="Application time for "
                + str(AppName[0])
                + " has been successfully modified",
            )
            self.AppName["values"] = self.TopLabel.cget("text").split("\n")[:-1]


class DeleteAppWindow:
    def __init__(self, master, TopLabel):
        self.master = master
        self.TopLabel = TopLabel

        AppTimes = TopLabel.cget("text").split("\n")[:-1]

        self.AppLabel = ttk.Label(self.master, text="Select an Application").grid(
            row=0, column=0, padx=3, pady=3
        )
        self.AppName = ttk.Combobox(self.master, values=(AppTimes), width=10)
        self.AppName.grid(
            row=0, column=1, padx=3, pady=3, sticky=tk.W + tk.E + tk.N + tk.S
        )

        self.DeleteButton = ttk.Button(
            self.master, text="Apply", command=self.DeleteDict
        ).grid(columnspan=2, padx=3, pady=3)

    def DeleteDict(self):
        global process_time
        global totaltime
        AppName = self.AppName.get()
        if AppName == "":
            messagebox.showinfo(
                parent=self.master,
                title="Error",
                message="Please select an application!",
            )
        else:
            AppName = AppName.replace(":", "").split(" ")
            totaltime -= process_time[AppName[0]]
            del process_time[AppName[0]]
            messagebox.showinfo(
                parent=self.master,
                title="App Deleted",
                message="Application time for "
                + str(AppName[0])
                + " has been successfully deleted",
            )
            self.AppName["values"] = self.TopLabel.cget("text").split("\n")[:-1]


class AddAppWindow:
    def __init__(self, master):
        self.master = master

        self.AppLabel = ttk.Label(self.master, text="Enter an Application").grid(
            row=0, column=0, padx=3, pady=3
        )

        self.AppNameEntry = tk.StringVar()
        self.AppName = ttk.Entry(self.master, textvariable=self.AppNameEntry).grid(
            row=0, column=1, padx=3, pady=3
        )

        self.TimeLabel = ttk.Label(self.master, text="Enter a Time\n(hh:mm:ss)").grid(
            row=1, column=0, padx=3, pady=3, sticky=tk.W + tk.E + tk.N + tk.S
        )

        self.AppTimeEntry = tk.StringVar()
        self.AppTime = ttk.Entry(self.master, textvariable=self.AppTimeEntry).grid(
            row=1, column=1, padx=3, pady=3
        )

        self.AppTimeEntry.set("00:00:00")

        self.AddButton = ttk.Button(
            self.master, text="Apply", command=self.AddDict
        ).grid(columnspan=2, padx=3, pady=3)

    def AddDict(self):
        global process_time
        global timestamp
        global totaltime
        AppName = self.AppNameEntry.get()
        AppTime = self.AppTimeEntry.get()
        pattern = r"\d\d:\d\d:\d\d"
        result = re.match(pattern, AppTime)
        if not result or AppName == "":
            messagebox.showinfo(
                parent=self.master,
                title="Error",
                message="Please select an application and enter the time in hh:mm:ss format!",
            )
        elif AppName.lower() in set(k.lower() for k in process_time):
            messagebox.showinfo(
                parent=self.master,
                title="Error",
                message="This application already in application list!",
            )
        else:
            date_time = datetime.datetime.strptime(AppTime, "%H:%M:%S")
            a_timedelta = date_time - datetime.datetime(1900, 1, 1)
            process_time[AppName] = a_timedelta.total_seconds()
            totaltime += a_timedelta.total_seconds()
            messagebox.showinfo(
                parent=self.master,
                title="App Added!",
                message="Application time for "
                + str(AppName)
                + " has been successfully added",
            )


def CalculatePercentages(seconds):
    return "%.0f%%" % (100 * seconds / totaltime)


def getActiveWindowTitle():
    try:
        return win32gui.GetWindowText(win32gui.GetForegroundWindow())
    except:
        return "Undefined App"


def taskDistribution(title):
    try:
        pid = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())
        activeapp = psutil.Process(pid[-1]).name().replace(".exe", "")
    except:
        return "Undefined App"

    AppDistribution = {
        "Zoom": "Meetings",
        "xlsm": "Macro Files",
        "Outlook": "Mail",
        "lync": "Skype",
    }

    res1 = [val for key, val in AppDistribution.items() if key in title]
    res2 = [val for key, val in AppDistribution.items() if key in activeapp]

    if res1:
        return " ".join([str(elem) for elem in res1])
    elif res2:
        return " ".join([str(elem) for elem in res2])
    else:
        return activeapp


def convertSeconds(seconds):
    h = round(seconds // (60 * 60))
    m = round((seconds - h * 60 * 60) // 60)
    s = round(seconds - (h * 60 * 60) - (m * 60))
    return str(h).zfill(2) + ":" + str(m).zfill(2) + ":" + str(s).zfill(2)


def sort_dictonary(process_time):
    return dict(sorted(process_time.items(), key=lambda x: x[1], reverse=True)).items()


def maincode(quote):
    global totaltime

    current_app = getActiveWindowTitle()
    current_app = taskDistribution(current_app)
    if len(current_app) > 20:
        current_app = current_app[:20] + "..."

    if current_app not in process_time.keys() and current_app.strip() != "":
        process_time[current_app] = 1
        totaltime += 1
    else:
        totaltime += 1
        process_time[current_app] = process_time[current_app] + 1
    for key, value in sort_dictonary(process_time):
        quote += f"{key}: {convertSeconds(value)} {CalculatePercentages(value)}\n"

    return quote


if __name__ == "__main__":
    process_time = {}
    timestamp = {}
    totaltime = 0
    quote = ""
    status = "Active"
    root = ThemedTk()
    TimeTrackerGUI = TimeTrackerGUI(root)
    root.mainloop()
