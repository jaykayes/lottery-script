import tkinter as tk
from tkinter import filedialog

import datetime as dt
from pathlib import Path
import sys
import os

import lottery

DEFAULT_BASE = Path("./lotteries/")
DEFAULT_INVENTORY = Path(DEFAULT_BASE, "inventory.xlsx")

APPLICATION_PERIOD = dt.timedelta(weeks = 2)

def make_deadline():
    '''Return datetime of 16:00 today

    Useful to use as a default suggestion for a deadline.
    '''
    dl = dt.datetime.today()
    dl = dl.replace(hour = 16, minute = 0, second = 0, microsecond = 0)

    return dl

def date_as_string(date):
    '''Returns given date as a string for displaying'''
    return date.strftime("%Y-%m-%d %H:%M")

def suggest_lotteryid():
    '''Gives a suggestion for the lottery identifier built from the current week number'''
    return dt.datetime.today().strftime("%Y-W%V")

def suggest_inventory() -> Path:
    '''Gives a suggestion for the default inventory file

    Checks if DEFAULT_INVENTORY exists and returns that if it does. If this
    file does not exist, does not suggest anything.
    '''
    inv = DEFAULT_INVENTORY

    if inv.exists():
        return inv
    else:
        return Path()

def suggest_applications() -> Path:
    '''Suggests applications file if one is found

    Checks if there is a file with the word "applications" (case-insensitive)
    in the lotteries/<lotteryid> directory. If so, this is returned as the
    applications file suggestion. Lottery id is gotten with
    suggest_lotteryid(). Otherwise an empty path is returned.
    '''
    lotterydir = suggest_resultsdir()

    if lotterydir.exists():
        ls = os.listdir(lotterydir)

        for filename in ls:
            if "applications" in filename.lower():
                return lotterydir / filename
    else:
        return Path()


def suggest_resultsdir() -> Path:
    '''Suggests a folder for the results based on the lottery id

    If the DEFAULT_BASE path exists, suggests a folder inside that based on the
    lottery id from suggest_lotteryid().
    '''
    if DEFAULT_BASE.exists():
        return Path(DEFAULT_BASE, suggest_lotteryid())
    else:
        return Path()

class SEGui(tk.Frame):
    # Lottery id
    def set_lotteryid(self, lotteryid):
        self.ent_lotteryid.delete(0, tk.END)
        self.ent_lotteryid.insert(0, lotteryid)

    def get_lotteryid(self) -> dt.datetime:
        return self.ent_lotteryid.get()

    # Inventory file
    def pick_inventory(self):
        file = filedialog.askopenfilename()

        if file:
            self.set_inventory(file)

    def set_inventory(self, inventory_file):
        self.ent_inventory.delete(0, tk.END)
        self.ent_inventory.insert(0, inventory_file)

    def get_inventory(self) -> Path:
        '''Returns the chosen inventory file'''
        return Path(self.ent_inventory.get())

    # Applications file
    def pick_applications(self):
        file = filedialog.askopenfilename()

        if file:
            self.set_applications(file)

    def set_applications(self, applications_file):
        self.ent_applications.delete(0, tk.END)
        self.ent_applications.insert(0, applications_file)

    def get_applications(self) -> Path:
        '''Returns the chosen applications file'''
        return Path(self.ent_applications.get())

    # Results folder
    def pick_resultsdir(self):
        file = filedialog.askdirectory()

        if file:
            self.set_resultsdir(file)

    def set_resultsdir(self, applications_file):
        self.ent_resultsdir.delete(0, tk.END)
        self.ent_resultsdir.insert(0, applications_file)

    def get_resultsdir(self) -> Path:
        '''Returns the chosen applications file'''
        return Path(self.ent_resultsdir.get())

    # Lottery opening date
    def set_opendate(self, opendate):
        self.ent_opendate.delete(0, tk.END)
        self.ent_opendate.insert(0, opendate)

    def get_opendate(self) -> dt.datetime:
        od = self.ent_opendate.get()
        return dt.datetime.strptime(od, '%Y-%m-%d %H:%M')

    # Application deadline
    def set_deadline(self, deadline):
        self.ent_deadline.delete(0, tk.END)
        self.ent_deadline.insert(0, deadline)

    def get_deadline(self) -> dt.datetime:
        dl = self.ent_deadline.get()
        return dt.datetime.strptime(dl, '%Y-%m-%d %H:%M')

    # Lottery process
    def run_lottery(self):
        '''Runs the Student Equipment lottery

        Calls the lottery function in the lottery module with the paramters
        provided in the GUI.
        '''
        print("SEGUI is running the lottery with following parameters:")

        inventory = self.get_inventory()
        applications = self.get_applications()
        resultsdir = self.get_resultsdir()

        results_filename = "{}_results.xls".format(self.get_lotteryid())
        results_file = resultsdir / results_filename

        opendate = self.get_opendate()
        deadline = self.get_deadline()

        print("{:>15}: {}".format("Inventory", inventory))
        print("{:>15}: {}".format("Applications", applications))
        print("{:>15}: {}".format("Results file", results_file))

        print("{:>15}: {}".format("Opening date", opendate))
        print("{:>15}: {}".format("Deadline", deadline))

        print("\nSEGUI is running the lottery!")

        lottery.lottery(inventory, applications, results_file, deadline, opendate)

    # GUI
    def build_window(self):
        '''Builds the GUI window

        Sets up all the GUI window widgets and callbacks.
        '''
        # Row counter 
        r = 0

        # Lottery open date
        self.lbl_lotteryid = tk.Label(self.master, text = "Lottery identifier")
        self.ent_lotteryid = tk.Entry(self.master)

        self.lbl_lotteryid.grid(column = 0, row = r, sticky = "WE")
        self.ent_lotteryid.grid(column = 1, row = r, sticky = "WE")

        r = r + 1

        # Inventory file chooser widgets
        self.btn_inventory = tk.Button(self.master,
                text    = "Select inventory file",
                command = self.pick_inventory)
        self.ent_inventory = tk.Entry(self.master)

        self.btn_inventory.grid(column = 0, row = r, sticky = "WE")
        self.ent_inventory.grid(column = 1, row = r, sticky = "WE")

        r = r + 1

        # Applications file chooser widgets
        self.btn_applications = tk.Button(self.master,
                text    = "Select applications file",
                command = self.pick_applications)
        self.ent_applications = tk.Entry(self.master)

        self.btn_applications.grid(column = 0, row = r, sticky = "WE")
        self.ent_applications.grid(column = 1, row = r, sticky = "WE")

        r = r + 1

        # Output chooser widgets
        self.btn_resultsdir = tk.Button(self.master,
                text    = "Select folder for results",
                command = self.pick_resultsdir)
        self.ent_resultsdir = tk.Entry(self.master)

        self.btn_resultsdir.grid(column = 0, row = r, sticky = "WE")
        self.ent_resultsdir.grid(column = 1, row = r, sticky = "WE")

        r = r + 1

        # Lottery open date
        self.lbl_opendate = tk.Label(self.master, text = "Lottery open date")
        self.ent_opendate = tk.Entry(self.master)

        self.lbl_opendate.grid(column = 0, row = r, sticky = "WE")
        self.ent_opendate.grid(column = 1, row = r, sticky = "WE")

        r = r + 1

        # Lottery deadline date
        self.lbl_deadline = tk.Label(self.master, text = "Lottery deadline")
        self.ent_deadline = tk.Entry(self.master)

        self.lbl_deadline.grid(column = 0, row = r, sticky = "WE")
        self.ent_deadline.grid(column = 1, row = r, sticky = "WE")

        r = r + 1

        # Run lottery button
        self.btn_run = tk.Button(self.master,
                text    = "Run lottery!!!!! :D",
                command = self.run_lottery)

        self.btn_run.grid(columnspan = 2, row = r, sticky = "WE")

        r = r + 1

    def __init__(self, master = None):
        tk.Frame.__init__(self, master, bd=10)
        self.master = master

        self.build_window()

        self.master.columnconfigure(1, weight=1)

        deadline = make_deadline()
        opendate = deadline - APPLICATION_PERIOD

        self.set_lotteryid(suggest_lotteryid())
        self.set_inventory(suggest_inventory())
        self.set_applications(suggest_applications())
        self.set_resultsdir(suggest_resultsdir())
        self.set_opendate(date_as_string(opendate))
        self.set_deadline(date_as_string(deadline))

if __name__ == "__main__":
    root = tk.Tk()
    root.wm_title("Student Equipment Lottery")

    win = SEGui(root)

    root.mainloop()
