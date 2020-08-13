"""
    Script to handle production of GUI for PF loader

"""
import tkinter as tk
import tkinter.filedialog
import tkinter.scrolledtext
import tkinter.messagebox as messagebox
import tkinter.ttk as ttk
import sys
import os
from tkinter import *

import pf_control
import pf_control.constants as constants

import subprocess

import sys
sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2020\python\3.8")
import powerfactory


class MainGui:
    """
        Main class to produce the GUI for user interaction
        Allows the user to set up the parameters and define the cases to run the studies

    """

    def __init__(self, pf_version, title=constants.GuiDefaults.gui_title):
        """
        Initialise GUI
        :param str pf_version: - This is the version of the running PowerFactory instance, if none then will populate
        a drop down list of available PowerFactory versions
        :param str title: (optional) - Title to be used for main window
        """
        self.logger = constants.logger

        # Initialise constants and Tk window
        tk.Tk.report_callback_exception = self.show_error
        self.master = tk.Tk()
        self.master.title(title)

        # General constants which needs to be initialised
        self._row = 0
        self._col = 0

        # Configure styles
        self.styles = CustomStyles()

        # Get a reference to all PowerFactory versions
        pf_constants = constants.PowerFactory()
        if pf_version:
            # If a PowerFactory version is already running then just display that version and disable the button
            dropdown_state = tk.DISABLED
        else:
            # Set the default value as the most recent version and enable dropdown
            pf_version = pf_constants.available_power_factory_versions[-1]
            dropdown_state = tk.NORMAL
        # Add a label and DropDown box to select the PowerFactory version
        _ = self.add_minor_label(
            row=self.row(1), col=self.col(), label='Select PowerFactory version:', columnspan=1,
            style=self.styles.label_general_left
        )
        self.selected_pf_version = self.add_dropdown_list(
            row=self.row(), col=self.col() + 1, values=pf_constants.available_power_factory_versions,
            def_value=pf_version, state=dropdown_state
        )

        # Add checkbox for each simulation module
        self.protection = self.add_checkbox(
            row=self.row(2), col=self.col(), text="Power Quality"
        )
        self.protection = self.add_checkbox(
            row=self.row(), col=self.col()+1, text="Contingency Analysis"
        )
        self.protection = self.add_checkbox(
            row=self.row(3), col=self.col(), text="Quasi-Dynamic Simulation"
        )
        self.protection = self.add_checkbox(
            row=self.row(), col=self.col()+1, text="Scripting and Automation"
        )
        self.protection = self.add_checkbox(
            row=self.row(4), col=self.col(), text="Stability Analysis"
        )
        self.protection = self.add_checkbox(
            row=self.row(), col=self.col()+1, text="Small Signal Stability"
        )
        self.protection = self.add_checkbox(
            row=self.row(5), col=self.col(), text="Network Reduction"
        )
        self.protection = self.add_checkbox(
            row=self.row(), col=self.col()+1, text="System Parameter Identification"
        )
        self.protection = self.add_checkbox(
            row=self.row(6), col=self.col(), text="Overcurrent Protection"
        )
        self.protection = self.add_checkbox(
            row=self.row(), col=self.col()+1, text="Arc-Flash Analysis"
        )

        self.master.mainloop()

    def row(self, i=0):
        """
            Returns the current row number + i
        :param int i: (optional=0) - Will return the current row number + this value
        :return int _row:
        """
        self._row += i
        return self._row

    def col(self, i=0):
        """
            Returns the current col number + i
        :param int i: (optional=0) - Will return the current col number + this value
        :return int _row:
        """
        self._col += i
        return self._col

    def add_minor_label(self, row, col, label, style, columnspan=2):
        """
            Function to add the name to the GUI
        :param row: Row number to use
        :param col: Column number to use
        :param str label: Label to use for header
        :param sty style: Style to use
        :param int columnspan:  (optional) - Number of columns to span
        :return ttk.Label lbl:  Reference to the newly created label
        """
        # Add label with the name to the GUI
        lbl = ttk.Label(self.master, text=label, style=style)
        lbl.grid(row=row, column=col, columnspan=columnspan, pady=5)
        return lbl

    def add_checkbox(self, row, col, text):
        """
            Function to add checkbox to the GUI
        :param row: Row number to use
        :param col: Column number to use
        :param str label: Label to use for header
        :param sty style: Style to use
        :param int columnspan:  (optional) - Number of columns to span
        :return ttk.Label lbl:  Reference to the newly created label
        """
        # Add label with the name to the GUI
        var = tk.IntVar(self.master)

        cbx = tk.Checkbutton(self.master, text=text, variable=var, selectcolor="white")
        cbx.grid(row=row, column=col, sticky=W, pady=5)
        return var

    def add_dropdown_list(self, row, col, values, def_value, state=tk.NORMAL):
        """
            General function just adds the list to the transformer rating (RPF)
        :param int row: Row number to use
        :param int col: Column number to use
        :param list values:  Values to populate dropdown box with
        :param str def_value:  Default value to initially populate box with
        :param str state:  Initial state of the dropdown list
        :return tk.StringVar variable:  Returns a reference to the DropDown box which contains the string variable
        """
        # Declare variable with initial default value
        variable = tk.StringVar(self.master)
        variable.set(def_value)

        # Create the drop down list to be shown in the GUI
        option_menu = ttk.OptionMenu(
            self.master, variable, def_value, *values, style=self.styles.option_menu
        )
        option_menu.grid(row=row, column=col, padx=6)
        option_menu.configure(state=state)
        return variable

    def show_error(self, *args):
        """
            Function to deal with error handling that occurs when running tkinter
        :param args:
        :return:
        """
        # Close all windows and exit Python
        self.master.destroy()
        self.logger.exception_handler(*args)

        # Close all tkinter windows
        sys.exit(1)


class CustomStyles:
    """ Class used to customize the layout of the GUI """

    def __init__(self):
        """
        Initialise the reference to the style names
        """
        # Constants for styles
        # Style for Loading the SAV Button
        self.cmd_buttons = 'General.TButton'
        self.option_menu = 'TMenubutton'

        self.label_general = 'TLabel'
        self.label_general_left = 'Left.TLabel'
        self.label_mainheading = 'MainHeading.TLabel'
        self.label_version_number = 'Version.TLabel'
        self.label_psc_info = 'PSCInfo.TLabel'
        self.label_psc_phone = 'PSCPhone.TLabel'
        self.label_hyperlink = 'Hyperlink.TLabel'

        self.configure_styles()

    def configure_styles(self):
        """
            Function configures all the ttk styles used within the GUI
            Further details here:  https://anzeljg.github.io/rin2/book2/2405/docs/tkinter/ttk-style-layer.html
            :return:
        """
        # Tidy up the repeat ttk.Style() calls
        # Switch to a different theme
        styles = ttk.Style()
        styles.theme_use('clam')

        # Configure the same font in all labels
        standard_font = constants.GuiDefaults.font_family
        bg_color = constants.GuiDefaults.color_main_window

        s = ttk.Style()
        s.configure('.', font=(standard_font, '8'))

        # General style for all buttons and active color changes
        s.configure(self.cmd_buttons, height=2, width=25)

        s.configure(self.option_menu, height=2, width=25)

        s.configure(self.label_general, background=bg_color)
        s.configure(self.label_general_left, background=bg_color, justify=tk.LEFT)
        s.configure(self.label_mainheading, font=(standard_font, '10', 'bold'), background=bg_color,
                    foreground=constants.GuiDefaults.font_heading_color)
        s.configure(self.label_version_number, font=(standard_font, '7'), background=bg_color, justify=tk.CENTER)
        s.configure(self.label_hyperlink, foreground='Blue', font=(standard_font, '7'), justify=tk.CENTER)

        s.configure(
            self.label_psc_info, font=constants.GuiDefaults.psc_font,
            color=constants.GuiDefaults.psc_color_web_blue, justify='center', background=bg_color
        )

        s.configure(
            self.label_psc_phone, font=(constants.GuiDefaults.psc_font, '8'),
            color=constants.GuiDefaults.psc_color_grey, background=bg_color
        )

        return None

    def command_button_color_change(self, color):
        """
            Force change in command button color to highlight error
        """
        s = ttk.Style()
        s.configure(self.cmd_buttons, background=color)

        return None


def running_in_powerfactory():
    """
        This function determines whether has been launched from PowerFactory or from a python terminal.
        If the former then will return the running PowerFactory version otherwise returns None
    :return str pf_version: Returns running PowerFactory version if applicable
    """

    # Determine if this script is being run from PowerFactory or plain Python.
    full_path_executable = sys.executable
    # Remove the folder path and keep only the executable file (in lower case).
    executable = os.path.basename(full_path_executable).lower()

    if executable in ['python.exe', 'pythonw.exe']:
        # If the executable was one of the above, it is a Python session.
        pf_version = None
    else:
        pf_version = os.path.basename(os.path.dirname(full_path_executable))

    return pf_version


if __name__ == '__main__':
    # Determine if running from PowerFactory and if so retrieve the current power factory version
    pf_version = running_in_powerfactory()

    main_gui = pf_control.gui.MainGui(pf_version=pf_version)

    # Get selected PowerFactory version
    selected_pf_version = main_gui.selected_pf_version.get()

    # Get selected settings
    protection_license = main_gui.protection.get()

    print("end of programme %s" % selected_pf_version)
    print("end of programme %d" % protection_license)

    # Find the installation directory for the PowerFactory paths
    DIG_PATH = os.path.join(constants.PowerFactory.default_install_directory, selected_pf_version)

    app = powerfactory.GetApplicationExt()
    user = app.GetCurrentUser()
    name = user.loc_name
    print('Current user is %s' % name)
    user.prot = protection_license

    subprocess.Popen(os.path.join(DIG_PATH, 'PowerFactory.exe'))

    print('PowerFactory Opened')