"""
    Script to handle production of GUI for PF loader

"""
import tkinter as tk
import tkinter.messagebox as messagebox
import tkinter.ttk as ttk
import os
from tkinter import *

from PIL import Image, ImageTk

import pf_control
import pf_control.constants as constants
import webbrowser

import subprocess
import platform

import sys


class MainGui:
    """
        Main class to produce the GUI for PowerFactory Loader

    """

    def __init__(self, title=constants.GuiDefaults.gui_title):
        """
        Initialise GUI
        :param str title: (optional) - Title to be used for main window
        """
        self.pf = pf_control.pf.PowerFactory()
        self.logger = constants.logger

        self.power_factory_launch_button = 0

        self.pf_version = None

        # Initialise constants and Tk window
        tk.Tk.report_callback_exception = self.show_error
        self.master = tk.Tk()
        self.master.title(title)

        # Change color of main window
        self.master.configure(bg=constants.GuiDefaults.color_main_window)

        # Ensure that on_closing is processed correctly
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

        # General constants which needs to be initialised
        self._row = 0
        self._col = 0

        # Configure styles
        self.styles = CustomStyles()

        # Get a reference to all PowerFactory versions
        self.c = constants.PowerFactory()

        # Set the default value as the most recent version and enable dropdown
        default_pf_version = self.c.available_power_factory_versions[-1]
        dropdown_state = tk.NORMAL

        # Add a label and DropDown box to select the PowerFactory version
        _ = self.add_minor_label(
            row=self.row(), col=self.col(), label='Select PowerFactory Version:', columnspan=1,
            style=self.styles.label_general_left
        )
        self.selected_pf_version = self.add_dropdown_list(
            row=self.row(), col=self.col() + 1, values=self.c.available_power_factory_versions,
            def_value=default_pf_version, state=dropdown_state
        )

        # Get selected PowerFactory version and Define the powerfactory application path
        #self.selected_pf_version_get = self.selected_pf_version.get()

        # Add checkbox for each simulation module
        self.power_quality = self.add_checkbox(
            row=self.row(1), col=self.col(), text="Power Quality"
        )
        self.contingency = self.add_checkbox(
            row=self.row(), col=self.col() + 1, text="Contingency Analysis"
        )
        self.quasi_dynamic = self.add_checkbox(
            row=self.row(1), col=self.col(), text="Quasi-Dynamic Simulation"
        )
        self.scripting = self.add_checkbox(
            row=self.row(), col=self.col() + 1, text="Scripting and Automation"
        )
        self.stability = self.add_checkbox(
            row=self.row(1), col=self.col(), text="Stability Analysis"
        )
        self.small_signal = self.add_checkbox(
            row=self.row(), col=self.col() + 1, text="Small Signal Stability"
        )
        self.network_reduction = self.add_checkbox(
            row=self.row(1), col=self.col(), text="Network Reduction"
        )
        self.parameter_identification = self.add_checkbox(
            row=self.row(), col=self.col() + 1, text="System Parameter Identification"
        )
        self.overcurrent_protection = self.add_checkbox(
            row=self.row(1), col=self.col(), text="Overcurrent Protection"
        )
        self.arc_flash = self.add_checkbox(
            row=self.row(), col=self.col() + 1, text="Arc-Flash Analysis"
        )

        # Add button for user to confirm selection and open PF in engine mode to change licence settings
        self.button_confirm_settings = self.add_cmd(
            label=constants.GuiDefaults.button_select_settings_label,
            cmd=self.change_licence_settings, tooltip='Click to confirm the selected settings',
            row=self.row(1), col=self.col()
        )

        # Add button for user to launch PF
        self.button_launch_powerfactory = self.add_cmd(
            label=constants.GuiDefaults.button_launch_powerfactory_label,
            cmd=self.launch_powerfactory, tooltip='Click to launch PowerFactory', state=tk.DISABLED,
            row=self.row(), col=self.col()+1
        )

        # Separator
        self.add_sep(row=self.row(1), col_span=2)

        _ = self.add_instruction_label(
            row=self.row(1), col=self.col(), label='Action and Status:', columnspan=1,
            style=self.styles.label_general_left
        )
        self.status_bar = self.add_status_label(
            row=self.row(1), col=self.col(), label='Please select PF version and licences, and confirm selection.', columnspan=2,
            style=self.styles.label_general_left
        )

        _ = self.add_note_label(
            row=self.row(1), col=self.col(), label='Note: Base package licences will be enabled by default.', columnspan=2,
            style=self.styles.label_general_left
        )
        # Add PSC logo in Windows Manager
        self.add_psc_logo_wm()

        # Add PSC logo with hyperlink to the website
        self.add_logo(
            row=self.row(1), col=self.col(),
            img_pth=constants.GuiDefaults.img_pth_psc_main,
            hyperlink="None",
            size=constants.GuiDefaults.img_size_psc
        )

        self.master.mainloop()

    def change_licence_settings(self):

        self.status_bar.configure(text="Check VPN connection")
        self.master.update()
        response = self.ping(constants.PowerFactory.power_factory_host)
        # and then check the response, if no VPN connection, indicate this in the GUI status output
        if response == 0:

            # Open PF in engine mode and get current user
            self.status_bar.configure(text="Setting up PowerFactory licence (take around 10 seconds)")
            self.master.update()

            self.pf_version = self.selected_pf_version.get()
            app = self.pf.initialise_power_factory(pf_version=self.pf_version)
            user = app.GetCurrentUser()
            # Set selected license
            user.check_adv = 0
            user.harm = self.power_quality.get()
            user.contingency = self.contingency.get()
            user.qdynsim = self.quasi_dynamic.get()
            user.script = self.scripting.get()
            user.stab = self.stability.get()
            user.smallsig = self.small_signal.get()
            user.netred = self.network_reduction.get()
            user.paramid = self.parameter_identification.get()
            user.prot = self.overcurrent_protection.get()
            user.arcflash = self.arc_flash.get()

            # Enable PowerFactory Launching Button
            self.status_bar.configure(text="Licence selection updated, click Launch PowerFactory")
            self.master.update()
            self.button_launch_powerfactory.configure(state=tk.NORMAL)
        else:
            self.status_bar.configure(text="No VPN Connection")

    def ping(self, host):
        """
        Returns True if host (str) responds to a ping request.
        Remember that a host may not respond to a ping (ICMP) request even if the host name is valid.
        """

        # Option for the number of packets as a function of
        param = '-n' if platform.system().lower() == 'windows' else '-c'

        # Building the command. Ex: "ping -c 1 google.com"
        command = ['ping', param, '1', host]

        return subprocess.call(command)

    def launch_powerfactory(self):

        self.power_factory_launch_button = 1

        self.master.destroy()

    def add_psc_logo_wm(self):
        """
            Function just adds the PSC logo to the windows manager in GUI
        :return: None
        """
        # Create the PSC logo for including in the windows manager
        logo = tk.PhotoImage(file=constants.GuiDefaults.img_pth_psc_window)
        # noinspection PyProtectedMember
        self.master.tk.call('wm', 'iconphoto', self.master._w, logo)
        return None

    def add_logo(self, row, col, img_pth, hyperlink=None, size=constants.GuiDefaults.img_size_psc):
        """
            Function to add an image which when clicked is a hyperlink to the companies logo.
            Image is added using a label and changing the it to be an image and binding a hyperlink to it
        :param int row:  Row number to use
        :param int col:  Column number to use
        :param str img_pth:  Path to image to use
        :param str hyperlink:  (optional=None) Website hyperlink to use
        :param str tooltip:  (Optional=None) Popup message to use for mouse over
        :param tuple size: (optional) - Size to make image when inserting
        :return ttk.Label logo:  Reference to the newly created logo
        """
        # Load the image and convert into a suitable size for displaying on the GUI
        img = Image.open(img_pth)
        img.thumbnail(size)
        # Convert to a photo image for inclusion on the GUI
        img_to_include = ImageTk.PhotoImage(img)

        # Add image to GUI
        logo = tk.Label(self.master, image=img_to_include,
                        bg='white')
        logo.photo = img_to_include
        logo.grid(row=row, column=col, columnspan=2, pady=10)

        # Associate clicking of the button as opening a web browser with the provided hyperlink
        if hyperlink:
            logo.bind(
                constants.GuiDefaults.mouse_button_1,
                lambda e: webbrowser.open_new(hyperlink)
            )

        return logo

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

    def add_sep(self, row, col_span):
        """
            Function just adds a horizontal separator
        :param row: Row number to use
        :param col_span: Column span number to use
        :return None:
        """
        # Add separator
        sep = ttk.Separator(self.master, orient="horizontal")
        sep.grid(row=row, sticky=tk.W + tk.E, columnspan=col_span, pady=10)
        return None

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
        lbl = ttk.Label(self.master, text=label, style=style, font=(constants.GuiDefaults.font_family, 9))
        lbl.grid(row=row, column=col, columnspan=columnspan, pady=5, sticky=W)
        return lbl

    def add_instruction_label(self, row, col, label, style, columnspan=2):
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
        lbl = ttk.Label(self.master, text=label, style=style, font=(constants.GuiDefaults.font_family, 10))
        lbl.grid(row=row, column=col, columnspan=columnspan, pady=5, sticky=W)
        return lbl

    def add_status_label(self, row, col, label, style, columnspan=2):
        """
            Function to add the status to the GUI
        :param row: Row number to use
        :param col: Column number to use
        :param str label: Label to use for header
        :param sty style: Style to use
        :param int columnspan:  (optional) - Number of columns to span
        :return ttk.Label lbl:  Reference to the newly created label
        """
        # Add label with the name to the GUI
        lbl = ttk.Label(self.master, text=label, style=style, font=(constants.GuiDefaults.font_family, 9),
                        background="white", foreground="blue")
        lbl.grid(row=row, column=col, columnspan=columnspan, pady=5)
        return lbl

    def add_note_label(self, row, col, label, style, columnspan=2):
        """
            Function to add the status to the GUI
        :param row: Row number to use
        :param col: Column number to use
        :param str label: Label to use for header
        :param sty style: Style to use
        :param int columnspan:  (optional) - Number of columns to span
        :return ttk.Label lbl:  Reference to the newly created label
        """
        # Add label with the name to the GUI
        lbl = ttk.Label(self.master, text=label, style=style, font=(constants.GuiDefaults.font_family, 10))
        lbl.grid(row=row, column=col, columnspan=columnspan, pady=10, sticky=W)
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

    def add_cmd(self, label, cmd, state=tk.NORMAL, tooltip=str(), row=None, col=None):
        """
            Function just adds the command button to the GUI which is used for loading the SAV case
        :param int row: (optional) Row number to use
        :param int col: (optional) Column number to use
        :param str label:  Label to use for button
        :param func cmd: Command to use when button is clicked
        :param int state:  Tkinter state for button initially
        :param str tooltip:  Message that pops up if hover over button
        :return None:
        """
        # If no number is provided for row or column then assume to add 1 to row and 0 to column
        if not row:
            row = self.row(1)
        if not col:
            col = self.col()

        button = ttk.Button(
            self.master, text=label, command=cmd, style=self.styles.cmd_buttons, state=state)
        button.grid(row=row, column=col, padx=5, pady=5, sticky=tk.W + tk.E)
        CreateToolTip(widget=button, text=tooltip)

        return button

    def on_closing(self):
        """
            Function runs when window is closed to determine if user actually wants to cancel running of study
        :return None:
        """
        # Ask user to confirm that they actually want to close the window
        result = messagebox.askquestion(
            title='Exit PowerFactory Loader?',
            message='Are you sure you want to close the loader'
        )

        # Test what option the user provided
        if result == 'yes':
            self.abort = True
            self.master.destroy()
            return None
        else:
            return None

    def show_error(self, *args):
        """
            Function to deal with error handling that occurs when running tkinter
        :param args:
        :return:
        """
        # Close all windows and exit Python
        self.master.destroy()
        # self.logger.exception_handler(*args)

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


class CreateToolTip(object):
    """
		Function to create a popup tool tip for a given widget based on the descriptions provided here:
			https://stackoverflow.com/questions/3221956/how-do-i-display-tooltips-in-tkinter
	"""

    def __init__(self, widget, text="widget info"):
        """
			Establish link with tooltip
		:param widget: Tkinter element that tooltip should be associated with
		:param text:    Message to display when hovering over button
		"""
        self.wait_time = 500  # milliseconds
        self.wrap_length = 450  # pixels
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None

    def enter(self, event=None):
        del event
        self.schedule()

    def leave(self, event=None):
        del event
        self.unschedule()
        self.hidetip()

    def schedule(self, event=None):
        del event
        self.unschedule()
        self.id = self.widget.after(self.wait_time, self.showtip)

    def unschedule(self, event=None):
        del event
        _id = self.id
        self.id = None
        if _id:
            self.widget.after_cancel(_id)

    def showtip(self):
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        # creates a top level window
        self.tw = tk.Toplevel(self.widget)
        self.tw.attributes('-topmost', 'true')
        # Leaves only the label and removes the app window
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(
            self.tw, text=self.text, justify='left', background="#ffffff", relief='solid', borderwidth=1,
            wraplength=self.wrap_length
        )
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tw
        self.tw = None
        if tw:
            tw.destroy()


if __name__ == '__main__':

    main_gui = pf_control.gui.MainGui()

    selected_dig_path = os.path.join(constants.PowerFactory.default_install_directory, main_gui.pf_version)

    if main_gui.power_factory_launch_button == 1:
        subprocess.Popen(os.path.join(selected_dig_path, 'PowerFactory.exe'))
