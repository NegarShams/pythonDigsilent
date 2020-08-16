"""
#######################################################################################################################
###											Constants																###
###		Central point to store all constants associated with PSC harmonics											###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###																													###
#######################################################################################################################
"""

# import pandas as pd
# import numpy as np
import os
import sys
import glob
import time

# Meta Data
__author__ = 'David Mills'
__version__ = '1.0.0'
__email__ = 'david.mills@pscconsulting.com'
__phone__ = '+44 7899 984158'
__status__ = 'In Development - Beta'

# Label used when displaying messages
__title__ = 'PowerFactory Loader'

logger_name = 'PSC'
logger = None

DEBUG = True

# Reference to local directory used by other packages
local_directory = os.path.abspath(os.path.dirname(__file__))


class PowerFactory:
    """
    Constants used in this script
    """
    # Populated with available installed PowerFactory versions on initialisation
    available_power_factory_versions = list()
    target_power_factory = 'PowerFactory 2019'

    # Default PowerFactory installation directories
    default_install_directory = r'C:\Program Files\DIgSILENT'
    power_factory_search = 'PowerFactory 20*'

    power_factory_host = 'digsilent2'

    def __init__(self):
        """
            Initialises the relevant python paths depending on the version that has been loaded
        """

        # Find all PowerFactory versions installed in this location
        power_factory_paths = glob.glob(os.path.join(self.default_install_directory, self.power_factory_search))
        self.available_power_factory_versions = [os.path.basename(x) for x in power_factory_paths]
        self.available_power_factory_versions.sort()


class GuiDefaults:
    gui_title = 'PSC - PowerFactory Loader'

    # Default labels for buttons (only those which get changed during running)
    button_select_settings_label = 'Confirm Selection'
    button_launch_powerfactory_label = 'Launch PowerFactory'

    # Default extensions used in file type selection windows
    xlsx_types = (('xlsx files', '*.xlsx'), ('All Files', '*.*'))

    font_family = 'Helvetica'
    psc_uk = 'PSC UK'
    psc_phone = '\nPSC UK:  +44 1926 675 851'
    psc_font = ('Calibri', '10', 'bold')
    psc_color_web_blue = '#%02x%02x%02x' % (43, 112, 170)
    psc_color_grey = '#%02x%02x%02x' % (89, 89, 89)
    font_heading_color = '#%02x%02x%02x' % (0, 0, 255)
    img_size_psc = (120, 120)

    # TODO: Test logos exist
    img_pth_psc_main = os.path.join(local_directory, 'PSC Logo RGB Vertical.png')
    img_pth_psc_window = os.path.join(local_directory, 'PSC Logo no tag-1200.gif')

    # TODO: Test hyperlink works
    hyperlink_psc_website = 'https://www.pscconsulting.com/'

    # Colors
    color_main_window = '#%02x%02x%02x' % (239, 243, 241)
    # Color of pop-up windows which may be different to main window
    color_pop_up_window = color_main_window
    error_color = '#%02x%02x%02x' % (255, 32, 32)

    # Reference to the Tkinter binding of a mouse button
    mouse_button_1 = '<Button - 1>'
