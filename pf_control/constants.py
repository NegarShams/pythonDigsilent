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
    # Constants relating to the paths
    pf_year = 2019
    year_max_tested = 2019
    pf_service_pack = ''
    dig_path = str()
    dig_python_path = str()
    # Populated with available installed PowerFactory versions on initialisation
    available_power_factory_versions = list()
    target_power_factory = 'PowerFactory 2019'

    # The following list details python versions which are non compatible
    non_compatible_python_versions = ['3.5']

    # Default PowerFactory installation directories
    default_install_directory = r'C:\Program Files\DIgSILENT'
    power_factory_search = 'PowerFactory 20*'

    # Constants associated with the handling of PowerFactory initialisation and
    # potential intermittent errors
    # Number of attempts to obtain a license
    license_activation_attempts = 5
    # Number of seconds to wait between license attempts
    license_activation_delay = 5.0
    # Error codes which could be intermittent and therefore the script should try again
    # Description in PowerFactory help file:  ErrorCodeReference_en.pdf
    license_activation_error_codes = (3000, 3002, 3005, 3011, 3012, 4000, 4002, 5000)

    # User default settings
    user_default_settings = 'Set\Def\Settings.SetUser'

    class ComRes:
        # Power Factory class name
        pf_comres = 'ComRes'
        # Com Res setting constants

        # File export type:
        #	0 = Output window
        #	1 = Windows clipboard
        #	2 = Measurement file (ElmFile)
        #	3 = Comtrade
        #	4 = Testfile
        #	5 = PSSPLT Version 2.0
        #	6 = Commas Separated Values (*.csv)
        export_type = 'iopt_exp'
        # Name of file to export to (if appropriate)
        file = 'f_name'
        # Type of separators to use (0 = Custom, 1 = system defaults)
        separators = 'iopt_sep'
        # Export object headers only (0 = all data, 1 = headers only)
        object_head_only = 'iopt_honly'
        # Variables to extract (0 = all, 1 = custom list)
        variables_all = 'iopt_csel'
        # Name of result file from PF to export
        result = 'pResult'
        # Details to export from element:
        # 	0 = None,
        # 	1 = Name,
        # 	2 = Short path and name,
        # 	3 = Path and name,
        # 	4 = Foreign key
        element = 'iopt_locn'
        # Details to export from variable:
        #	0 = None,
        #	1 = Parameters name,
        #	3 = Short description,
        #	4 = Full description
        variable = 'ciopt_head'
        # Custom of full dataset (0 = full, 1 = custom)
        user_interval = 'iopt_tsel'
        # Export values (0 = values, 1 = variable descriptors only)
        export_values = 'iopt_vars'
        # Shift time of results (0 = none, 1 = shift)
        shift_time = 'iopt_rscl'
        # Filter time of results (0 = None, 1 = filter)
        filtered_time = 'filtered'

    def __init__(self):
        """
            Initialises the relevant python paths depending on the version that has been loaded
        """
        # Get reference to logger
        self.logger = logger

        # Find all PowerFactory versions installed in this location
        power_factory_paths = glob.glob(os.path.join(self.default_install_directory, self.power_factory_search))
        self.available_power_factory_versions = [os.path.basename(x) for x in power_factory_paths]
        self.available_power_factory_versions.sort()

    def select_power_factory_version(self, pf_version=None, mock_python_version=str()):
        """
            Function allows the user to select a specific PowerFactory version, if none is selected then
            the most recent version of PowerFactory is used
        :param str pf_version: (optional) - If None then the most recent PowerFactory version is used
        :param str mock_python_version:  For TESTING only, gets replaced with a different version to check correct
                                        errors thrown if incorrect version provided

        """

        # If no pf_version is provided then the default version defined is used if it exists in the avaiable versions
        # otherwise the latest version
        if pf_version is None:
            if self.target_power_factory not in self.available_power_factory_versions:
                # Rather than assuming a particular version just default to the latest version
                self.target_power_factory = self.available_power_factory_versions[-1]
        elif pf_version in self.available_power_factory_versions:
            self.target_power_factory = pf_version
        else:
            self.logger.critical(
                (
                    'The PowerFactory version {} has been selected but does not exist in the installed versions:\n\t{}'
                ).format(pf_version, '\n\t'.join([x for x in self.available_power_factory_versions]))
            )
            raise EnvironmentError('Invalid PowerFactory version')

        self.logger.debug('PowerFactory version <{}> will be used'.format(self.target_power_factory))

        # Find year from selected PowerFactory version
        # pf_version is assumed to take the format PowerFactory #### and therefore the #### can be extracted
        year = [int(s) for s in self.target_power_factory.split() if s.isdigit()][0]

        # Confirm the year is > 2017 and < 2020 otherwise warn that hasn't been fully tested
        if int(year) < 2018:
            self.logger.warning(
                (
                    'You are using PowerFactory version {}.\n'
                    'In the 2018 version there were some API changes which have been considered in this script.  The '
                    'previous versions may still work but they have not been considered as part of the development '
                    'testing and so you are advised to carefully check your results.'
                ).format(year)
            )
        elif int(year) > self.year_max_tested:
            self.logger.warning(
                (
                    'You are using PowerFactory version {}.\n'
                    'This script has only been tested up to year {} and therefore changes in the PowerFactory API may '
                    'impact on the running and results you produce.  You are advised to check the results carefully or '
                    'consider updating the developments testing for this version.  For further details contact:\n{}'
                ).format(year, self.year_max_tested, Author.contact_summary)
            )

        # Find the installation directory for the PowerFactory paths
        self.dig_path = os.path.join(self.default_install_directory, self.target_power_factory)

        # Now checks for Python versions within this PowerFactory
        if mock_python_version:
            # Running in a test mode to check an error is created
            self.logger.warning(
                'TESTING - Testing with a mock python version to raise exception, if not expected then there is a '
                'script error! - TESTING')
            current_python_version = mock_python_version
        else:
            # Formulate string for python version
            current_python_version = '{}.{}'.format(sys.version_info.major, sys.version_info.minor)

        # Get list of supported python versions
        list_of_available_versions = [os.path.basename(x) for x in
                                      glob.glob(os.path.join(self.dig_path, 'Python', '*'))]
        if current_python_version in self.non_compatible_python_versions:
            self.logger.critical(
                (
                    'You are using Python version {}, this script is not compatible with that version or the following '
                    'versions: \n\t Python {}\n  Additionally, the PowerFactory version you have selected ({}) is only compatible '
                    'with the following Python versions: \n\t Python {}'
                ).format(
                    current_python_version, '\n\t Python '.join(self.non_compatible_python_versions),
                    self.target_power_factory, '\n\t Python '.join(list_of_available_versions)
                )
            )
            raise EnvironmentError('Non Compatible Python Version')

        # Define the Python Path for PowerFactory
        self.dig_python_path = os.path.join(self.dig_path, 'Python', current_python_version)
        if not os.path.isdir(self.dig_python_path):
            self.logger.critical(
                (
                    'You are running python version: {} but only the following versions are supported by this version of'
                    'PowerFactory ({}):\n\t Python {}'
                ).format(current_python_version, self.target_power_factory,
                         '\n\t Python '.join(list_of_available_versions))
            )
            raise EnvironmentError('Incompatible Python version')


class GuiDefaults:
    gui_title = 'PSC - PowerFactory Loader'

    # Default labels for buttons (only those which get changed during running)
    button_select_settings_label = 'Select Settings File'

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
