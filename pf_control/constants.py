"""
#######################################################################################################################
###											Constants																###
###		Central point to store all constants associated with PF Loader  											###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###																													###
#######################################################################################################################
"""

import os
import sys
import glob

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
    pf_year = 2020
    year_max_tested = 2020
    pf_service_pack = ''
    dig_path = str()
    dig_python_path = str()
    # Populated with available installed PowerFactory versions on initialisation
    available_power_factory_versions = list()
    target_power_factory = 'PowerFactory 2020'

    # The following list details python versions which are non compatible
    non_compatible_python_versions = ['3.5']

    # Default PowerFactory installation directories
    default_install_directory = r'C:\Program Files\DIgSILENT'
    power_factory_search = 'PowerFactory 20*'

    power_factory_host = 'digsilent2'

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
                self.target_power_factory = self.available_power_factory_versions[0]
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
