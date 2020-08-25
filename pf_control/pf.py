import os
import sys
import pf_control.constants as constants

# powerfactory will be defined after initialisation by the PowerFactory class
powerfactory = None
app_pf = None


class PowerFactory:
    """
        Class to deal with system level interfacing in PowerFactory
    """

    def __init__(self):
        """ Gets the relevant powerfactory version and initialises """
        # Get reference to the constants and carry out a search for all of the available PowerFactory versions
        self.c = constants.PowerFactory()
        self.logger = constants.logger

    def add_python_paths(self):
        """
            Function retrieves the relevant python paths, adds them and then imports the powerfactory module
            Importing of the powerfactory module has to happen here due to the location
        """
        self.logger.debug('Searching of paths to PowerFactory and adding the Python search path')
        # Get the python paths if not already populated
        if not (self.c.dig_path and self.c.dig_python_path):
            # Initialise so that the paths are looked for and the
            self.c = self.c()

        # Add the paths to system and the environment and then try and import powerfactory
        sys.path.append(self.c.dig_path)
        sys.path.append(self.c.dig_python_path)
        os.environ['PATH'] = '{};{}'.format(os.environ['PATH'], self.c.dig_path)

        self.logger.debug(
            'The following paths have been added to the Python modules search path:\n\t{}\n\t{}'.format(
                self.c.dig_path, self.c.dig_python_path
            )
        )

        # Try and import the powerfactory module
        try:
            self.logger.debug('Importing the powerfactory module')
            global powerfactory
            import powerfactory
            self.logger.debug('Imported successfully')
        except ImportError:
            self.logger.critical(
                (
                    'It has not been possible to import the powerfactory module and therefore the script cannot be run.\n'
                    'This is most likely due to there not being a powerfactory.pyc file located in the following path:\n\t'
                    '<{}>\n'
                    'Please check this exists and the error messages above.'
                ).format(self.c.dig_python_path)
            )
            raise ImportError('PowerFactory module not found')

        return None

    def initialise_power_factory(self, pf_version=None):
        """
            Function initialises powerfactory and provides an object reference to it
        :param str pf_version:  Will initialise power_factory based on the version provided
        :return None:
        """
        # Check if already running from PowerFactory and if so then update to use that power factory version

        # Check the paths have already been found and if not call the relevant function
        if not (self.c.dig_path and self.c.dig_python_path):
            # Initialise so that the paths are looked for with the provided pf_version
            self.c.select_power_factory_version(pf_version=pf_version)
            self.add_python_paths()

        # Different APIs exist for different PowerFactory versions, if an old version is run then different
        # initialisation route.  When initialising need to warn user that old version is being used
        global app_pf
        app_pf = powerfactory.GetApplication()

        return app_pf
