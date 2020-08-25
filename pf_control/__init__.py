"""
#######################################################################################################################
###											Initialisation															###
###		Script deals with initialising the PF_Control scripts														###
###																													###
###		Code developed by Jinsheng Peng (jinsheng.peng@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###																													###
#######################################################################################################################
"""
import importlib
import logging

import pf_control.constants as constants
import pf_control.gui as gui
import pf_control.pf as pf

# Reload all modules
constants = importlib.reload(constants)
gui = importlib.reload(gui)

if constants.logger is None:
	logging.basicConfig()
	constants.logger = logging.getLogger()
	constants.logger.setLevel(level=logging.DEBUG)
