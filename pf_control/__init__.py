"""
#######################################################################################################################
###											Initialisation															###
###		Script deals with initialising the PF_Controlscripts														###
###																													###
###		Code developed by David Mills (david.mills@pscconsulting.com, +44 7899 984158) as part of PSC UK Ltd. 		###
###																													###
#######################################################################################################################
"""
import importlib
import sys

import pf_control
import pf_control.constants as constants
import pf_control.gui as gui

# Reload all modules
constants = importlib.reload(constants)
gui = importlib.reload(gui)
