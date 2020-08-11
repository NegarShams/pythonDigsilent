"""
#######################################################################################################################
###											PF Control																###
###		This script deals with loading PowerFactory with the relevant license modules pre-selected					###
###																													###
###		Code developed by:																							###
###			David Mills (david.mills@PSCconsulting.com, +44 7899 984158)											###
###																													###
#######################################################################################################################

TODO - GUI to request the following details:
TODO 1. user_name (is it possible to get the default user_name from powerfactory)
TODO 2. pf_version
TODO 3. tick boxes for licenses to activate

TODO - Script will also need to have functions to:
TODO 1. Search PC to find installed PF versions (detect versions which are not compatible with the running version of Python)
TODO 2. Store previous values and directories for faster loading
"""



import os
import sys
import subprocess

user_name = 'HoomanAndami'
protection = 0
harmonics = 1
arcflash = 0

pf_version = '2018'

DIG_PATH_2016 = r'C:\Program Files\DIgSILENT\PowerFactory 2016 SP5'
DIG_PATH_2018 = r'C:\Program Files\DIgSILENT\PowerFactory 2018 SP5'
DIG_PYTHON_PATH_2016 = r'C:\Program Files\DIgSILENT\PowerFactory 2016 SP5\Python\3.5'
DIG_PYTHON_PATH_2018 = r'C:\Program Files\DIgSILENT\PowerFactory 2018 SP5\Python\3.5'

if pf_version == '2016':
	DIG_PATH = DIG_PATH_2016
	DIG_PYTHON_PATH = DIG_PYTHON_PATH_2016
elif pf_version == '2018':
	DIG_PATH = DIG_PATH_2018
	DIG_PYTHON_PATH = DIG_PYTHON_PATH_2018
else:
	print('ERROR python version not found')
	raise SyntaxError('ERROR')

sys.path.append(DIG_PATH)
# #sys.path.append(DIG_PATH_2018)
sys.path.append(DIG_PYTHON_PATH)
# #sys.path.append(DIG_PYTHON_PATH_2018)

os.environ['PATH'] = os.environ['PATH'] + ';' + DIG_PATH

import powerfactory

# Load application as Administrator
# #app = powerfactory.GetApplication(username='Administrator', password='Administrator')
# No need to load as Administrator when controlling through Python
app = powerfactory.GetApplication()

# Get list of all users
users = app.GetAllUsers()

# Loop through until found this user
for user in users:
	if user.loc_name == user_name:
		user.prot = protection
		user.harm = harmonics
		user.arcflash = arcflash
		break

app = None

# #app.Show()
subprocess.Popen(os.path.join(DIG_PATH, 'PowerFactory.exe'))

print('PowerFactory Opened')