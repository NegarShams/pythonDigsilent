"""
#######################################################################################################################
###											PF Control																###
###		This script deals with loading PowerFactory with the relevant license modules pre-selected					###
###																													###
###		Code developed by:																							###
###			David Mills (david.mills@PSCconsulting.com, +44 7899 984158)											###
###																													###
#######################################################################################################################

TODO - Build GUI to request the following details:
TODO 1. user_name (is it possible to get the default user_name from powerfactory)
TODO 2. drop down selector for power factory version (see JA7896 project:  pscharmonics.gui.MainGui lines 449-466)
TODO 2.1 - JA7896 project:  pscharmonics.pf.PowerFactory has a search routine for finding all installed versions
TODO 3. tick boxes for licenses to activate
TODO 4. Storing previous details of installed PowerFactory versions to avoid searching every time
TODO:    (only need to do this if actually takes a long time)


TODO: The code to actually change the selected license, etc. also needs to be written into functions
"""



import os
import sys
import subprocess

user_name = 'NegarShams'
protection = 0
harmonics = 1
arcflash = 0

pf_version = '2020'

DIG_PATH_2016 = r'C:\Program Files\DIgSILENT\PowerFactory 2016 SP5'
DIG_PATH_2018 = r'C:\Program Files\DIgSILENT\PowerFactory 2018 SP5'
DIG_PATH_2019 = r'C:\Program Files\DIgSILENT\PowerFactory 2019'
DIG_PATH_2020 = r'C:\Program Files\DIgSILENT\PowerFactory 2020'
DIG_PYTHON_PATH_2016 = r'C:\Program Files\DIgSILENT\PowerFactory 2016 SP5\Python\3.5'
DIG_PYTHON_PATH_2018 = r'C:\Program Files\DIgSILENT\PowerFactory 2018 SP7\Python\3.5'
DIG_PYTHON_PATH_2019 = r'C:\Program Files\DIgSILENT\PowerFactory 2019\Python\3.5'
DIG_PYTHON_PATH_2020 = r'C:\Program Files\DIgSILENT\PowerFactory 2020\Python\3.8'

if pf_version == '2016':
	DIG_PATH = DIG_PATH_2016
	DIG_PYTHON_PATH = DIG_PYTHON_PATH_2016
elif pf_version == '2018':
	DIG_PATH = DIG_PATH_2018
	DIG_PYTHON_PATH = DIG_PYTHON_PATH_2018
elif pf_version == '2019':
	DIG_PATH = DIG_PATH_2019
	DIG_PYTHON_PATH = DIG_PYTHON_PATH_2019
elif pf_version == '2020':
	DIG_PATH = DIG_PATH_2020
	DIG_PYTHON_PATH = DIG_PYTHON_PATH_2020
else:
	print('ERROR python version not found')
	raise SyntaxError('ERROR')

sys.path.append(DIG_PATH)
# #sys.path.append(DIG_PATH_2018)
sys.path.append(DIG_PYTHON_PATH)
# #sys.path.append(DIG_PYTHON_PATH_2018)

os.environ['PATH'] = os.environ['PATH'] + ';' + DIG_PATH
# noqa
import powerfactory

# Load application as Administrator
# #app = powerfactory.GetApplication(username='Administrator', password='Administrator')
# No need to load as Administrator when controlling through Python
app = powerfactory.GetApplication()

# Get list of all users
users = app.GetAllUsers()
k=1
# Loop through until found this user
# for user in users:
# 	if user.loc_name == user_name:
# 		user.prot = protection
# 		user.harm = harmonics
# 		user.arcflash = arcflash
# 		break
#
# app = None

# #app.Show()
subprocess.Popen(os.path.join(DIG_PATH, 'PowerFactory.exe'))

print('PowerFactory Opened')