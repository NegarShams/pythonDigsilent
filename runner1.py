import sys
sys.path.append(r"C:\DigSILENT15p1p7\python")

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
#user = app.GetCurrentUser()
k=1


# #app.Show()
#subprocess.Popen(os.path.join(DIG_PATH, 'PowerFactory.exe'))

print('PowerFactory Opened')


# activate project
project = app.ActivateProject("Test1")
prj = app.GetActiveProject()

#FolderWhereStudyCasesAreSaved= app.GetProjectFolder('study')
#FolderWhereStudyCasesAreSaved= app.GetProjectFolderType('study')
folders = prj.GetContents('*.IntPrjfolder')
temp_filtered = filter(lambda folders: folders.loc_name=='Study Cases', folders)
studyCases_folder=list(temp_filtered)
studyCases = studyCases_folder[0].GetContents('*.IntCase')
temp_filtered = filter(lambda folders: folders.loc_name=='Library', folders)
Library=list(temp_filtered)
Library_temp=Library[0]
Library_subfolders=Library_temp.GetContents('*.')
temp_filtered = filter(lambda Library_subfolders: Library_subfolders.loc_name=='Operational Library', Library_subfolders)
oper_folder=list(temp_filtered)
oper_temp=oper_folder[0]
oper_subfolders=oper_temp.GetContents('*.')
temp_filtered = filter(lambda oper_subfolders: oper_subfolders.loc_name=='Characteristics', oper_subfolders)
char_folder=list(temp_filtered)
char_temp=char_folder[0]
char_subfolders=char_temp.GetContents('*.')
temp_filtered = filter(lambda char_subfolders: char_subfolders.loc_name=='Oyster Creek 800MW', char_subfolders)
scen_folder=list(temp_filtered)
scen_temp=scen_folder[0]
scen_subfolders=scen_temp.GetContents('*.')
temp_filtered = filter(lambda scen_subfolders: scen_subfolders.loc_name=='OC_Intact_V', scen_subfolders)
scen_param_folder=list(temp_filtered)
scen_param_temp=scen_param_folder[0]
scen_param_subvalues=scen_param_temp.GetContents('*.')
#studyCases = folders[4].GetContents('*.IntCase')
temp_filtered_L = filter(lambda scen_param_subvalues: scen_param_subvalues.loc_name=='L', scen_param_subvalues)
temp_filtered_R = filter(lambda scen_param_subvalues: scen_param_subvalues.loc_name=='R', scen_param_subvalues)
L=list(temp_filtered_L)
R=list(temp_filtered_R)
i=L[0].vector
i[0]=1.2
L[0].vector=i
Lib_folders = folders[1].GetContents('*.IntPrjfolder')
Charc_scen_folders=Lib_folders[1].GetContents('*.IntPrjfolder')
scn_folders=Charc_scen_folders[10].GetContents('*.')
scn_folders_1=scn_folders[3].GetContents('*.')
temp_filtered = filter(lambda scn_folders_1: scn_folders_1.loc_name=='Vref STATCOM OC', scn_folders_1)
scn_folders=Charc_scen_folders[10].GetContents('*.IntPrjfolder')
#column_list_o = filter(lambda local_x: local_x.startswith(object_name), df.columns)
L_R_temp=Charc_scen_folders[6].GetContents('*.ChaVec')
studycase=studyCases[0]
s=studycase.Activate()
ldf = app.GetFromStudyCase("ComLdf")
Hldf = app.GetFromStudyCase("ComHLdf")
Fsweep = app.GetFromStudyCase("ComFsweep")
ini = app.GetFromStudyCase('ComInc')
terminals = app.GetCalcRelevantObjects("*.ElmTerm")
elmres = app.GetFromStudyCase('Freq.Sweep.ElmRes')

# for terminal in terminals:
#     elmres.AddVars(terminal, 'm:u')

Outs=elmres.GetContents()
Hldf.Execute()
Fsweep.Execute()


#comres = app.GetFromStudyCase('new.ComRes');
comres = app.GetFromStudyCase('FTest1.ComRes');

comres.iopt_csel = 0
comres.iopt_tsel = 0
comres.iopt_locn = 2
comres.ciopt_head = 1
comres.pResult = elmres
comres.f_name = r'C:\Negar\Test1\hola.csv'
comres.iopt_exp = 6
comres.Execute()


n1=studyCases[4].loc_name
k=1
#app.PrintPlain(FolderWhereStudyCasesAreSaved)
# AllStudyCasesInProject= FolderWhereStudyCasesAreSaved.GetContents()
# for StudyCase in AllStudyCasesInProject:
#    app.PrintPlain(StudyCase)
#    StudyCase.Activate()
#
k=1



ldf = app.GetFromStudyCase("ComLdf")
ini = app.GetFromStudyCase('ComInc')

#lines = app.GetCalcRelevantObjects("*.ElmTerm")




sim = app.GetFromStudyCase('ComSim')
Shc_folder = app.GetFromStudyCase('IntEvt');

terminals = app.GetCalcRelevantObjects("*.ElmTerm")
lines = app.GetCalcRelevantObjects("*.ElmTerm")
syms = app.GetCalcRelevantObjects("*.ElmSym")

Shc_folder.CreateObject('EvtSwitch', 'evento de generacion');
EventSet = Shc_folder.GetContents();
evt = EventSet[0];

evt.time = 1.0

evt.p_target = syms[1]

ldf.iopt_net = 0

ldf.Execute()

elmres = app.GetFromStudyCase('Results.ElmRes')

for terminal in terminals:
    elmres.AddVars(terminal, 'm:u', 'm:phiu', 'm:fehz')

for sym in syms:
    elmres.AddVars(sym, 's:xspeed')

ini.Execute()
sim.Execute()

evt.Delete()

comres = app.GetFromStudyCase('ComRes');
comres.iopt_csel = 0
comres.iopt_tsel = 0
comres.iopt_locn = 2
comres.ciopt_head = 1
comres.pResult = elmres
comres.f_name = r'C:\Users\jmmauricio\hola.txt'
comres.iopt_exp = 4
comres.Execute()