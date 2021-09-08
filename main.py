import clr
clr.AddReference('C:\\Program Files\\Siemens\\Automation\\Portal V16\\PublicAPI\V16\\Siemens.Engineering.dll')
from System.IO import DirectoryInfo, FileInfo
import Siemens.Engineering as tia
import Siemens.Engineering.HW.Features as hwf


project_path = DirectoryInfo ('C:\\Jonas\\TIA')
project_name = 'PythonTest'


processes = tia.TiaPortal.GetProcesses()
mytia = processes[0].Attach()
myproject = mytia.Projects[0]


plc1 = myproject.Devices[0]

print(type(plc1))

print(myproject.Author)

software_container = tia.IEngineeringServiceProvider(plc1.DeviceItems[1]).GetService[hwf.SoftwareContainer]()
software_base = software_container.Software

plc_block = software_base.BlockGroup.Blocks.Find("funkcja1")

# plc_block.Export(FileInfo('C:\\pythonProject\\openess\\imports\\funkcja1.xml'), tia.ExportOptions.WithDefaults)
software_base.BlockGroup.Blocks.Import(FileInfo('C:\\pythonProject\\openess\\imports\\funkcja2.xml'), tia.ImportOptions.Override)



# #Starting TIA
# print ('Starting TIA with UI')
# mytia = tia.TiaPortal(tia.TiaPortalMode.WithUserInterface)

# #Creating new project
# print ('Creating project')
# myproject = mytia.Projects.Create(project_path, project_name)

# #Addding Stations
# print ('Creating station 1')
# station1_mlfb = 'OrderNumber:6ES7 515-2AM01-0AB0/V2.6'
# station1 = myproject.Devices.CreateWithItem(station1_mlfb, 'station1', 'station1')

# print ('Creating station 2')
# station2_mlfb = 'OrderNumber:6ES7 518-4AP00-0AB0/V2.6'
# station2 = myproject.Devices.CreateWithItem(station2_mlfb, 'station2', 'station2')

# print ("Press any key to quit")
# input()
# quit()
