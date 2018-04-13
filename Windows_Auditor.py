import wmi
import win32com.client
import os
#import time
from progressbar import ProgressBar

c = wmi.WMI()
pbar = ProgressBar()
def main():
	global hostname,file,count
	clean_run()		
	print "Entering Step 1"
	for opsys in c.Win32_OperatingSystem():
		hostname = opsys.CSName
		write_title("Operating System")
		write_to_output("OS", 'Operating system == %s' % opsys.Caption)
	print "Leaving Step 1"

	print "Entering Step 2"
	write_title("Shares")
	for share in c.win32_share():	
		information = "Share Name: " + share.Name + "\n\t Path: " +share.Path + "\n"
		write_to_output("Shares", information)
	print "Leaving Step 2"
	
	print "Entering Step 3"
	for process in c.Win32_Process():
		Runing_Process = "|Process Name: \t" + "%s" %process.Caption + "\tProcess ID: \t" + "%s" %process.ProcessId + "|\n\n"
		write_to_output("Running Processes", Runing_Process)
	print "Leaving Step 3"
	
	print "Entering Step 4"
	for localuser in c.Win32_UserAccount():
		user_info = "|Username: %s" %localuser.Caption + "\nDisabled Status: %s" %localuser.Disabled + "\tPassword Changeable: %s" %localuser.PasswordChangeable + "\nPassword Expires: %s" %localuser.PasswordExpires + "\tPassword Required: %s" %localuser.PasswordRequired +"\nSID: %s" %localuser.SID + "\n\n"
		write_to_output("Local_User_Info", user_info)
	print "Leaving Step 4"

	print "Entering Step 5"
	stopped_services = c.Win32_Service (StartMode="Auto", State="Stopped")
	if stopped_services:
	  for stopped in stopped_services:
		write_to_output("Automatic Services Stopped", "Automatic Service: %s" %stopped.Caption + "\tservice is not running\n")
	else:
	  print "No auto services stopped"
	print "Leaving Step 5"
	  
	print "Entering Step 6"
	runnning_services = c.Win32_Service (StartMode="Auto", State="Running")
	if runnning_services:
	  for running in runnning_services:
		write_to_output("Automatic Services Running", "Automatic Service: %s" %running.Caption + "\tservice is running\n")
	print "Leaving Step 6"
	
	print "Entering Step 7"
	for sitems in c.Win32_StartupCommand ():
	  write_to_output("Startup Items",  "[%s] %s <%s>" % (sitems.Location, sitems.Caption, sitems.Command) + "\n")
	print "Leaving Step 7"
	  

	print "Entering Step 8"
	os.system("secedit /export /cfg group-policy /log export.log")
	print "Leaving Step 8"
	print "Entering Step 9"
	ip_info = os.system('ipconfig /all > ip_info.txt')
	routing_info = os.system('route print > routing_info.txt')
	os.system('regedit /e .\HKEY_LOCAL_MACHINE.reg "HKEY_LOCAL_MACHINE"')
	os.system('regedit /e .\HKEY_USERS.reg "HKEY_USERS"')
	os.system('regedit /e .\HKEY_CURRENT_CONFIG.reg "HKEY_CURRENT_CONFIG"')
	os.system('regedit /e .\HKEY_CURENT_USER.reg "HKEY_CURENT_USER"')
	os.system('regedit /e .\HKEY_CLASSES_ROOT.reg "HKEY_CLASSES_ROOT"')
	os.system('w32tm /query /source > ntpserver.txt')
	os.system('accesschk.exe C:\ -s > Accesscheck.txt')
	os.system('net localgroup > LocalGroups.txt')
	os.system('net Accounts > AccountSettings.txt')
	print "Leaving Step 9"
	
	print "Entering Step 10"
	connected_drives()
	print "Leaving Step 10"
	print "Entering Step 11"
	group_users()
	print "Leaving Step 11"
	print "Entering Step 12"
	patch_lists()
	print "Leaving Step 12"
	
	
	
	
	
	
	cleanup()
	
	
def patch_lists():
	strComputer = "."
	objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
	objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
	colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_QuickFixEngineering")
	for objItem in colItems:
		if objItem.Caption != None:
			tmp =  "Caption: %s" %objItem.Caption + "\n"
		if objItem.Description != None:
			tmp = tmp + "Description: %s" %objItem.Description + "\n"
		if objItem.HotFixID != None:
			tmp = tmp +  "HotFixID: %s" %objItem.HotFixID + "\n"
		if objItem.InstallDate != None:
			tmp = tmp +  "InstallDate: %s" % WMIDateStringToDate(objItem.InstallDate) + "\n"
		if objItem.InstalledBy != None:
			tmp = tmp +  "InstalledBy: %s" % objItem.InstalledBy + "\n"
		if objItem.InstalledOn != None:
			tmp = tmp +  "InstalledOn: %s" %objItem.InstalledOn + "\n"
		if objItem.Name != None:
			tmp = tmp +  "Name: %s" %objItem.Name + "\n\n"
		write_to_output("Patches", tmp)
	
	
def connected_drives():
	strComputer = "."
	objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
	objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
	colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_SystemNetworkConnections")
	for objItem in colItems:
		if objItem.GroupComponent != None:
			tmp =  "GroupComponent: %s"  %objItem.GroupComponent + "\n"
		if objItem.PartComponent != None:
			tmp =  "PartComponent: %s" %objItem.PartComponent + "\n\n"
		write_to_output("Connected Drives", tmp)
			
def group_users():
	strComputer = "."
	objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
	objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
	colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_GroupUser")
	for objItem in colItems:
		if objItem.GroupComponent != None:
			temp =  "GroupComponent: %s" %objItem.GroupComponent + "\n"
		if objItem.PartComponent != None:
			temp = temp + "PartComponent: %s" %objItem.PartComponent + "\n\n"
		write_to_output("Group_User_Info", temp)
		
def WMIDateStringToDate(dtmDate):
    strDateTime = ""
    if (dtmDate[4] == 0):
        strDateTime = dtmDate[5] + '/'
    else:
        strDateTime = dtmDate[4] + dtmDate[5] + '/'
    if (dtmDate[6] == 0):
        strDateTime = strDateTime + dtmDate[7] + '/'
    else:
        strDateTime = strDateTime + dtmDate[6] + dtmDate[7] + '/'
        strDateTime = strDateTime + dtmDate[0] + dtmDate[1] + dtmDate[2] + dtmDate[3] + " " + dtmDate[8] + dtmDate[9] + ": %s" + dtmDate[10] + dtmDate[11] +':' + dtmDate[12] + dtmDate[13]
    return strDateTime		
		
def cleanup():
	os.system('mkdir "Final Folder"')
	os.system('move *.txt ".\Final Folder\"')
	os.system('move *.reg ".\Final Folder\"')
	os.system('move *.log ".\Final Folder\"')
	os.system('move .\group-policy ".\Final Folder\"')

		
def write_to_output(item, information):
	complete_file(item, information)
	individual_file(item, information)
	
def complete_file(item, information):
	file = open('%s.txt' % hostname, 'a') 
	file.write(information)
	file.write("\n")

def write_title(item):
	file = open('%s.txt' % hostname, 'a')
	file.write("\n##############################################################")
	file.write("\n##################" + "\t%s\t" %item+ "##########################\n\n")

def individual_file(item,information):
	file = open('%s.txt' % item, 'a')
	file.write(information)
	file.close()

def clean_run():
	import os
	dir = "./"
	files = os.listdir(dir)
	for file in files:
		if file.endswith(".txt"):
			os.remove(os.path.join(dir,file))
	
if __name__ == "__main__":
	main()