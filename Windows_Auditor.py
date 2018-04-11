import wmi
import os
import time
import progressbar

c = wmi.WMI ()

def main():
	global hostname,file,count
	clean_run()	
	for i in progressbar.progressbar(range(100)):
		time.sleep(0.02)
	
	for os in c.Win32_OperatingSystem():
		hostname = os.CSName
		write_title("Operating System")
		write_to_output("OS", 'Operating system == %s' % os.Caption)
	
	write_title("Shares")
	for share in c.win32_share():	
		information = "Share Name: " + share.Name + "\n\t Path: " +share.Path + "\n"
		write_to_output("Shares", information)
	
	for process in c.Win32_Process():
		Runing_Process = "|Process Name: \t" + "%s" %process.Caption + "\tProcess ID: \t" + "%s" %process.ProcessId + "|\n\n"
		write_to_output("Running Processes", Runing_Process)

	
	
	
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