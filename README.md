# hyperlinks
Manipulating Office word/excel hyperlinks with powershell
 

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

### Prerequisites

What things you need to install the software

1. Powershell v5
2. Windows 10 or Windows Server 2016
3. .Net Framework at least 4.6.2


### Installing

A step by step series of examples that tell you how to prepare the script for running

1. Run powershell as admin
	* Write: Set-Execution Policy Unrestricted.

__WARNING__: Before working in the Windows Registry, it is always a good idea to back it up first, so that you have the option of restoration, should something go wrong.

2. Open run(hit windows key+r)
	* Write regedit
	* Go to: HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\FileSystem
	* Set the value to 1 in LongPathsEnabled
	* Go to: HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\FileSystem
	* Set the value to 1 in LongPathsEnabled


## Running the tests

__WARNING__: Before running the script, it is always recommended to backup the files (word or excel) which contain the hyperlinks, so that you have the option of restoration, should something go wrong.

1. Create a new folder
	* Copy the script in the folder
2. Open Powershell ISE as administrator
	* Write the command to change directory to the folder where script is
	For example: CD "C:\Users\John\Desktop\Script folder"
	* In tab "File", select "Open" and choose the script(   .ps1)
	* Run the script with F5 or click "Play" button in toolbar.
3. A menu will appear
	* 1. Edit
	* 2. Copy
	* 3. Exit
4. If you choose 1 or 2 a second menu appears
	* 1. Word
	* 2. Excel
	* 3. Exit
5. If you chose the "Edit" option in main menu and then either chose Word or Excel
	* A window will popup. This is the path to search for Words or Excels(depends on choice in 2nd 	menu). The script will search all files(word or excel) under this path which contain hyperlinks.

	__WARNING__: If you want to cancel the operation from pop-uped window, it is recommended NOT to choose "CANCEL" but to stop the script.

	* After the window, a text will be appeared in command line: Write the string of path to be 	replaced.
	For example there is a hyperlink in word or excel which targets the path:C:\Users\Desktop\Books
and you decided to change the Books to library,write as the old string the "Books"
	* Then a text will be displayed to write the new string(write "library").
	* It will search the words or excels under the path from popup windows and will replace the part of target of hyperlinks which target to this path.
6. If you chose the "Copy" option in main menu ande then either chose Word or Excel
	* Two windows will popup.
	* The first is the path to search for Words or Excels(depends on choice in 2nd 	menu). The script will search all files(word or excel) under this path which contain hyperlinks.
	* The second will be the destination to copy the hyperlink folders and files. It will create automatically a folder which name is the word or excel with hyperlnks and then it will paste the hyperlinked folders and files in these folders.

	__WARNING__: If you want to cancel the operation from pop-uped windows, it is recommended NOT to choose "CANCEL" but to stop the script.

7. If you want the script to run in shared folder(for example \\192.9.168.2\SharedFolder), in function Find-Folders, change the $browse.SelectedPath="" to $browse.SelectedPath="\\192.9.168.2\SharedFolder"
	
8. If you finished your job go to heading "Installation",follow the steps and change the "Unrestricted" to "Undefined"(step 1). Also in step 2, change the values from 1 to 0.

## Authors

* **Anastasios Antoniou** - *Initial work* - [stashs89](https://github.com/stashs89)

