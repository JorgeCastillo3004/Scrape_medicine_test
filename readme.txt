
# Term and conditions applied to account creation.

	https://www.prepssm.it/legal/cgv


#####################################################################
#																	#
#				Package installation instructions.					#
#																	#
#####################################################################

	1- Create a virtual environment.

	2- Execute the following command line: pip install -r requirements.txt

	It will to procedd to install the next libraries:

		-selenium
		-pandas
		-chromedriver-autoinstaller
		-openpyxl

#####################################################################
#																	#
#				Profile chrome configuration.						#
#																	#
#####################################################################

	1- In Chrome navigator type chrome://version/ into address bar.
	2- Copy the default path address (line 374).
	3- Check the profile name in url, chrome://settings/manageProfile
	3- Copy the profile name (line 376).

#####################################################################
#																	#
#				Instructions of use.								#
#																	#
#####################################################################

	1 - Execute in cosole windows the next command line: python mainprepssm.py
	2 - First option ask about get or update all the list of links, or load a previos list loaded, in case that desire to update the list of links remember make zoom out to visulize all the exams links and scroll down until the end of the list of exams, the list of links can be checked in the folder "files_prepssm" in file "examslinks.json".

	2 - If the option selected is load a previos list, the Script load "examslinks.json" file, and in the next step ask if desire load a previos check point or restart the process.

	3 - If the option was continue in previos checkpoint the script continue working, in the las exam link processed.