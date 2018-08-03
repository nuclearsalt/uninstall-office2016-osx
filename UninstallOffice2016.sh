#!/bin/bash

### This script removes Microsoft Office 2016 and all associated files, as specified on the 
### official Microsoft site: https://support.office.com/en-us/article/uninstall-office-201or-mac-eefa1199-5b58-43af-8a3d-b73dc1a8cae3

### Script created and updated by John Tyler -- jtyler203@gmail.com

read -p "Do you want to completely uninstall Microsoft Office 2016? Y/N? " choice

if [[ $choice = "Y" || $choice = "y" ]] ; then
	echo "Starting Office Removal Script. Please enter password"

	# Move office apps to the OfficeFileDump folder

	mkdir ~/OfficeFileDump

	sudo mv -v /Applications/Microsoft\ Excel.app ~/OfficeFileDump/
	sudo mv -v /Applications/Microsoft\ OneNote.app ~/OfficeFileDump/
	sudo mv -v /Applications/Microsoft\ Outlook.app ~/OfficeFileDump/
	sudo mv -v /Applications/Microsoft\ Powerpoint.app ~/OfficeFileDump/
	sudo mv -v /Applications/Microsoft\ Word.app ~/OfficeFileDump/
	sudo mv -v /Applications/OneDrive.app ~/OfficeFileDump/

	# Move files from user library to OfficeFileDump folder

	mv -v ~/Library/Containers/com.microsoft.errorreporting ~/OfficeFileDump/
	mv -v ~/Library/Containers/com.microsoft.Excel ~/OfficeFileDump/
	mv -v ~/Library/Containers/com.microsoft.netlib.shipassertprocess ~/OfficeFileDump/
	mv -v ~/Library/Containers/com.microsoft.Office365ServiceV2 ~/OfficeFileDump/
	mv -v ~/Library/Containers/com.microsoft.Outlook ~/OfficeFileDump/
	mv -v ~/Library/Containers/com.microsoft.Powerpoint ~/OfficeFileDump/
	mv -v ~/Library/Containers/com.microsoft.RMS-XPCService ~/OfficeFileDump/
	mv -v ~/Library/Containers/com.microsoft.Word ~/OfficeFileDump/
	mv -v ~/Library/Containers/com.microsoft.onenote.mac ~/OfficeFileDump/

	# Move more files from user library to OfficeFileDump folder

	mv -v ~/Library/Group\ Containers/UBF8T346G9.ms ~/OfficeFileDump/
	mv -v ~/Library/Group\ Containers/UBF8T346G9.Office ~/OfficeFileDump/
	mv -v ~/Library/Group\ Containers/UBF8T346G9.OfficeOsfWebHost ~/OfficeFileDump/

	# Move OfficeFileDump folder to Trash

	sudo chown -R $USER:staff ~/OfficeFileDump
	sudo chmod -R 755 ~/OfficeFileDump
	echo "Please wait. This may take a while..."
	sudo rsync -a --remove-source-files ~/OfficeFileDump ~/.Trash/ && rm -R ~/OfficeFileDump


	## choose whether or not to empty trash
	
	read -p "Do you want to empty the Trash? y/n " empty

	if [[ $empty = "y" || $empty = "Y" ]] ; then
		Echo "Please wait. Thanks."
		sudo rm -R ~/.Trash/*
		Echo "The trash has been emptied, unlike the one in your kitchen."
		sleep 2 
		echo "Office 2016 uninstall is complete."
	else
		Echo "Office 2016 uninstall is complete." 
		sleep 2
		echo "You may now return to looking at memes on the internet"
		exit 0	
	fi


else
	echo "Office will not be uninstalled. Office will not be told what to do."	
	exit 1
fi


