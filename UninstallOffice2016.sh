#!/bin/bash

### This script removes Microsoft Office 2016 and all associated files, as specified on the 
### official Microsoft site: https://support.office.com/en-us/article/uninstall-office-2016-for-mac-eefa1199-5b58-43af-8a3d-b73dc1a8cae3

### Script created and updated by John Tyler -- jtyler203@gmail.com

read -p "Do you want to completely uninstall Microsoft Office 2016? Y/N? " choice

if [[ $choice = "Y" || $choice = "y" ]] ; then
	echo "Starting Office Removal Script."
	echo "Please enter password: "

	# Move office apps to the trash

	sudo mv -v /Applications/Microsoft\ Excel.app ~/.Trash
	sudo mv -v /Applications/Microsoft\ OneNote.app ~/.Trash
	sudo mv -v /Applications/Microsoft\ Outlook.app ~/.Trash
	sudo mv -v /Applications/Microsoft\ Powerpoint.app ~/.Trash
	sudo mv -v /Applications/Microsoft\ Word.app ~/.Trash

	# Move files from user library to trash

	mv ~/Library/Containers/com.microsoft.errorreporting ~/.Trash
	mv ~/Library/Containers/com.microsoft.Excel ~/.Trash
	mv ~/Library/Containers/com.microsoft.netlib.shipassertprocess ~/.Trash
	mv ~/Library/Containers/com.microsoft.Office365ServiceV2 ~/.Trash
	mv ~/Library/Containers/com.microsoft.Outlook ~/.Trash
	mv ~/Library/Containers/com.microsoft.Powerpoint ~/.Trash
	mv ~/Library/Containers/com.microsoft.RMS-XPCService ~/.Trash
	mv ~/Library/Containers/com.microsoft.Word ~/.Trash
	mv ~/Library/Containers/com.microsoft.onenote.mac ~/.Trash

	# Move more files from user library to trash

	mv ~/Library/Group\ Containers/UBF8T346G9.ms ~/.Trash
	mv ~/Library/Group\ Containers/UBF8T346G9.Office ~/.Trash
	mv ~/Library/Group\ Containers/UBF8T346G9.OfficeOsfWebHost ~/.Trash

#   Echo "Done Uninstalling."
#	osascript -e 'tell application (path to frontmost application as text) to display dialog "Office 2016 has been uninstalled. You may now go back to...uhhh...whatever it is you were doing. Yeah." buttons {"OK"} with icon caution'
	
	read -p "Do you want to empty the Trash? y/n " empty

	if [[ $empty = "y" || $empty = "Y" ]] ; then
		sudo rm -R ~/.Trash/*
		Echo "Trash has been emptied. Office 2016 uninstall is complete."
	else
		exit 0
		Echo "Office 2016 uninstall is complete. You may now return to your regularly scheduled programming."
	fi


else
	echo "Office will not be uninstalled."	
	exit 1
fi


