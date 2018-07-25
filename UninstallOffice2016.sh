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

	## empty trash
	
	read -p "Do you want to empty the Trash? y/n " empty

	if [[ $empty = "y" || $empty = "Y" ]] ; then
		sudo rm -R ~/.Trash/*
		Echo "The trash has been emptied, unlike the one in your kitchen. Office 2016 uninstall is complete."
	else
		Echo "Office 2016 uninstall is complete. You may now return to looking at memes on the internet"
		exit 0	
	fi


else
	echo "Office will not be uninstalled. Office will not be told what to do."	
	exit 1
fi


