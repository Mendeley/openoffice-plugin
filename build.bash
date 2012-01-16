# This is hard-coded for my Windows machine at the moment

# Used to quickly re-build and install the template on a dev machine,
# and open OO writer to start testing

# If the unopkg install fails, it tries installing with the graphical 
# extension manager which will display the error

/c/windows/sysnative/tskill.exe soffice
buildScript.pl 1.1.1 true
echo ""
echo "Running unopkg..."
"/c/Program Files (x86)/OpenOffice.org 3/program/unopkg.exe" add --force Mendeley-1.1.1.oxt
if [ $? -eq 0 ]; then 
	echo "Extension installed successfully"
	/c/windows/sysnative/tskill.exe soffice
	"/c/Program Files (x86)/OpenOffice.org 3/program/swriter.exe" &
else 
	echo "Error with extension, launching extension manager..."
	"/c/Program Files (x86)/OpenOffice.org 3/program/soffice.exe" Mendeley-1.1.1.oxt &
fi
