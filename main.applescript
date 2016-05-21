set instructions to "Please start by reproducing the issue you are having. Press OK, then choose the application you want to troubleshoot" --instructions text
display dialog instructions with title "Microsoft Office for MacOS logs gatherer v0.1"

set officeapp to {"Word", "Excel", "PowerPoint", "Outlook", "OneNote"}

set selectedofficeapp to {choose from list officeapp}

set logword to "syslog | grep -i Word > /tmp/logword.txt"
set logexcel to "syslog | grep -i Excel > /tmp/logexcel.txt"
set logpowerpoint to "syslog | grep -i PowerPoint > /tmp/logpowerpoint.txt"
set logoutlook to "syslog | grep -i Outlook > /tmp/logoutlook.txt"
set logonenote to "syslog | grep -i OneNote > /tmp/logonenote.txt"
if (selectedofficeapp = {{"Word"}}) then
	do shell script logword
	display dialog "/tmp/logword.txt generated"
else if (selectedofficeapp = {{"Excel"}}) then
	do shell script logexcel
	display dialog "/tmp/logexcel.txt generated"
else if (selectedofficeapp = {{"PowerPoint"}}) then
	do shell script logpowerpoint
	display dialog "/tmp/logpowerpoint.txt"
else if (selectedofficeapp = {{"Outlook"}}) then
	do shell script logoutlook
	display dialog "/tmp/logoutlook.txt"
else if (selectedofficeapp = {{"OneNote"}}) then
	do shell script logonenote
	display dialog "/tmp/logonenote.txt"
else
	display dialog "operation cancelled"
end if


