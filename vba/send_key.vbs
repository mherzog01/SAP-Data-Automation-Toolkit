' Waits for a window with a given name, and sends a given key
'
' Parameters
'   1 - Window title
'   2 - Key to send
'   3 - Timeout (optional)

Dim pTargetWindowTitle
Dim pKeyToSend
Dim cmdArguments
Dim pTimeout

Dim numTries
Dim foundTarget 

do 'main
	logmsg01 "Initializing"

	if wscript.arguments.count < 2 then
	    logmsg "Error:  2 parameters are required"
	    exit do
	end if
	pTargetWindowTitle = Wscript.Arguments(0)
	pKeyToSend = Wscript.Arguments(1)
	if WScript.Arguments.count >= 3 then
	    pTimeout = Wscript.Arguments(2)
	else
		pTimeout = 10
	end if

	logmsg "Target Window Title=" & pTargetWindowTitle
	logmsg "Key to Send=" & pKeyToSend

	Set WshShell = WScript.CreateObject("WScript.Shell")

	logmsg01("Searching for desired window")
	numTries = 0
	do
		numTries = numTries + 1
		foundTarget = WshShell.appactivate(pTargetWindowTitle) 
		if foundTarget then
			exit do
		elseif numTries > pTimeout then 
			logmsg "Timeout reached"
			exit do
		end if
		wscript.sleep 1 * 1000
		logmsg01 "Continuing to search"
	loop

	if not foundTarget then
		exit do
	end if

	logmsg "Sending key"
	wshShell.sendkeys(pKeyToSend)

	exit do
loop 'main
	
'finalize
logmsg01 "End"

sub logmsg(pMsg)
    wscript.echo pMsg
end sub

sub logmsg01(pMsg)
    wscript.echo now() & " - " & pMsg
end sub
