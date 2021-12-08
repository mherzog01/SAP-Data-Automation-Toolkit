Attribute VB_Name = "ExcelUtil"
Option Explicit

Dim gCurUser

' From https://stackoverflow.com/questions/19185260/open-a-pdf-using-vba-in-excel
Private Declare PtrSafe Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" _
    (ByVal lpFile As String, _
     ByVal lpDirectory As String, _
     ByVal lpResult As String) As Long


Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
     ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


' Returns time in milliseconds since Windows was started
' From https://stackoverflow.com/questions/8631975/measuring-query-processing-time-in-microsoft-access
' Also see https://docs.microsoft.com/en-us/windows/desktop/api/timeapi/nf-timeapi-timegettime
Public Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long


Type typeFileAttr
  FilePath As String
  filePathDir As String
  FileName As String
  fileExtn As String
  fileBase As String
  filePathBase As String
End Type


Function cleanPath(pPath)
    
    Dim tPath
    
    tPath = pPath
    
    'Replace / with \
    tPath = Replace(tPath, "/", "\")
    
    'Remove ending \
    'If Left(tPath, 1) = "\" Then
    '    tPath = Mid(tPath, 2, Len(tPath))
    'End If
    If Right(tPath, 1) = "\" Then
        tPath = Left(tPath, Len(tPath) - 1)
    End If
    
    cleanPath = tPath
End Function


'https://stackoverflow.com/questions/41095060/how-to-get-running-application-name-by-vbscript

Function KillProcessbyName(FileName)
    On Error Resume Next
    'Dim WshShell
    Dim strComputer, objWMIService, colProcesses, objProcess
    'Set WshShell = CreateObject("Wscript.Shell")
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("SELECT ProcessId,Name,CommandLine FROM Win32_Process where name = '" & FileName & "'")
    For Each objProcess In colProcesses
        'If InStr(objProcess.CommandLine, FileName) > 0 Then
        '    If err <> 0 Then
        '        MsgBox err.Description, vbCritical, err.Description
        '    Else
        With objProcess
            Debug.Print "Killing process PID=" & .ProcessID & ", Name=" & .Name & ", Cmd Line=" & .CommandLine
            objProcess.Terminate (0)
        End With
        '    End If
        'End If
    Next
End Function


' -------------------------------------------------------
' PURPOSE: Extract parts of a File Path
' https://www.thespreadsheetguru.com/the-code-vault/2014/3/2/retrieving-the-file-name-extension-from-a-file-path-string
' -------------------------------------------------------
Function getFileAttr(pFilePath) As typeFileAttr

    Dim lastSlash As Long
    Dim extnLen As Integer
   
    ' filePath
    getFileAttr.FilePath = pFilePath
   
    ' fileExtn
    If InStr(1, pFilePath, ".", vbTextCompare) > 0 Then
        getFileAttr.fileExtn = Right(pFilePath, Len(pFilePath) - InStrRev(pFilePath, "."))
    Else
        getFileAttr.fileExtn = ""
    End If
    
    ' filePathDir and fileName
    If InStr(1, pFilePath, "\", vbTextCompare) > 0 Then
        lastSlash = InStrRev(pFilePath, "\")
        getFileAttr.filePathDir = Left(pFilePath, lastSlash - 1)
        getFileAttr.FileName = Right(pFilePath, Len(pFilePath) - lastSlash)
    Else
        getFileAttr.filePathDir = ""
        getFileAttr.FileName = pFilePath
    End If
    
    ' fileBase
    extnLen = Len(getFileAttr.fileExtn)
    getFileAttr.fileBase = Left(getFileAttr.FileName, Len(getFileAttr.FileName) - (extnLen + 1))
    getFileAttr.filePathBase = Left(getFileAttr.FilePath, Len(getFileAttr.FilePath) - (extnLen + 1))
    
End Function


Public Sub DocDatabase()
 '====================================================================
 ' Name:    DocDatabase
 ' Purpose: Documents the database to a series of text files
 '
 ' Author:  Arvin Meyer
 ' Date:    June 02, 1999
 ' Comment: Uses the undocumented [Application.SaveAsText] syntax
 '          To reload use the syntax [Application.LoadFromText]
 ' From:    http://www.accessmvp.com/Arvin/DocDatabase.txt
 '====================================================================
On Error GoTo Err_DocDatabase
Dim dbs As Database
Dim cnt As Container
Dim doc As Document
Dim i As Integer
Dim dbPath, strCodeExportFolder, strCodeExportFilesPattern As String
Dim ans
Dim strDir

Set dbs = CurrentDb() ' use CurrentDb() to refresh Collections

ans = MsgBox("Exporting code for " & dbs.Name, vbOKCancel)
If ans <> vbOK Then
    Exit Sub
End If

Debug.Print "Process begins"

dbPath = Application.CurrentProject.Path
strCodeExportFolder = dbPath & "\Code Export"

strCodeExportFilesPattern = strCodeExportFolder & "\*.*"
strDir = Dir(strCodeExportFilesPattern)
If strDir <> "" Then
    Debug.Print "Deleting files in folder " & strCodeExportFolder
    Kill strCodeExportFolder & "\*.*"
End If

Set cnt = dbs.Containers("Forms")
For Each doc In cnt.Documents
    Application.SaveAsText acForm, doc.Name, strCodeExportFolder & "\" & doc.Name & ".txt"
Next doc

Set cnt = dbs.Containers("Reports")
For Each doc In cnt.Documents
    Application.SaveAsText acReport, doc.Name, strCodeExportFolder & "\" & doc.Name & ".txt"
Next doc

Set cnt = dbs.Containers("Scripts")
For Each doc In cnt.Documents
    Application.SaveAsText acMacro, doc.Name, strCodeExportFolder & "\" & doc.Name & ".txt"
Next doc

Set cnt = dbs.Containers("Modules")
For Each doc In cnt.Documents
    Application.SaveAsText acModule, doc.Name, strCodeExportFolder & "\" & doc.Name & ".txt"
Next doc

For i = 0 To dbs.QueryDefs.Count - 1
    Application.SaveAsText acQuery, dbs.QueryDefs(i).Name, strCodeExportFolder & "\" & dbs.QueryDefs(i).Name & ".txt"
Next i

Set doc = Nothing
Set cnt = Nothing
Set dbs = Nothing

Exit_DocDatabase:
    Debug.Print "Process complete"
    Exit Sub


Err_DocDatabase:
    Select Case err

    Case Else
        MsgBox err.Description
        Resume Exit_DocDatabase
    End Select

End Sub


Public Function GetFileAssociation(ByVal sFilepath As String) As String
Dim i               As Long
Dim e               As String
    GetFileAssociation = "File not found!"
    If Dir(sFilepath) = vbNullString Or sFilepath = vbNullString Then Exit Function
    GetFileAssociation = "No association found!"
    e = String(260, Chr(0))
    i = FindExecutable(sFilepath, vbNullString, e)
    If i > 32 Then GetFileAssociation = Left(e, InStr(e, Chr(0)) - 1)
End Function


'=================================
' OpenFile
'=================================
Sub openFile(pPath, Optional pWindowSettings = vbNormalFocus)
    ShellExecute 0, "Open", pPath, "", "", pWindowSettings
End Sub


'=================================
' fileAction
'=================================
' Function succeeds if return code > 32
' https://docs.microsoft.com/en-us/windows/desktop/api/shellapi/nf-shellapi-shellexecutea
Function fileAction(pPath, Optional pParams = "", Optional pAction = "")
    Dim txtParams, txtAction
    Dim retCode
    
    retCode = ShellExecute(0, txtAction, pPath, pParams, "", vbNormalNoFocus)
    fileAction = retCode
End Function


Sub testgetIE()
Dim BrowserExe
    BrowserExe = GetReg("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\IEXPLORE.EXxE\")
Debug.Print BrowserExe
End Sub


'-----------------------------------------
' openFileInBrowser
'-----------------------------------------
Function openFileInBrowser(pURL, Optional pWindowState = vbMaximizedFocus)
    Dim BrowserExe
    Dim retCode
    
    'https://stackoverflow.com/questions/4212002/how-to-find-out-from-the-windows-registry-where-ie-is-installed
    BrowserExe = ""
    On Error Resume Next
    BrowserExe = GetReg("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\IEXPLORE.EXE\")
    On Error GoTo 0
    'Debug.Print IEexe
    
    'Gives error:  Debug.Print "Assoc=" & GetFileAssociation(txtURL)
    
    '--Fails openFile """" & txtURL & """"
    '--Fails:  Shell """c:\program files (x86)\Adobe\Acrobat Reader DC\Reader\acrord32.exe"" """ & txtURL & """"
    '**Works:  Shell """C:\Program Files (x86)\Internet Explorer\iexplore.exe"" """ & txtURL & """"
    'Shell """c:\tmp\test.bat"" """ & txtURL & """"
    'retCode = fileAction(txtURL)
    'Debug.Print retCode
    
    'fileAction """" & IEexe & """", """" & txtURL & """", "explore"
    'fileAction """" & IEexe & """", "http://www.google.com"

    If BrowserExe <> "" Then
        retCode = Shell("""" & BrowserExe & """ """ & pURL & """", pWindowState)
    Else
        openFile pURL
    End If
    'Debug.Print retCode

End Function


Function getErrorDesc()
    If err.Number = 0 Then
        getErrorDesc = ""
    Else
        getErrorDesc = "Error:  " & err.Description & ".  Number=" & err.Number & ".  Source=" & err.Source
    End If
End Function


Sub DispMsg(pMsg)
    ' Ideally, we'd poll an application data structure for status info,and display default
    ' values if no updates are needed.  This will clear old statuses
    ' Or use a data structure, not passing parameters
    'Application.SysCmd acSysCmdSetStatus, pMsg
End Sub


Sub ClearMsg()
    DispMsg " "
End Sub

Sub DispMsg01(pMsg)
    Debug.Print pMsg
    DispMsg (pMsg)
End Sub

Sub DispMsg02(pMsg)
    DispMsg pMsg
End Sub



Function LogMsg(pMsg)
    
    LogMsg = createMessage(pMsg, "Message")
    
End Function


' Create an entry in the log table with no UI interaction unless there is an error
Function createMessage(pMsg, pMsgType)
    Dim i
    'Dim rsLog As DAO.Recordset
    Dim tErrorsCount
    Dim tCurError
    Dim tStrError
    
    'Disable write to db for now
    Debug.Print Now() & " - " & pMsg
    Exit Function
    
End Function

'Sub ClearStatusMsg()
'
'End Sub

Sub LogMsg01(pMsg)
    DispMsg Now & ":  " & pMsg
    LogMsg pMsg
End Sub

Sub LogMsg02(pMsg)
    LogMsg pMsg
End Sub

Sub LogError01()
    Dim sMsg As String
    
    sMsg = getErrorDesc
    LogError sMsg
End Sub


Sub LogError(pMsg)
    If Left(pMsg, 6) = "Error:" Then
        LogMsg01 pMsg
    Else
        LogMsg01 "Error:  " & pMsg
    End If
End Sub




'Make function in case user presses ESC or otherwise interrups process
'Trapping for this is not yet supported
Function AppWait(pWaitSec)

    Dim curTime, initTime As Date
    
    initTime = Now()
    Do
        DoEvents
        curTime = Now()
    Loop While initTime + pWaitSec / 3600 / 24 > curTime

End Function



' From https://stackoverflow.com/questions/424331/get-the-current-temporary-directory-path-in-vbscript
'
' Can also use
'    CreateObject("WScript.Shell").ExpandEnvironmentStrings("%Temp%")
Function getTempFolder()
    Const WindowsFolder = 0
    Const SystemFolder = 1
    Const TemporaryFolder = 2
    
    Dim fso
    Dim tempFolder

    Set fso = CreateObject("Scripting.FileSystemObject")
    tempFolder = fso.GetSpecialFolder(TemporaryFolder)
    getTempFolder = tempFolder
End Function


'Replace ' with '' in incoming string and perform other escaping
'Use to prepare SQL
Function escapeStringForSQL(pString)
    escapeStringForSQL = Replace(pString, "'", "''")
End Function



' From https://stackoverflow.com/questions/35997892/show-users-on-access-database
Function listLoggedInUsers()
    Dim cn 'As New ADODB.Connection
    Dim rs 'As New ADODB.Recordset
    Dim i, j As Long
    Dim loggedInUsers
    
    Set cn = CurrentProject.Connection

    ' The user roster is exposed as a provider-specific schema rowset
    ' in the Jet 4.0 OLE DB provider.  You have to use a GUID to
    ' reference the schema, as provider-specific schemas are not
    ' listed in ADO's type library for schema rowsets

    ' https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/schemaenum?view=sql-server-2017
    Const adSchemaProviderSpecific = -1
    Set rs = cn.OpenSchema(adSchemaProviderSpecific, _
    , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")

    'Output the list of all users in the current database.
    loggedInUsers = "x"
    'Debug.Print rs.Fields(0).Name, "", rs.Fields(1).Name, "", rs.Fields(2).Name, rs.Fields(3).Name

    While Not rs.EOF
        'Debug.Print rs.Fields(0), rs.Fields(1), rs.Fields (2), rs.Fields(3)
        ' Is user connected?
        If rs.Fields(2) Then
            If loggedInUsers > "" Then loggedInUsers = loggedInUsers & ","
            loggedInUsers = loggedInUsers & Trim(rs.Fields(0)) & "\" & Trim(rs.Fields(1))
        End If
        rs.MoveNext
    Wend
    
    listLoggedInUsers = loggedInUsers

End Function



Function GetReg(pRegPath)
Dim txtGetReg
'https://stackoverflow.com/questions/32345238/read-and-write-from-to-registry-in-vba
txtGetReg = CreateObject("WScript.Shell").RegRead(pRegPath)
GetReg = txtGetReg
End Function




Function getPrinterListFromOS(Optional pComputer) As String()

Dim objWMIService, objSWbemServices, colItems, objItem
Dim printerList() As String
Dim i
Dim strComputer

    If IsMissing(pComputer) Then strComputer = "."

    Set objWMIService = CreateObject("WbemScripting.SWbemLocator")
    Set objSWbemServices = objWMIService.ConnectServer(strComputer, "root\cimv2")
    Set colItems = objSWbemServices.ExecQuery("Select * from Win32_PrinterConfiguration")
    i = 0
    For Each objItem In colItems
       ReDim Preserve printerList(i)
       printerList(i) = objItem.Name
       i = i + 1
       'cbPrinter.AddItem objItem.Name
    Next
    'For i = 0 To UBound(printerList)
    '    Debug.Print i, printerList(i)
    'Next
    
    getPrinterListFromOS = printerList
End Function


Function getDefaultPrinterFromApp()
    Dim txtDefaultPrinter
    txtDefaultPrinter = Application.Printer.DeviceName
    
    'NOTE: To use this default printer to print, you must
    '      ensure case (upper/lower) of defaultPrinter matches list
    '           e.g.
    '               If objItem.Name = defaultPrinter Then defaultPrinter = objItem.Name
    getDefaultPrinterFromApp = txtDefaultPrinter
End Function


Function getFriendlyPrinterName(pPrinter)
    Dim aStrings() As String
    'For printers of form "\\<server>\<printer>", return only "<printer>"
    aStrings = Split(pPrinter, "\")
    getFriendlyPrinterName = aStrings(UBound(aStrings))
End Function




' Error codes
'   10 - Invalid printer
Function validatePrinter(pPrinter)

    Dim printerList
    Dim boolFound
    Dim retCode
    Dim i
    
    printerList = getPrinterList
    boolFound = False
    retCode = 10 'Invalid printer
    For i = 0 To UBound(printerList)
        If printerList(i) = pPrinter Then
            boolFound = True
            Exit For
        End If
    Next
    If boolFound Then
        retCode = 0
    End If
    
    GoTo finalize

finalize:
    validatePrinter = retCode
End Function



Function openFileExplorer(pFolder, Optional pFocus = vbNormalFocus)
    Dim fso
    Dim retCode
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(pFolder) Then
        retCode = 1
        GoTo finalize
    End If
    
    Shell "Explorer """ & pFolder & """", pFocus
    retCode = 0
    
finalize:
    openFileExplorer = retCode
End Function


' from https://stackoverflow.com/questions/42934778/how-to-open-the-print-queue-window-in-vb-net
Sub showPrinterQueueWindow(pPrinterName)
    Dim OpenCMD
    Set OpenCMD = CreateObject("wscript.shell")
    OpenCMD.Run ("rundll32.exe printui.dll,PrintUIEntry /o /n """ & pPrinterName)
End Sub


Sub showAllPrintersAndDevices(Optional pWindowSettings = vbNormalFocus)
    Shell "control printers", pWindowSettings
End Sub


Function openFolder(pFolderPath)

    Dim txtFolderPath
    Dim fso
    Dim retCode

    retCode = 1
    txtFolderPath = pFolderPath

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(txtFolderPath) Then
        Debug.Print txtFolderPath
        MsgBox "Error - invalid folder.  Path=" & txtFolderPath
        GoTo finalize
    End If
    DispMsg "Opening folder " & txtFolderPath
    openFileExplorer txtFolderPath
    ClearMsg
    retCode = 0
finalize:
    openFolder = retCode
    Exit Function
End Function




Function AddCommandBarCtrl(ByRef pCmdShortcutMenu, pCaption, pAction, Optional pParam = "", Optional pBeginGroup = False)

Dim retCode

On Error GoTo errhandler

    With pCmdShortcutMenu
        With .Controls.Add(Type:=1)
            .BeginGroup = pBeginGroup
            .Caption = pCaption
            .OnAction = pAction
            .Parameter = pParam
        End With

    End With
    
    retCode = 0
    
ExitSub:
    AddCommandBarCtrl = retCode
    Exit Function
errhandler:
    retCode = err.Number
    Debug.Print "AddCommandBarCtrl", getErrorDesc
    Resume ExitSub
End Function


Function DeleteShortCutMenu(MenuName As String)
Dim retCode
Dim boolExistsInCollection As Boolean
    On Error GoTo errhandler
    boolExistsInCollection = ExistsInCollection(CommandBars, MenuName)
    If boolExistsInCollection Then
        CommandBars(MenuName).Delete
        retCode = 0
    Else
        retCode = 1
    End If
ExitSub:
    DeleteShortCutMenu = retCode
    Exit Function
errhandler:
    retCode = err.Number
    Debug.Print "DeleteShortCutMenu", getErrorDesc
    Resume ExitSub
End Function


Function viewTextFile(pFilePath, Optional pWindowSetings = vbNormalFocus)
    Dim txtValidatedFilePath
    txtValidatedFilePath = Dir(pFilePath)
    If txtValidatedFilePath = "" Then
        MsgBox "File not found " & pFilePath
        GoTo finalize
    End If
    Shell "notepad.exe """ & pFilePath & """", pWindowSetings
finalize:
End Function


'-----------------------------------------------------
' ExistsInCollection - does an object exist in a collection
'
' From https://stackoverflow.com/questions/137845/determining-whether-an-object-is-a-member-of-a-collection-in-vba
Public Function ExistsInCollection(col, key As Variant) As Boolean
'Public Function ExistsInCollection(col, As Collection, key As Variant) As Boolean
    On Error GoTo err
    ExistsInCollection = True
    IsObject (col.Item(key))
    Exit Function
err:
    ExistsInCollection = False
End Function


Function openAcrobatViaShell(pParams, Optional pWindowState = vbMaximizedFocus)
    Dim AcrobatExe
    Dim retCode
    
    'https://stackoverflow.com/questions/4212002/how-to-find-out-from-the-windows-registry-where-ie-is-installed
    AcrobatExe = ""
    On Error Resume Next
    AcrobatExe = GetReg("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\ACROBAT.EXE\")
    On Error GoTo 0
    If AcrobatExe = "" Then
        MsgBox "Unable to find Acrobat"
        GoTo finalize
    End If
    Debug.Print AcrobatExe, pParams
    
    On Error Resume Next
    retCode = 0
    retCode = Shell("""" & AcrobatExe & """ " & pParams, pWindowState)
    'retCode = fileAction(pParams, "")
    On Error GoTo 0
    If retCode = 0 Then
        MsgBox "Error starting Acrobat.  Parameters=" & pParams & " (" & retCode & ")"
        GoTo finalize
    End If
    'Debug.Print retCode
    
finalize:

End Function


Function getCurUser()
    If IsEmpty(gCurUser) Then
       'gCurUser = Environ("UserName")
       ' Environ is less secure....
       gCurUser = CreateObject("WScript.Network").UserName
    End If

    getCurUser = gCurUser
End Function


'https://stackoverflow.com/questions/14219455/excel-vba-code-to-copy-a-specific-string-to-clipboard
'Object is MSForms.DataObject - Microsoft Forms 2.0 Object Library at C:\windows\system32\fm20.dll - https://stackoverflow.com/questions/5552299/how-to-copy-to-clipboard-using-access-vba
Sub CopyTextToClipboard(text As String)
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.setText text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub



Function loadTextDictToClipboard(pDict As Dictionary, Optional pAddCrLF = True)
    CopyTextToClipboard textDictToString(pDict, pAddCrLF)
End Function

'TODO Return focus to current window - use https://stackoverflow.com/questions/51726104/how-to-set-a-reference-to-a-running-object-in-access-vba
Function waitForWindow(pTitle, Optional pTimeout = 60, Optional pCheckFreq = 1, Optional pReturnFocus = False, Optional pLogLevel = 0)

    Dim foundTitle
    Dim startTime
    Dim timeout
    Dim retCode
    
    startTime = Now
    foundTitle = False
    timeout = False
    Do While Not foundTitle
        'On Error GoTo end_try01
        On Error Resume Next
        err.Clear
        AppActivate pTitle
        If err.Number = 0 Then
            foundTitle = True
            Exit Do
        End If
        On Error GoTo 0
        If (Now - startTime) * 3600 * 24 > pTimeout Then
            timeout = True
            Exit Do
        End If
        AppWait pCheckFreq
        If pLogLevel > 0 Then Debug.Print Now & " - Waiting"
        DoEvents
    Loop
    
    
    If foundTitle Then
        retCode = 0
        If pReturnFocus Then
            'TODO Return focus to source window
        End If
    ElseIf timeout Then
        retCode = 1
    Else
        retCode = 99 'Some other error
    End If
    
    If pLogLevel > 0 Then Debug.Print Now & " - exiting.  Code=" & retCode
    
    waitForWindow = retCode

End Function


'https://stackoverflow.com/questions/40729968/exported-file-opens-after-macro-completes-unwanted
Function waitForExcelWorkbook(pWbkName, Optional pTimeout = 120, Optional pCheckFreq = 1, Optional pLogLevel = 0)

    Dim foundApp As Boolean
    Dim foundWbk As Boolean
    Dim startTime As Date
    Dim timeout As Boolean
    Dim retCode
    Dim xclApp As Object
    Dim xclWbk As Object
    Dim retStx(3)
    
    'Get Excel app object
    startTime = Now
    foundApp = False
    timeout = False
    Do While Not foundApp
        'On Error GoTo end_try01
        On Error Resume Next
        err.Clear
        Set xclApp = GetObject(, "Excel.Application")
        If err.Number = 0 Then
            foundApp = True
            Exit Do
        End If
        On Error GoTo 0
        If (Now - startTime) * 3600 * 24 > pTimeout Then
            timeout = True
            Exit Do
        End If
        AppWait pCheckFreq
        If pLogLevel > 0 Then Debug.Print Now & " - Waiting for Excel application"
        DoEvents
    Loop 'Get Excel app object
    
    If timeout Or Not foundApp Then GoTo finalize
    
    'Get Excel workbook
    foundWbk = False
    timeout = False
    Do While Not foundWbk
        'On Error GoTo end_try01
        On Error Resume Next
        err.Clear
        Set xclWbk = xclApp.Workbooks.Item(pWbkName)
        If err.Number = 0 Then
            foundWbk = True
            Exit Do
        End If
        On Error GoTo 0
        If (Now - startTime) * 3600 * 24 > pTimeout Then
            timeout = True
            Exit Do
        End If
        AppWait pCheckFreq
        If pLogLevel > 0 Then Debug.Print Now & " - Waiting for workbook " & pWbkName
        DoEvents
    Loop 'Get Excel app object
    
finalize:
    If foundWbk Then
        retCode = 0
    ElseIf Not foundApp Then
        If timeout Then
            retCode = 1
        Else
            retCode = 90
        End If
    ElseIf Not foundWbk Then
        If timeout Then
            retCode = 11
        Else
            retCode = 91
        End If
    End If
    
    If pLogLevel > 0 Then Debug.Print Now & " - exiting.  Code=" & retCode
    
    retStx(1) = retCode
    Set retStx(2) = xclWbk
    retStx(3) = pTimeout
    
    waitForExcelWorkbook = retStx

End Function


'Remove "." and other characters that the MS Access import cannot process
Function cleanExcelHeaderForAccess(pFilePath, Optional pLogLevel = 0)
    Dim xclWkb
    Dim xclApp
    Dim xclWs
    Dim xclRange
    Dim xclCell
    Dim cleanValue
    Dim retCode
    Dim colNum
    
    If pLogLevel >= 1 Then Debug.Print "Starting process"
    
    Set xclApp = CreateObject("Excel.Application")
    If pLogLevel >= 2 Then xclApp.Visible = True
    Set xclWkb = xclApp.Workbooks.Open(pFilePath, ReadOnly:=False, IgnoreReadOnlyRecommended:=True)
    
    'Assume there is just one worksheet in workbook
    Set xclWs = xclWkb.Worksheets(1)
    'TODO select only columns which have data, not entire row
    xclRange = xclWs.Range("1:1")
    If pLogLevel >= 2 Then Debug.Print "# cells selected=" & UBound(xclRange, 2)
    colNum = 0
    For Each xclCell In xclRange
        colNum = colNum + 1
        If IsEmpty(xclCell) Then GoTo end_loop01
        cleanValue = xclCell
        cleanValue = Replace(cleanValue, ". ", " ")
        cleanValue = Replace(cleanValue, ".", " ")
        xclWs.Cells(1, colNum).Value = cleanValue
end_loop01:
    Next
    
    xclWkb.Save
    xclWkb.Close
    xclApp.Quit
    
    If pLogLevel >= 1 Then Debug.Print "Process ends"
    
End Function

Sub testclean()
    Debug.Print cleanExcelHeaderForAccess("c:\users\mherzo\downloads\test.xlsx", 2)
End Sub


Function windowExists(pTitle) As Boolean
    windowExists = (waitForWindow(pTitle, 0, 0) = 0)
End Function


'https://stackoverflow.com/questions/25424469/vba-get-taskbar-applications
Sub dispProcesses()
    Dim W As Object
    Dim ProcessQuery As String
    Dim processes As Object
    Dim process As Object
    Set W = GetObject("winmgmts:")
    ProcessQuery = "SELECT * FROM win32_process"
    Set processes = W.ExecQuery(ProcessQuery)
    For Each process In processes
        'Debug.Print process.Name, process.Description
    Next
    Debug.Print "#=" & processes.Count
    Set W = Nothing
    Set processes = Nothing
    Set process = Nothing

End Sub



Function textDictToString(pDict As Dictionary, Optional pAddCrLF = True) As String
    Dim text
    Dim d
    text = ""
    For Each d In pDict
        text = text & pDict(d)
        If pAddCrLF Then text = text & vbCrLf
    Next d
    textDictToString = text
End Function





Sub TestOpenAcrobat()
    Dim strFilePath
    strFilePath = "//allergan.sharepoint.com@ssl/davwwwroot/sites/TissuePortal/tds/TST/AppTest/CTDN/DMWEST-18-T1600_Final_Record_Release_(2of2).pdf"
    strFilePath = "//allergan.sharepoint.com@ssl/davwwwroot/sites/TissuePortal/tds/TST/AppTest/TDS/TEST - TDS18-2228TDS18-2228.pdf"
    openAcrobatViaShell strFilePath
End Sub


Function testx()
Debug.Print Now
End Function

