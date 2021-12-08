Attribute VB_Name = "vbaSAPScriptingLibrary"
'Setup Required:
' 1.  References needed:
'     - "SAP GUI Scripting API - C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapfewse.ocx
'     - "Microsoft Scripting Runtime"
'
' 2.  Turn off SAP scripting prompts.  See https://blogs.sap.com/2015/06/09/tips-stop-the-pop-up-sap-gui-security-remeber-my-decision/
'

Option Explicit

Type reportInfoType
    TCode As String
    outputFileName As String
    outputFolder As String
    outputFilePath As String
    params As New Dictionary
End Type

'Dim SAPGUIAuto As Object
'Dim App, Connection As Object
Dim session As Object


' Do needed initialization
'
' Return "session" object so it can be worked with in the calling procedure,
' but also set it globally in this module so it's available to all code
'
' Assumptions:
' - Start at SAP GUI menu
' - Set curReportInfo and params (passed in pReportInfo and pParams)
'
' TODO - detect if not starting at SAP menu
' TODO - capture elapsed time of process
' TODO - move init of SAP Scripting objects into a function or subroutine
' TODO - Handle situation if SAP is already launched, or is not on the main screen.
'        The following will detect if an SAP window exists:
'           '--https://stackoverflow.com/questions/50049675/how-to-logon-on-sap-with-sso-via-vbs
'           WinTitle = "SAP"
'           While Not WSHShell.AppActivate(WinTitle)
'               WScript.Sleep 250
'           Wend
'
Function InitSAP(ByRef pReportInfo As reportInfoType) As Object 'session

    Dim retCode

    LogMsg01 "initSAP - Process begins"

    Set session = getSession(1)
    
    Set InitSAP = session
    
    ' Launch TCode
    retCode = launchTCode(pReportInfo.TCode, 1)
    
    LogMsg01 "initSAP - Process ends.  Code=" & retCode
    
End Function



Function launchTCode(pTCode, Optional logLevel = 0)
    
    Dim retCode
    
    err.Clear
    On Error GoTo catch_block01
    If logLevel >= 1 Then LogMsg01 "Executing TCode=" & pTCode
    
    ' Insert Transaction
    session.FindById("wnd[0]/tbar[0]/okcd").text = pTCode
        
    ' Launch Transaction
    session.FindById("wnd[0]/tbar[0]/btn[0]").Press
    ' Could also send the "Enter" keystroke:
    'session.findById("wnd[0]").sendVKey 0
    
    retCode = 0
    
    GoTo after_catch_block01
    On Error Resume Next
    
catch_block01:
    retCode = err.Number
    LogError01
after_catch_block01:
    launchTCode = retCode
End Function



Function setParam(ByRef pParamDict, pName, Optional pLowValue, Optional pHighValue, Optional pIncludeDict, Optional pExcludeDict, Optional pBaseName)
    Dim curParamValues As New vbaSAPParamValues
    
    With curParamValues
        .LabelName = pName
        If Not IsMissing(pLowValue) Then
            .LowValue = pLowValue
        End If
        If Not IsMissing(pHighValue) Then
            .HighValue = pHighValue
        End If
        If Not IsMissing(pIncludeDict) Then
            Set .IncludeDict = pIncludeDict
        End If
        If Not IsMissing(pExcludeDict) Then
            Set .ExcludeDict = pExcludeDict
        End If
        If Not IsMissing(pBaseName) Then
            .BaseName = pBaseName
        End If
    End With
    
    pParamDict.Add pName, curParamValues
    
End Function



' Assume session is already set
Sub ClearAllFields()

    'Dim SAPGUIAuto As Object
    'Dim sapapp As Object
    'Dim sapcon As Object
    'Dim session As Object
    Dim Area As Object
    Dim i As Long
    Dim Children As Object
    Dim obj As Object
    Dim curName As String

    'Set SAPGUIAuto = GetObject("SAPGUI")
    'Set sapapp = SAPGUIAuto.GetScriptingEngine
    'Set sapcon = sapapp.Children(0)
    'Set session = sapcon.Children(0)
    Set Area = session.FindById("wnd[0]/usr")
    Set Children = Area.Children()

    'The variable Children seems to get reset bo selecting the clear menu
    'Debug.Print "User Area=" & Area.ID
    For i = 0 To Area.Children.Count() - 1
        Do
            'Debug.Print i, Obj.Name
            'Set Obj = Children(CInt(i))
            Set obj = Area.Children(CInt(i))
            clearField obj
            Exit Do
        Loop
    Next i
    Set Children = Nothing
    Set obj = Nothing

End Sub


Sub clearField(pObj)

    If Not canUpdateObj(pObj.TypeAsNumber) Or Not pObj.Changeable Then GoTo finalize
    
    'Debug.Print "Name=" & pObj.Name & "  Value=" & pObj.Text
    pObj.SetFocus
    session.FindById("wnd[0]/mbar/menu[1]/menu[6]").Select
finalize:
End Sub


Function parseObject(pObject) As vbaSAPGUIObjectInfo

    ' retVals map
    ' Element 1 = Object Path
    '         2 = Object Name
    '         3 = Object Value
    '         4 = SAP Object Type (Obj.ObjectType)
    '         5 = SAP Object Type as Number (Obj.ObjectTypeAsNumber)
    '         6 = Business Object Type (Criteria Button, Label-Field, Label-To, Low Value, High Value)
    '         7 = Object Base
    Dim curGUIObjectInfo As New vbaSAPGUIObjectInfo
    Dim objTypeSAP
    Dim objName
    Dim BaseName
    Dim baseNameEndPos
    Dim objTypeBus
    
    objName = pObject.Name
    objTypeSAP = pObject.Type
    
    With curGUIObjectInfo
        .ObjectPath = pObject.id
        .ObjectName = objName
        .ObjectValue = pObject.text
        .SAPObjectType = objTypeSAP
        .SAPObjectTypeAsNumber = pObject.TypeAsNumber
    End With
    
    ' Assume GUI objects have the following structure/naming convention
    '  Label:  %_<base>_%_APP_%-TEXT
    '  Low:    <base>-LOW
    '  To label: %_<base>_%_APP_%-TO_TEXT
    '  High:   <base>-HIGH
    '  Criteria button:  %_<base>_%_APP_%-VALU_PUSH
    BaseName = ""
    objTypeBus = ""
    If objTypeSAP = "GuiTextField" Or objTypeSAP = "GuiCTextField" Or (objTypeSAP = "GuiButton" And Right(objName, 9) = "VALU_PUSH") Then
        'Label-Field, Label-To, Low Value, High Value)
        If Left(objName, 2) = "%_" Then
            baseNameEndPos = InStr(3, objName, "%") - 2
            BaseName = Mid(objName, 3, baseNameEndPos - 2)
            If Right(objName, 15) = "%_APP_%-TO_TEXT" Then
                objTypeBus = "Label-To"
            ElseIf Right(objName, 12) = "%_APP_%-TEXT" Then
                objTypeBus = "Label-Field"
            ElseIf Right(objName, 17) = "%_APP_%-VALU_PUSH" Then
                objTypeBus = "Criteria Button"
            End If
        Else
            If Right(objName, 4) = "-LOW" Then
                objTypeBus = "Low Value"
            ElseIf Right(objName, 5) = "-HIGH" Then
                objTypeBus = "High Value"
            End If
            If objTypeBus <> "" Then
                baseNameEndPos = InStr(objName, "-") - 1
                BaseName = Left(objName, baseNameEndPos)
            Else
                BaseName = objName
            End If
        End If
    Else
        objTypeBus = objTypeSAP
        BaseName = objName
    End If
    With curGUIObjectInfo
        .BusinessObjectType = objTypeBus
        .BaseName = BaseName
        .Changeable = pObject.Changeable
    End With
    
    Set parseObject = curGUIObjectInfo

End Function


' Assume session is already set
Function parseObjects(Optional pDispObjects = False)

    Dim guiObjectList As New Dictionary
    
    Dim curGUIObjectInfo As New vbaSAPGUIObjectInfo

    'Dim SAPGUIAuto As Object
    'Dim sapapp As Object
    'Dim sapcon As Object
    'Dim session As Object
    Dim Area As Object
    Dim i As Long
    Dim Children As Object
    Dim obj As Object
    Dim curName As String
    'Dim parsedObjValues() As String

    'Set SAPGUIAuto = GetObject("SAPGUI")
    'Set sapapp = SAPGUIAuto.GetScriptingEngine
    'Set sapcon = sapapp.Children(0)
    'Set session = sapcon.Children(0)
    If session Is Nothing Then
        Set session = getSession(1)
    End If
    ' Currently, hardcode user area of main window.  This can be made soft using an optional parameter
    Set Area = session.FindById("wnd[0]/usr")
    Set Children = Area.Children()
    
    'The variable Children seems to get reset by selecting the clear menu
    'Debug.Print "User Area=" & Area.ID
    For i = 0 To Area.Children.Count() - 1
        Do
            'Set Obj = Children(CInt(i))
            Set obj = Area.Children(CInt(i))
            curName = obj.Name
            Set curGUIObjectInfo = parseObject(obj)
            If pDispObjects Then
                Debug.Print i, "Name=" & obj.Name & "  Value=" & obj.text
                With curGUIObjectInfo
                    Debug.Print .ObjectName, .ObjectPath, .ObjectValue, .SAPObjectType, .BusinessObjectType, .BaseName, .Changeable
                End With
            End If
            guiObjectList.Add i, curGUIObjectInfo
            Exit Do
        Loop
    Next i
    
    Set parseObjects = guiObjectList
    
    Set Children = Nothing
    Set obj = Nothing

End Function


Sub SetValue(pPath, pValue)
    session.FindById(pPath).text = pValue
End Sub

' If pValue = "", then don't set value -- assume it was already cleared
Sub SetValue01(pPath, pValue)
    If pValue = "" Then GoTo finalize
    SetValue pPath, pValue
finalize:
End Sub


' Only supports 1 session
'
' TODO Give error if find more than one session
Function getSession(pLogLevel)

    Dim SapGuiAuto As Object
    Dim App, Connection, session As Object

    ' Initiate SAP Scripting objects
    Dim numTries As Integer
    Dim retCode As Integer
    
    ' Initiate SAP Scripting objects
    LogMsg "Initiating SAP Scripting objects"
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    On Error GoTo 0
    If SapGuiAuto Is Nothing Then
        ' TODO - make startup/logon parameters soft, or at least don't use Accounting
        ' https://archive.sap.com/discussions/thread/3856038
        ' https://stackoverflow.com/questions/16809492/how-do-i-login-through-sap-using-command-line
        ' TODO - turn off log messages, but allow for them to be enabled if we want to debug this process
        LogMsg "Starting SAP"
        'Shell """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe"" -sysname=""[001]   PA1 [Accounting]"" -client=50 -command=" & pReportInfo.TCode
        Shell """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe"" -sysname=""[001]   PA1 [Manufacturing]"" -client=50"
        numTries = 0
        retCode = -1
        Do While retCode <> 0 And numTries < 10
            numTries = numTries + 1
            AppWait 1
            LogMsg "Try #" & numTries
            retCode = checkSAPSessionObjects
        Loop
        If numTries >= 10 Then
            LogError "Unable to initiate SAP GUI"
            Exit Function
        End If
        'Kludge - Wait 1 sec for application to finish initializing
        'TODO search for main menu bar to input tCode
        AppWait 1
        Set SapGuiAuto = GetObject("SAPGUI")
    End If
    
    LogMsg "Initiating existing SAP session"
    Set App = SapGuiAuto.GetScriptingEngine
    If Not IsObject(App) Then
        LogError "Unable to initiate SAP GUI App object"
        Exit Function
    End If
    LogMsg "Got App"
    
    'TODO - if the Logon Pad is running but there are no sessions, the code below will fail.  Clean up the login pad window and create a fresh login.  Can use https://stackoverflow.com/questions/56895430/run-time-error-91-when-sap-connection-is-empty to find open windows.
    '     - Possibly use https://stackoverflow.com/questions/25424469/vba-get-taskbar-applications
    Set Connection = App.Children(0)
    If Not IsObject(Connection) Then
        LogError "Unable to initiate SAP GUI Connection object"
        Exit Function
    End If
    LogMsg "Got Connection"
    
    Set session = Connection.Children(0)
    If Not IsObject(session) Then
        LogError "Unable to initiate SAP GUI Session object"
        Exit Function
    End If
    LogMsg "Got Session"

    Set getSession = session
End Function


Function checkSAPSessionObjects()
        
    Dim SapGuiAuto As Object
    Dim App, Connection, session As Object
    
    Dim gotAllObjects As Boolean
        
    On Error GoTo end_loop
    
    gotAllObjects = False
    
    Set SapGuiAuto = CreateObject("SAPGUI")
    LogMsg "Got SAPGUI"
    
    Set App = SapGuiAuto.GetScriptingEngine
    LogMsg "Got App"
    
    ' Ensure the object's children has been populated
    If App Is Nothing Then GoTo end_loop
    If App.Children Is Nothing Then GoTo end_loop
    If App.Children.Count >= 0 Then
        Set Connection = App.Children(0)
    End If
    LogMsg "Got Connection"
    
    ' Ensure the object's children has been populated
    Debug.Print 1
    If Connection Is Nothing Then GoTo end_loop
    Debug.Print 2
    If Connection.Children Is Nothing Then GoTo end_loop
    Debug.Print 3
    If Connection.Children.Count >= 0 Then
        Debug.Print "4:  #=" & Connection.Children.Count
        Debug.Print "4b:  Connection.children.Type=" & Connection.Children.Type
        If IsEmpty(Connection.Children.Type) Then GoTo end_loop
        Debug.Print 5
        Set session = Connection.Children(0)
    End If
    LogMsg "Got Session"
    
    gotAllObjects = True
    
    On Error GoTo 0
    
end_loop:
    Set SapGuiAuto = Nothing
    Set App = Nothing
    Set Connection = Nothing
    Set session = Nothing
    
    If gotAllObjects Then
        checkSAPSessionObjects = 0
    Else
        checkSAPSessionObjects = 1
    End If
    
    Exit Function 'Must exit or Resume to reset error handler - https://stackoverflow.com/questions/6028288/properly-handling-errors-in-vba-excel

End Function


Function canUpdateObj(pObjTypeNum)
    canUpdateObj = (pObjTypeNum = GuiComponentType.GuiTextField Or pObjTypeNum = GuiComponentType.GuiCTextField)
End Function


'Process:
'1.  Load to clipboard
'2.  Click Paste button
'3.  Close (F8)
Function copyDictToSelectionList(pDict As Dictionary, Optional pCloseWindow = True)
    'Load to clipboard
    loadTextDictToClipboard pDict
    
    ' Click Paste button
    session.FindById("wnd[1]/tbar[0]/btn[24]").Press
    
    ' Close window, saving entered data
    If pCloseWindow Then
        session.FindById("wnd[1]/tbar[0]/btn[8]").Press
    End If
End Function



'TODO detect if not on home screen - give error
Function executeTransaction()
    Debug.Print "executeTransaction - begin"
    ' Execute Transaction
        session.FindById("wnd[0]").SendVKey 8
    Debug.Print "executeTransaction - end"
        
End Function



Sub initReportInfo(pReportInfo As reportInfoType)

With pReportInfo
    If .outputFolder = "" Then
        'TODO Get from registry or other source
        .outputFolder = "c:\users\" & getCurUser & "\Documents\SAP\SAP GUI"
    End If
    
    If .outputFileName = "" Then
        .outputFileName = "export.xlsx"
    End If
    
    .outputFilePath = .outputFolder & "\" & .outputFileName
    
End With

End Sub



Function exportData(pReportInfo As reportInfoType, Optional pExportType = "ALV-GRID-EXCEL", Optional pOutputFormat = "unconverted", Optional pTimeout = 600)

    'TODO handle case where no data is found

    Dim retCode
    
    With pReportInfo
    
        initReportInfo pReportInfo
        Debug.Print "exportData - begin.  Output file=" & .outputFilePath
        
        'TODO handle error if file can't be deleted
        If Dir(.outputFilePath) > "" Then Kill .outputFilePath
        
        If session Is Nothing Then Set session = getSession(1)
        
        'If IsMissing(pExportType) Then
        '    Set sessionInfo = session.Info
        '    pExportType = sessionInfo.Program
        'End If
    
        Select Case pExportType
            'ALV grid control
            Case "ALV-GRID"
                
                ' Invoke export via the List menu
                ' TODO Make selection soft - select by menu name, not position
                session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
                setOutputFormat01 pOutputFormat
                
                ' Insert File Path into "Directory" field
                session.FindById("wnd[1]/usr/ctxtDY_PATH").text = .outputFolder
                'Insert File Name into "File Name" field
                session.FindById("wnd[1]/usr/ctxtDY_FILENAME").text = .outputFileName
                
            
                ' Replace File
                session.FindById("wnd[1]/tbar[0]/btn[11]").Press
            
            ' Used by SQVI
            Case "SAPLAQRUNT"
                
                ' Invoke export
                session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
                session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").SelectContextMenuItem "&PC"
                session.FindById("wnd[1]/tbar[0]/btn[0]").Press
                
                ' Insert File Path into "Directory" field
                session.FindById("wnd[1]/usr/ctxtDY_PATH").text = .outputFolder
                'Insert File Name into "File Name" field
                session.FindById("wnd[1]/usr/ctxtDY_FILENAME").text = .outputFileName
            
                ' Replace File
                session.FindById("wnd[1]/tbar[0]/btn[11]").Press
                
            ' ALV grid control - Excel - invoke via "List" menu
            '
            Case "ALV-GRID-EXCEL"
                getALVGridExcel pReportInfo, pTimeout
                
            Case "LIST-SPREADSHEET"
                getListSpreadsheet pReportInfo, pTimeout
                
        End Select
            
            
        'session.FindById("wnd[0]/tbar[0]/btn[3]").Press
        'session.FindById("wnd[0]").SendVKey 3
        
finalize:
        ' Since we deleted the file at the beginning of the routine, if it exists, assume success
        If Dir(.outputFilePath) > "" Then
            retCode = 0
        Else
            retCode = 1
        End If
    End With 'pReportInfo
    
    exportData = retCode
    
    Debug.Print "exportData - end.  Code=" & retCode

End Function

' Sets radio set in prompt for data format from ALV Grid - "Local File" export
Sub setOutputFormat01(Optional pOutputFormat = "unconverted", Optional pExecute = True)
    Dim id
    Dim idBase
    Dim idFull
    
    idBase = "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/"
    Select Case UCase(pOutputFormat)
        Case "UNCONVERTED"
            id = "radSPOPLI-SELFLAG[0,0]"
        Case "SPREADSHEET"
            id = "radSPOPLI-SELFLAG[1,0]"
        Case "RICH TEXT"
            id = "radSPOPLI-SELFLAG[2,0]"
        Case "HTML"
            id = "radSPOPLI-SELFLAG[3,0]"
        Case "CLIPBOARD"
            id = "radSPOPLI-SELFLAG[4,0]"
    End Select
    idFull = idBase & id
    session.FindById(idFull).Select
    
    If pExecute Then
        session.FindById("wnd[1]/tbar[0]/btn[0]").Press
    End If
End Sub


Function getALVGridExcel(pReportInfo As reportInfoType, Optional pTimeout = 600)
    
    Dim saveAsTitle As String
    Dim retCode As Integer
    
    saveAsTitle = "Save As" 'TODO Support other languages
    
    With pReportInfo
        If windowExists(saveAsTitle) Then
            LogError "'Save As' dialog is already open.  Search title=" & saveAsTitle
            GoTo finalize
        End If
        
        ' Run export
        '
        ' Notes:
        ' 1.  This export method blocks for processing until the "Save As" dialog is cleared.
        '     Launch asynchronous process to interact with "Save As" dialog
        '
        ' 2.  Choose via menu bar to make more generic than a context menu in the grid.  E.g. FBL1N has a different path than MB51.
        '     Also avoids the need to select export type
        
        'TODO try a variety of ways to export to SAP
        '     - If can find the grid control, use the context menu
        '     - If the first menu has an option called "Spreadsheet", select it.  However, IQ09 has a number of steps
        'TODO make path to VBS soft
        'TODO ** use different timeout for SaveAs and closing Excel workbook
        Shell ("cscript ""\\ussomgensvm00.allergan.com\lifecell\Depts\Tissue Services\TS Field\bat\drive_save_as_dialog.vbs"" """ & .outputFilePath & """")
        session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
        
        'Set obj = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
        'obj.SetCurrentCell -1, obj.ColumnOrder(0)
        'obj.ContextMenu
        'obj.SelectContextMenuItem "&XXL"
        'retCode = waitForWindow(saveAsTitle)
        'If retCode <> 0 Then
        '    LogError "Unable to find 'Save As' dialog.  Search title=" & saveAsTitle & ".  Return code=" & retCode
        '    GoTo finalize
        'End If
        'SendKeys "%n"
        'SendKeys .outputFilePath
        'SendKeys "%n"
        
        'Wait for Excel to launch
        retCode = waitForExcelThenClose(pReportInfo)
        If retCode <> 0 Then
            GoTo finalize
        End If
    End With
    
    retCode = 0
    
finalize:
    getALVGridExcel = retCode

End Function


Function getListSpreadsheet(pReportInfo As reportInfoType, Optional pTimeout = 600)
    
    Dim saveAsTitle As String
    Dim retCode As Integer
    
    saveAsTitle = "Save As" 'TODO Support other languages
    
    With pReportInfo
        If windowExists(saveAsTitle) Then
            LogError "'Save As' dialog is already open.  Search title=" & saveAsTitle
            GoTo finalize
        End If
        
        ' Run export
        '
        ' Notes:
        ' 1.  This export method blocks for processing until the "Save As" dialog is cleared.
        '     Launch asynchronous process to interact with "Save As" dialog
        '
        ' 2.  Choose via menu bar to make more generic than a context menu in the grid.  E.g. FBL1N has a different path than MB51.
        '     Also avoids the need to select export type
        
        'TODO try a variety of ways to export to SAP
        '     - If can find the grid control, use the context menu
        '     - If the first menu has an option called "Spreadsheet", select it.  However, IQ09 has a number of steps
        'TODO make path to VBS soft
        'TODO ** use different timeout for SaveAs and closing Excel workbook
        Shell ("cscript ""\\ussomgensvm00.allergan.com\lifecell\Depts\Tissue Services\TS Field\bat\drive_save_as_dialog.vbs"" """ & .outputFilePath & """")
        session.FindById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").Select
        
        'Set obj = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
        'obj.SetCurrentCell -1, obj.ColumnOrder(0)
        'obj.ContextMenu
        'obj.SelectContextMenuItem "&XXL"
        'retCode = waitForWindow(saveAsTitle)
        'If retCode <> 0 Then
        '    LogError "Unable to find 'Save As' dialog.  Search title=" & saveAsTitle & ".  Return code=" & retCode
        '    GoTo finalize
        'End If
        'SendKeys "%n"
        'SendKeys .outputFilePath
        'SendKeys "%n"
        
        'Wait for Excel to launch
        retCode = waitForExcelThenClose(pReportInfo)
        If retCode <> 0 Then
            GoTo finalize
        End If
    End With
    
    retCode = 0
    
finalize:
    getListSpreadsheet = retCode

End Function


'TODO Display status of background processes
'TODO Log entry and exit of various subroutines/functions, or otherwise provide call stack trace
'TODO If more than one instance of Excel is found, check all of them, or give an error
Function waitForExcelThenClose(pReportInfo As reportInfoType)
    Dim retStx
    Dim retCode
    Dim xclWkb
    Dim xclApp
    
    With pReportInfo
        ' Get workbook
        retStx = waitForExcelWorkbook(.outputFileName)
        retCode = retStx(1)
        If retCode <> 0 Then
            LogError "Unable to find Excel file '" & .outputFileName & "' within timeout=" & retStx(3) & ".  Return code=" & retCode
            GoTo finalize
        End If
    
        ' Close workbook
        retCode = 99
        On Error GoTo close_wb_error
        Set xclWkb = retStx(2)
        Set xclApp = xclWkb.Application
        xclWkb.Close False
        xclApp.Application.Quit
        retCode = 0
        GoTo next_step
    
close_wb_error:
        If err.Number > 0 Then
            LogError01
        Else
            LogError "Error raised when attempting to close workbook " & .outputFileName
        End If
    
next_step:
    End With

finalize:
    waitForExcelThenClose = retCode
End Function

Function exitToHomeScreen()

Debug.Print "Exiting to home screen"
    ' TODO:  repeat sendVKey until detect reaching Home Screen
        session.FindById("wnd[0]").SendVKey 12

        session.FindById("wnd[0]").SendVKey 12

End Function


'----------------------------------------------------------------------------
'setGUIValues - set GUI objects with user-provided values
'
'Objective:  For each parameter in the "reportInfo" object, set all matching GUI objects to the values in the parameter
'
'Example (MB51)
'         Material   _________  to _________  <Button>
'         1          2          3  4          5
'
' #  Item name                    Value          Base Item
' -  ---------------------------- -------------- ---------
' 1  %_MATNR_%_APP_%-TEXT         Material       MATNR
' 2  MATNR-LOW                                   MATNR
' 3  %_MATNR_%_APP_%-TO_TEXT      to             MATNR
' 4  MATNR-HIGH                                  MATNR
' 5  %_MATNR_%_APP_%-VALU_PUSH                   MATNR
'
'Details:
'- The system first clears all GUI elements in the usr portion of the object hierarchy
'
'- A parameter is matched with one or more GUI objects.  In the parameter, a user supplies a label and/or a base item.
'
'    1.  User specifies a label
'        * Given a label, this process matches GUI objects which:
'          - Have the same label as the parameter, or
'          - Are a member of the same item group as a GUI object which matches by label
'        * Some GUI screens have more than one label with the same value.
'          - The process will match a parameter to only one of those labels, and any applicable objects which are in the same item group as that label.
'          - Additional parameters can be specified to match additional instances of a label value.
'
'    2.  User specifies a base item
'        * Label matching is bypassed.  Instead, GUI objects are selected using the base item
'        * If both a label and a base item are specified, the label is ignored
'
'- A parameter can have several values (Low, High, lists of items to include/exclude).  These are matched with GUI objects using object naming conventions.  See code below for details.
'----------------------------------------------------------------------------
Function setGUIValues(pRptInfo As reportInfoType, Optional pLogLevel = 0)
    'TODO - If can't find any of the given parameters, give an error.
    Dim guiObjectList As New Dictionary
    Dim curObj As vbaSAPGUIObjectInfo
    Dim curObj2 As vbaSAPGUIObjectInfo
    Dim curParam As vbaSAPParamValues
    Dim curObjIdx, curObj2Idx, curParamIdx
    Dim itemHasSetableFields As Boolean
    Dim setItemDirectly As Boolean
    Dim setOtherValue As Boolean
    Dim itemList As New Collection
    Dim userSuppliedBaseItem As Boolean
    
    Set guiObjectList = parseObjects
    
    LogMsg01 "Clearing fields"
    ClearAllFields
    
    ' curObj - Object linked to the parameter, based on label or base item
    ' curObj2 - Related objects using base item
    '
    LogMsg01 "Setting values"
    For Each curParamIdx In pRptInfo.params
        Set curParam = pRptInfo.params(curParamIdx)
        If pLogLevel > 0 Then LogMsg "Param:  Param label=" & curParam.LabelName & ", base item=" & curParam.BaseName & ", low=" & curParam.LowValue
        
        userSuppliedBaseItem = (curParam.BaseName > "")
        itemHasSetableFields = False
        'Loop through all GUI objects, looking for a match with the current parameter, using base item if specified, or label
        For Each curObjIdx In guiObjectList
            Set curObj = guiObjectList(curObjIdx)
            
            If pLogLevel >= 5 Then LogMsg "Examining obj: Obj Name=" & curObj.ObjectName & ", SAP Type=" & curObj.SAPObjectType & ", Bus Type=" & curObj.BusinessObjectType & ", Value=" & curObj.ObjectValue
            
            '1.  Only look at GUI elements which contain labels
            Select Case curObj.BusinessObjectType
                Case "Label-Field"
                    setItemDirectly = False
                Case "GuiRadioButton", "GuiCheckBox"
                    setItemDirectly = True
                Case Else
                    If Not userSuppliedBaseItem Then GoTo end_loop01
            End Select
            
            If pLogLevel >= 7 Then LogMsg "setItemDirectly=" & setItemDirectly & ", userSuppliedBaseItem=" & userSuppliedBaseItem
            
            '2.  If the user supplied a base item, skip all objects that do not match that item
            '3.  If the user did not supply a base item, but a base item is already set, then we already processed an object matching the label or base item, so exit
            If userSuppliedBaseItem Then
                If curObj.BaseName <> curParam.BaseName Then GoTo end_loop01
                If pLogLevel >= 10 Then LogMsg "After base name check - user supplied"
            Else
                'If the base name of the param is set by this process, we already processed the parameter, so don't process it again
                If curParam.BaseName > "" Then Exit For
                If pLogLevel >= 10 Then LogMsg "After base name check - curParam"
                
                'Try to match on static labels
                If curObj.ObjectValue <> curParam.LabelName Then GoTo end_loop01
                If pLogLevel >= 10 Then LogMsg "After check label match"
                
                If Not setItemDirectly And curObj.Changeable Then GoTo end_loop01
                If pLogLevel >= 10 Then LogMsg "After check changeable"
                
                ' If base item has already been processed (matched on another instance of the label), skip
                If ExistsInCollection(itemList, curObj.BaseName) Then GoTo end_loop01
                If pLogLevel >= 10 Then LogMsg "After check already processed"
            End If
                        
            If pLogLevel >= 7 Then LogMsg "After checks 2 and 3"
            
            '4.  Ensure at least one object for the given item is changeable
            For Each curObj2Idx In guiObjectList
                Set curObj2 = guiObjectList(curObj2Idx)
                If curObj2.BaseName <> curObj.BaseName Then GoTo end_loop03
                If Not curObj2.Changeable Then GoTo end_loop03
                itemHasSetableFields = True
                Exit For
end_loop03:
            Next
            If pLogLevel >= 10 Then LogMsg "ItemHasSetableFields=" & itemHasSetableFields
            If Not itemHasSetableFields Then GoTo end_loop01
            
            If pLogLevel > 1 Then LogMsg "Found obj:  Param label=" & curParam.LabelName & ", Matched Obj Name=" & curObj.ObjectName & ", SAP Type=" & curObj.SAPObjectType & ", Bus Type=" & curObj.BusinessObjectType & ", Value=" & curObj.ObjectValue
            
            If setItemDirectly Then
                Select Case curObj.BusinessObjectType
                    Case "GuiRadioButton"
                        session.FindById(curObj.ObjectPath).Select
                    Case "GuiCheckBox"
                        'Assume if we pass a parameter to a checkbox object, the parameter datatype is Boolean
                        If IsEmpty(curParam.LowValue) Then
                            session.FindById(curObj.ObjectPath).Selected = Not session.FindById(curObj.ObjectPath).Selected
                        Else
                            session.FindById(curObj.ObjectPath).Selected = curParam.LowValue
                        End If
                    Case Else
                        LogError "Invalid BusinessObjectType=" & curObj.BaseName & " found in the setItemDirectly portion of setGUIValues"
                        GoTo end_loop01
                End Select
                If pLogLevel > 1 Then Debug.Print "Set directly - name=" & curObj.ObjectName
                curParam.BaseName = curObj.BaseName
                itemList.Add curParam.BaseName, curParam.BaseName
                'GoTo end_loop04 'Next parameter
                Exit For 'Next parameter
            End If
            
            '--------------------------------------------------------------
            'From this point on, process all objects in the selected group
            '--------------------------------------------------------------
            If pLogLevel > 1 Then Debug.Print "Set " & curParam.LabelName & ", base name=" & curObj.BaseName & ", curObj Name=" & curObj.ObjectName
            
            'Set the needed values
            '
            'If the parameter has a Low Value, and a matching object (on BaseName) has no BusinessObjectType and is changeable,
            'consider this object as "Other" and set it.  Only do so for the first such object found.
            setOtherValue = False
            For Each curObj2Idx In guiObjectList
                Set curObj2 = guiObjectList(curObj2Idx)
                If pLogLevel >= 10 Then LogMsg "Set object base name=" & curObj2.BaseName & ", curObj Name=" & curObj2.ObjectName & ", curObj Value=" & curObj2.ObjectValue
                
                If curObj2.BaseName <> curObj.BaseName Then GoTo end_loop02
                
                'Set low, high, etc
                If pLogLevel >= 7 Then LogMsg "Setting object base name=" & curObj2.BaseName & ", curObj Name=" & curObj2.ObjectName & ", curObj Value=" & curObj2.ObjectValue
                Select Case curObj2.BusinessObjectType
                    Case "Label-Field", "Label-To"
                        GoTo end_loop02
                    Case "Low Value"
                        SetValue01 curObj2.ObjectPath, curParam.LowValue
                        setOtherValue = True
                    Case "High Value"
                        SetValue01 curObj2.ObjectPath, curParam.HighValue
                    Case "Criteria Button"
                        If curParam.IncludeDict Is Nothing And curParam.ExcludeDict Is Nothing Then GoTo end_loop02
                        session.FindById(curObj2.ObjectPath).Press
                        If Not curParam.IncludeDict Is Nothing Then
                            If curParam.IncludeDict.Count > 0 Then
                                'Select "Single Values" tab
                                session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA").Select
                                copyDictToSelectionList curParam.IncludeDict, False
                            End If
                        End If
                        If Not curParam.ExcludeDict Is Nothing Then
                            If curParam.ExcludeDict.Count > 0 Then
                                'Select Exclude Single Values tab
                                session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
                                copyDictToSelectionList curParam.ExcludeDict, False
                            End If
                        End If
                        ' Close window, saving entered data
                        session.FindById("wnd[1]/tbar[0]/btn[8]").Press
                    Case Else '"Other" fields
                        If curParam.LowValue <> "" And Not setOtherValue And curObj2.Changeable Then
                            SetValue01 curObj2.ObjectPath, curParam.LowValue
                            setOtherValue = True
                        End If
                End Select
end_loop02:
            Next 'GUI Objects with matching item code (CurObj2)
            
            If pLogLevel >= 3 Then Debug.Print "Finished setting values for " & curObj.BaseName
            itemList.Add curObj.BaseName, curObj.BaseName
            If IsEmpty(curParam.BaseName) Then
                curParam.BaseName = curObj.BaseName
            ElseIf Not userSuppliedBaseItem Then
                LogError "Processed curObj.BaseName=" & curObj.BaseName & " but curParam.BaseName=" & curParam.BaseName
            End If
            
            Exit For
end_loop01:
        Next 'All GUI Objects (CurObj)
end_loop04:
    Next 'Parameter
    LogMsg01 "Process complete"
End Function

'========================================================

'--------------------------------------------------
'GUI scripting template
'--------------------------------------------------
Function extractXXXXX()
    
    ' Base definitions
    Dim curReportInfo As reportInfoType
    Dim session
    Dim retCode
    
    With curReportInfo
    
        'Set parameters
        .TCode = "XXXXX"
        setParam .params, "Material", "1520320"
        setParam .params, "Equipment", "RH200*"
        initReportInfo curReportInfo
    
        'Get session
        Set session = InitSAP(curReportInfo)
        
        'Load parameter values into SAP GUI
        setGUIValues curReportInfo
        
        'Execute process
        executeTransaction
        
        'Export data
        retCode = exportData(curReportInfo)
        
        'Exit to home screen
        exitToHomeScreen
        
        If retCode <> 0 Then
            LogError "Error exporting data"
            GoTo finalize
        End If
        
    End With
    
finalize:
    extractXXXXX = retCode

End Function
