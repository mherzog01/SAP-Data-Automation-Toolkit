Attribute VB_Name = "A1_Download_SAP_Data_Functions"
Dim session

Sub Initiate_SAP()

Set session = getSession(1)

End Sub

Sub Close_SAP_Session()

    session.FindById("wnd[0]").Close
    session.FindById("wnd[1]/usr/btnSPOP-OPTION1").Press
    
    'cmd = "cscript ""\\ussomgensvm00.allergan.com\LifeCell\Depts\Planning Meetings\Alert List Priority Lot Summary\bat\send_key.vbs"" ""SAP Logon Pad 740"" ""%{F4}"""
    'Shell cmd

    KillProcessbyName "saplgpad.exe"
    KillProcessbyName "saplogon.exe"

End Sub

Sub Download_Snap_Data(Start_Date As Date, End_Date As Date, Export_Path As String, Export_File As String)

    ' Initiate SAP Scripting once open
    
        Set SapGuiAuto = GetObject("SAPGUI")
        If Not IsObject(SapGuiAuto) Then
            Exit Sub
        End If
        
        Set App = SapGuiAuto.GetScriptingEngine
        If Not IsObject(App) Then
            Exit Sub
        End If
        
        Set Connection = App.Children(0)
        If Not IsObject(Connection) Then
            Exit Sub
        End If
        
        Set session = Connection.Children(0)
        If Not IsObject(session) Then
            Exit Sub
        End If
    
    ' SQ01
    
        ' Handle notification alert
    
            On Error Resume Next
            
                session.FindById("wnd[1]/tbar[0]/btn[12]").Press
        
        ' Insert Transaction
            session.FindById("wnd[0]/tbar[0]/okcd").text = "SQ01"
            
        ' Launch Transaction
            session.FindById("wnd[0]/tbar[0]/btn[0]").Press
            
        ' Select Standard Area
            session.FindById("wnd[0]/mbar/menu[5]/menu[0]").Select
            session.FindById("wnd[1]/usr/radRAD1").Select
            session.FindById("wnd[1]/tbar[0]/btn[2]").Press
            
        ' Filter for BBG_PLANNING Group
            session.FindById("wnd[0]/tbar[1]/btn[19]").Press
            session.FindById("wnd[1]/tbar[0]/btn[29]").Press
            session.FindById("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").SelectedRows = "0"
            session.FindById("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/btnAPP_WL_SING").Press
            session.FindById("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON").Press
            session.FindById("wnd[3]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").Press
            session.FindById("wnd[4]/tbar[0]/btn[16]").Press
            session.FindById("wnd[4]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "BBG_PLANNING"
            session.FindById("wnd[4]/tbar[0]/btn[8]").Press
            session.FindById("wnd[3]/tbar[0]/btn[0]").Press
            session.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").SelectedRows = "0"
            session.FindById("wnd[1]/tbar[0]/btn[0]").Press
            
            
        ' Insert Query Name
            session.FindById("wnd[0]/usr/ctxtRS38R-QNUM").text = "MB51-SER-TIME"
            session.FindById("wnd[0]/tbar[1]/btn[8]").Press

            
    ' Number of Material Document (Selection Field 1)
            
        ' Open Multiple Selection for Number of Material Document (Selection Field 1)
            session.FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH").Press

        ' Delete Entire Selection Line command (Shift + F4)
            session.FindById("wnd[1]/tbar[0]/btn[16]").Press
        
        ' Copy Multiple Selection for Number of Material Document
            session.FindById("wnd[1]/tbar[0]/btn[8]").Press
            
    ' Number of Material Document Year (Selection Field 2)
            
        ' Open Multiple Selection for Number of Material Document Year (Selection Field 2)
            session.FindById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH").Press
             
        ' Delete Entire Selection Line command (Shift + F4)
            session.FindById("wnd[1]/tbar[0]/btn[16]").Press
            
        ' Copy Multiple Selection for Number of Material Document Year
            session.FindById("wnd[1]/tbar[0]/btn[8]").Press
            
    ' Movement Type (Inventory Management) (Selection Field 3)
    
        ' Open Multiple Selection for Movement Type (Inventory Management) (Selection Field 3)
            session.FindById("wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH").Press
            
        ' Delete Entire Selection Line command (Shift + F4)
            session.FindById("wnd[1]/tbar[0]/btn[16]").Press
            
        ' Enter Relevant Movement Types into multiple selection fields
            session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "101"
            session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "102"
            
        ' Copy Movement Type (Inventory Management) (Selection Field 3)
            session.FindById("wnd[1]/tbar[0]/btn[8]").Press

    ' Material (Selection Field 4)
    
        ' Open Material (Selection Field 4)
            session.FindById("wnd[0]/usr/btn%_SP$00004_%_APP_%-VALU_PUSH").Press
        
        ' Delete Entire Selection Line command (Shift + F4)
            session.FindById("wnd[1]/tbar[0]/btn[16]").Press
                
        ' Enter Intermediate Wildcard into Select Single Values tab
            session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "*-I*"
    
        ' Select Exclude Single Values tab
            session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            
        ' Enter Exclusion Wildcard into Exclude Single Values tab
            session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "0000*"
        
        ' Copy Material (Selection Field 4)
            session.FindById("wnd[1]/tbar[0]/btn[8]").Press

    ' Plant (Selection Field 5)
    
        ' Open Multiple Selection for Plant (Selection Field 5)
            session.FindById("wnd[0]/usr/btn%_SP$00005_%_APP_%-VALU_PUSH").Press
            
        ' Delete Entire Selection Line command (Shift + F4)
            session.FindById("wnd[1]/tbar[0]/btn[16]").Press
            
        ' Enter Plant 9541 (LifeCell, BBG) into multiple selection fields
            session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "9541"
            
        ' Copy Plant (Selection Field 5)
            session.FindById("wnd[1]/tbar[0]/btn[8]").Press
            
    ' Storage Location (Selection Field 6)
    
        ' Open Multiple Selection for Storage Location (Selection Field 7)
            session.FindById("wnd[0]/usr/btn%_SP$00006_%_APP_%-VALU_PUSH").Press
            
        ' Delete Entire Selection Line command (Shift + F4)
            session.FindById("wnd[1]/tbar[0]/btn[16]").Press
            
        ' Enter Storage Location 9542 (AlloDerm) into multiple selection fields
            session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "9542"
            
        ' Copy Plant (Selection Field 8)
            session.FindById("wnd[1]/tbar[0]/btn[8]").Press

    ' Posting Date (Selection Field 7)
    
        ' Open Multiple Selection for Storage Location (Selection Field 7)
            session.FindById("wnd[0]/usr/btn%_SP$00007_%_APP_%-VALU_PUSH").Press
            
        ' Delete Entire Selection Line command (Shift + F4)
            session.FindById("wnd[1]/tbar[0]/btn[16]").Press
            
        ' Copy Posting Date (Selection Field 6)
            session.FindById("wnd[1]/tbar[0]/btn[8]").Press
            
        ' Enter Start Date into Posting Date "From" Cell
            session.FindById("wnd[0]/usr/ctxtSP$00007-LOW").text = Start_Date
        
        ' Enter End Date into Posting Date "to" Cell
            session.FindById("wnd[0]/usr/ctxtSP$00007-HIGH").text = End_Date
            
    ' Batch (Selection Field 8)
    
        ' Open Batch (Selection Field 8)
            session.FindById("wnd[0]/usr/btn%_SP$00008_%_APP_%-VALU_PUSH").Press
        
        ' Delete Entire Selection Line command (Shift + F4)
            session.FindById("wnd[1]/tbar[0]/btn[16]").Press
                
        ' Enter RTU Batch Wildcard into Select Single Values tab
            session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "RH*"
        
        ' Copy Batch Selection
            session.FindById("wnd[1]/tbar[0]/btn[8]").Press
            
    ' Serial Number (Selection Field 9)
    
        ' Open Serial Number (Selection Field 9)
            session.FindById("wnd[0]/usr/btn%_SP$00009_%_APP_%-VALU_PUSH").Press
        
        ' Delete Entire Selection Line command (Shift + F4)
            session.FindById("wnd[1]/tbar[0]/btn[16]").Press
                
        ' Enter RTU Serial Number Wildcard into Select Single Values tab
            session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "RH*"
        
        ' Copy Serial Number Selection
            session.FindById("wnd[1]/tbar[0]/btn[8]").Press
            
            
    ' Execute Transaction
        session.FindById("wnd[0]").SendVKey 8
        
    ' Export as Unconverted Local File
        session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
        session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").SelectContextMenuItem "&PC"
        
        ' Continue
            session.FindById("wnd[1]/tbar[0]/btn[0]").Press
            
        ' Insert File Path into "Directory" field
            session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Export_Path
        
        'Insert File Name into "File Name" field
            session.FindById("wnd[1]/usr/ctxtDY_FILENAME").text = Export_File
        
        ' Replace File
            session.FindById("wnd[1]/tbar[0]/btn[11]").Press

    ' Exit Output to Transaction Input
        session.FindById("wnd[0]").SendVKey 12
        
    ' Exit Transaction Input to SQVI
        session.FindById("wnd[0]").SendVKey 12
        
    ' Exit SQVI to Home Screen
        session.FindById("wnd[0]").SendVKey 12
            
End Sub

