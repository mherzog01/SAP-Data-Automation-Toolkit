# SAP-Data-Automation-Toolkit (SDAT)
Use the tools in this project to automate SAP scripting and RFC calls.  Current clients include Microsoft VBA libraries for Excel and Access.

**SAP RFC Example**

The following code uses SDAT libraries to extract 30 days of data from the table MSEG and insert it into a MS Access table called [Stg - MSEG]:

```VBA
Sub RefreshMSEGStg()

    Dim retCode
    Dim fieldList
    
    DoCmd.SetWarnings False
    
    Dim aFilter(4)
    aFilter(1) = "WERKS = '9541'"
    aFilter(2) = " AND (BWART LIKE '1%' or BWART LIKE '2%' or BWART LIKE '5%')"
    aFilter(3) = " AND BUDAT_MKPF between '" & formatSAPDate(now - 30) & "' AND '" & formatSAPDate(now) & "'"
    aFilter(4) = " AND MATNR in ('111P5050','111P0001-002')"
    
    fieldList = "MANDT,MBLNR,MJAHR,ZEILE,BWART,MATNR,WERKS,LGORT,CHARG,INSMK,ZUSTD,SOBKZ,LIFNR,SHKZG,WAERS,DMBTR,BNBTR,BUALT,SHKUM,MENGE,MEINS,ERFMG,ERFME,EBELN,EBELP,KOSTL,BUKRS,LGNUM,LGTYP,LGPLA,BESTQ,BWLVS,TBNUM,TBPOS,XBLVS,PRCTR,SAKTO,VFDAT,BUSTM,BUSTW,HSDAT,ZUSTD_T156M,VGART_MKPF,BUDAT_MKPF,CPUDT_MKPF,CPUTM_MKPF,USNAM_MKPF,XBLNR_MKPF,TCODE2_MKPF,SGTXT,GRUND"
    
    retCode = transferSAPTableDataToAccess("MSEG", fieldList, aFilter, "Stg - MSEG", True)
    
    DoCmd.SetWarnings True

End Sub
```

**SAP Scripting Example**

The following code uses SDAT libraries to run the transaction FBLN for a list of vendors and export the result to a text file.

```vba
Sub extractFBL1_OpenItems()
    ' Base definitions
    Dim curReportInfo As reportInfoType
    Dim session
    
    ' Definitions for parameters
    Dim dIncludeList As New Dictionary
    
    curReportInfo.TCode = "FBL1N"
    
    Set session = InitSAP(curReportInfo)
    
    ' Vendor account
    dIncludeList.Add "1000000", "1000000"
    dIncludeList.Add "1000001", "1000001"
    dIncludeList.Add "1000002", "1000002"
    dIncludeList.Add "1000003", "1000003"
    dIncludeList.Add "1000004", "1000004"
    dIncludeList.Add "1000005", "1000005"
    dIncludeList.Add "1000006", "1000006"
    dIncludeList.Add "1000007", "1000007"
    dIncludeList.Add "1000008", "1000008"
    dIncludeList.Add "1000009", "1000009"
    dIncludeList.Add "1000010", "1000010"
    dIncludeList.Add "1000011", "1000011"
    dIncludeList.Add "1000012", "1000012"
    setParam curReportInfo.params, "Vendor account", pIncludeDict:=dIncludeList
    setParam curReportInfo.params, "Company code", "0141"
    
    setGUIValues curReportInfo
    executeTransaction
    exportData curReportInfo
    exitToHomeScreen

End Sub
```

**Installation**
1.  RFC<br>
1.1  Modify the procedure getSAPTableData in the appSAPRFCLibrary with your RFC connection parameters.<br>
1.2  Set up VBA references listed in module header comments

2.  Scripting<br>
2.1  Modify the procedure getSession in the appSAPScriptingLibrary with your -sysname name.<br>
2.2  Set up VBA references listed in module header comments
