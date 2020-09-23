VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDBSQLUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Class: CDBSQLUtil
'Author: J. Kamau
'Date: 4th Feb 2002
''''''''''''''Notes:'''''''''''''''''
'There are some bugs as regards dimming the sql objects Backup and Restore.
'If you create them with the WithEvents keyword you shall get an error
'
'
''''''''''''''New notes as of 23-10-2002-12.41am'''''''
'
'I've wisened up since I wrote this prj and there are some things
'wouldbe downloaders should consider.
'
'1. The SHBrowseForFolder function may have a bug in winNT 4.0 whereby
'   it does not allow you to select a valid folder, if you find this bug
'   just change the flags and it should be ok.
'2. The restore section of the class does not work very well,
'   its BADLY written and BADLY designed, I didn't get round to finishing it
'   back in february. Therefore don't rely on it, just use it for info on how
'   to perform a backup.
'3. I've removed some OLD code for backing up an access database using fso,
'   The class is still there in the synDBUtil dll however, it's not been used
'   anywhere in the UI.
'4. I've refused backup locations that contains spaces, you can remove it by
'   commenting out the section in the BackupDatabase() method. Note that I tested
'   it out on Windows 2000 Advanced Server with SQL Server 7.0 and I kept getting errors.
'5. Lastly, If I got the chance to re-write this prj again. There are many things I'd
'   do differently. I have wisened up much since I started it in Feb.
'
'''''''''''''''''''''''''''''''''''''
Private mobjSQLServer As SQLDMO.SQLServer 'Main SQL server object
Private mobjBackup  As SQLDMO.Backup 'Backup object
Private mobjRestore As SQLDMO.Restore 'Restore object
''''''''''''''''''''''''
'These events have been removed due to a funny bug where you recieve an error
'on classInit
'Public Event BackupPercentageCompleted(ByVal strMessage As String, ByVal intPercentComplete)
'Public Event RestorePercentageCompleted(ByVal strMessage As String, ByVal intPercentComplete)
''''''''''''''''''''''''
Private Const M_MODULENAME  As String = "CDBSQLUtil"
Private mblnConnected As Boolean 'Are we connected or not? This flag says it all
'''''''''''''''''''''''''''''''''
Private mstrBackupFileLocation As String 'Location of backup file
Private mstrBackupFileName As String 'Name of backup file
Private mstrBackupSetName As String 'Backup set name
Private mstrDatabaseName As String 'db name
''''''''''''''''''''''''''''''''''
Private Const G_JOBNAME As String = "Backup of "
Private Const G_SCHEDULE_NAME As String = "Backup Job Schedule"
Private Const G_BACKUP_STEP_NAME As String = "Perform Full Database Backup"
Private Const G_JOBDESCRIPTION As String = "Complete Database Backup"
Private Const G_JOBCATEGORY = "Database Maintenance"
'''''''''''''''''''''''''''''''''''
Public Enum enFrequencyTypes
'Map directly to SQLDMOFreq_Monthly, SQLDMOFreq_Weekly and SQLDMOFreq_Daily
    BackupMonthly = 16
    BackupWeekly = 8
    BackupDaily = 4
End Enum
Public Enum enLoginType
    LoginNormal = SQLDMO_LOGIN_TYPE.SQLDMOLogin_Standard
    LoginNT = SQLDMO_LOGIN_TYPE.SQLDMOLogin_NTUser
End Enum
''''''''''''''''''''''''''''''''''

Public Property Get BackupFileLocation() As String
    BackupFileLocation = mstrBackupFileLocation
End Property

Public Property Let BackupSetName(ByVal strBackupSetName As String)
    mstrBackupSetName = strBackupSetName
End Property

Public Property Let DatabaseName(ByVal strDatabaseName As String)
    mstrDatabaseName = strDatabaseName
End Property

Public Function GetListOfSQLServers() As Variant
    Dim vntList As Variant
    Dim nameList As nameList
    Dim objSQLApp As SQLDMO.Application
    Dim i As Long
    
    Set objSQLApp = New Application
    
    Set nameList = objSQLApp.ListAvailableSQLServers()
    
    If nameList.Count > 0 Then
        ReDim vntList(nameList.Count)
        For i = 0 To UBound(vntList)
            vntList(i) = nameList.Item(i)
        Next i
    End If
    
    GetListOfSQLServers = vntList
    
    Set nameList = Nothing
    Set objSQLApp = Nothing
End Function

Public Function GetDefaultBackupLocation() As String
 On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function GetDefaultBackupLocation"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    If mblnConnected Then
        GetDefaultBackupLocation = mobjSQLServer.Registry.SQLDataRoot & "\BACKUP\"
    End If
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next

Out:
    If Not blnInError Then
        Exit Function
    Else
        On Error GoTo 0

        Err.Raise lngErrorNumber, , strErrorDescription
    End If

Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
    End If

    Resume CleanUp
    
End Function

Private Function GetFileType(ByVal bytType As Byte) As String
    Select Case bytType
        Case 1
            GetFileType = "Database"
        Case 2
            GetFileType = "Transaction Log"
        Case 4
            GetFileType = "File"
        Case 5
            GetFileType = "Differential Database"
        Case Else
            GetFileType = "Unknown"
    End Select
End Function


Public Property Let BackupFileLocation(ByVal strFileLocation As String)
    mstrBackupFileLocation = strFileLocation
End Property

Public Property Let BackupFileName(ByVal strFileName As String)
    mstrBackupFileName = strFileName
End Property

Public Property Get BackupFileName() As String
    BackupFileName = mstrBackupFileName
End Property


Public Function GetListOfBackupSetsInFile(ByVal strFile As String) As Variant
On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function GetListOfBackupSetsInFile"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    Dim objQryRslts As SQLDMO.QueryResults
    Dim i As Integer
    Dim strListOfFiles As String
    
    Set mobjRestore = New Restore
    
    mobjRestore.Files = strFile
    mobjRestore.Action = SQLDMORestore_Database
    
    Set objQryRslts = mobjRestore.ReadBackupHeader(mobjSQLServer)
    
    For i = 1 To objQryRslts.Rows
        strListOfFiles = strListOfFiles & "," & objQryRslts.GetColumnString(i, 6) & ": Name: [" & objQryRslts.GetColumnString(i, 1) & "] DB: [" & objQryRslts.GetColumnString(i, 10) & "] Type: [" & GetFileType(objQryRslts.GetColumnString(i, 3)) & "]"
    Next i
    
    GetListOfBackupSetsInFile = Split(strListOfFiles, ",")
    
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next
        Set objQryRslts = Nothing
        Set mobjRestore = Nothing
Out:
    If Not blnInError Then
        Exit Function
    Else
        On Error GoTo 0
        
        Err.Raise lngErrorNumber, , strErrorDescription
    End If

Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
    End If

    Resume CleanUp
End Function

Public Function GetDatabaseSummary(strDbOwner, strCurrentSize, strDataSpaceSize, strIndexSpaceSize, strPrimaryFilePath, dtCreateDate, strMemoryUsage, arrDBFiles)
'On Error Resume Next

    Dim objDB As SQLDMO.Database
    Dim strDBFiles As String
    Dim objQryRslts As SQLDMO.QueryResults
    Dim i As Integer
    
    Set objDB = mobjSQLServer.Databases(mstrDatabaseName)
    
    strCurrentSize = objDB.Size
    strPrimaryFilePath = objDB.PrimaryFilePath
    dtCreateDate = objDB.CreateDate
    strDataSpaceSize = objDB.DataSpaceUsage
    strMemoryUsage = objDB.GetMemoryUsage()
    strIndexSpaceSize = objDB.IndexSpaceUsage
    strDbOwner = objDB.Owner
    
    Set objQryRslts = objDB.EnumFiles()
    
    For i = 1 To objQryRslts.Rows
        strDBFiles = strDBFiles & "," & Trim$(objQryRslts.GetColumnString(i, 8))
    Next i
    
    'objDB.GetMemoryUsage
    arrDBFiles = Split(strDBFiles, ",")
    
    Set objQryRslts = Nothing
    Set objDB = Nothing
End Function

Public Function RestoreDatabase(ByVal blnOverwrite As Boolean, ByVal intFileNumber As Integer)
On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function RestoreDatabase"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    Set mobjRestore = New SQLDMO.Restore
    With mobjRestore
        .Database = mstrDatabaseName
        .Files = mstrBackupFileLocation & mstrBackupFileName
        .FileNumber = intFileNumber
        .ReplaceDatabase = blnOverwrite
        .SQLRestore mobjSQLServer
    End With
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next

Out:
    If Not blnInError Then
        Exit Function
    Else
        On Error GoTo 0
        
        Err.Raise lngErrorNumber, , strErrorDescription
    End If

Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
    End If

    Resume CleanUp
    
End Function

Public Function VerifyBackupFile(ByVal strFile As String, ByVal strDatabaseName As String)
    On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function VerifyBackupFile"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    Set mobjRestore = New SQLDMO.Restore
    With mobjRestore
        .Database = strDatabaseName
        .Files = strFile
        .SQLVerify mobjSQLServer
    End With
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next

Out:
    If Not blnInError Then
        Exit Function
    Else
        On Error GoTo 0
        
        Err.Raise lngErrorNumber, , strErrorDescription
    End If

Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
    End If

    Resume CleanUp
    
End Function

Private Function GenerateBackupScript() As String
    On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function GenerateBackupScript"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    Dim strScript As String
    Dim strLocation As String
    
    strLocation = Trim$(mstrBackupFileLocation & mstrBackupFileName)
    
    Set mobjBackup = New SQLDMO.Backup
    
    With mobjBackup
        .BackupSetName = mstrBackupSetName
        .Database = mstrDatabaseName
        .Action = SQLDMOBackup_Database
        .BackupSetDescription = "Backup of " & mstrDatabaseName
        .MediaName = mstrDatabaseName & " MASTER"
        .Files = strLocation
        .Initialize = True
        strScript = .GenerateSQL
        'strScript = "DBCC CHECKDB ('" & strDatabaseName & "', REPAIR_FAST')"
    End With
    GenerateBackupScript = strScript
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next
        Set mobjBackup = Nothing
Out:
    If Not blnInError Then
        Exit Function
    Else
        On Error GoTo 0
        
        Err.Raise lngErrorNumber, , strErrorDescription
    End If

Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
    End If

    Resume CleanUp

End Function

Public Function BackupDatabase()
  On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function BackupDatabase"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    Dim strLocation As String
    Dim i
    Dim objFSO As Scripting.FileSystemObject
    Dim blnInitialiseMedia As Boolean
    
    Set objFSO = New FileSystemObject
    
    strLocation = Trim$(mstrBackupFileLocation & mstrBackupFileName)
    
    '''''''''''''''''''''''''''''''''''''''''''
    '
    '   Remove this section if you would like to attempt to set a
    '   backup file location that contains spaces.
    '
    ''''''''''''''''''''''''''''''''''''''''''
    If InStr(1, strLocation, " ") > 0 Then
        'I'm at my end here
        'I've tried all ways and means to set a backup location that has
        'a space in it but sql refuses, therefore I've added this simple check
        'to refuse spaces
        Err.Raise G_ERR_APP_DEFINED_ERROR, , "SQL Server does not allow backup folder locations that contain spaces."
    End If
    '''''''''''''''''''''''''''''''''''''''''''
    '
    '   End of Removable section
    '
    '''''''''''''''''''''''''''''''''''''''''''
    If objFSO.FileExists(strLocation) = True Then
        'Always reinit the media in the bak file. but only
        'if the file exists before
        blnInitialiseMedia = True
    Else
        blnInitialiseMedia = False
    End If
    
    Set mobjBackup = New SQLDMO.Backup
    With mobjBackup
        .Devices = ""
        .BackupSetName = mstrBackupSetName
        .Database = mstrDatabaseName
        .Action = SQLDMOBackup_Database
        .BackupSetDescription = "Backup of " & mstrDatabaseName
        .MediaName = mstrDatabaseName & " MASTER"
        .Files = strLocation
        .Initialize = blnInitialiseMedia
        .SQLBackup mobjSQLServer
    End With
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next
        Set objFSO = Nothing
        Set mobjBackup = Nothing
Out:
    If Not blnInError Then
        Exit Function
    Else
        On Error GoTo 0
        
        Err.Raise lngErrorNumber, , strErrorDescription
    End If

Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
    End If

    Resume CleanUp
End Function

Public Function GetListOfDatabases(ByVal blnIncludeSystem As Boolean) As Variant
   On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function GetListOfDatabases()"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
        Dim strDbList As String
        Dim objDB As Object
        Dim vntRetVal As Variant
        
        If mblnConnected = False Then
            Err.Raise G_ERR_APP_DEFINED_ERROR, , "Not Connected to a Valid SQL Server."
        End If
        
        If blnIncludeSystem = True Then
            For Each objDB In mobjSQLServer.Databases
                strDbList = objDB.Name & "," & strDbList
            Next
        Else
            For Each objDB In mobjSQLServer.Databases
                If objDB.SystemObject = False Then
                    strDbList = objDB.Name & IIf((strDbList <> vbNullString), "," & strDbList, "")
                End If
            Next
        End If
        
        GetListOfDatabases = Split(strDbList, ",")
        
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next
        Set objDB = Nothing
Out:
    If Not blnInError Then
        Exit Function
    Else
        On Error GoTo 0
        
        Err.Raise lngErrorNumber, , strErrorDescription
    End If

Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
    End If

    Resume CleanUp
End Function

Public Function DisConnectFromSQLServer()
   On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function DisConnectFromSQLServer()"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    If Not mobjSQLServer Is Nothing Then
        mobjSQLServer.DisConnect
    End If
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next

Out:
    If Not blnInError Then
        Exit Function
    Else
        On Error GoTo 0
        
        Err.Raise lngErrorNumber, , strErrorDescription
    End If

Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
    End If

    Resume CleanUp
    
End Function

Public Function ConnectToSQLServer(ByVal strServerName As String, ByVal strUserName As String, ByVal strpassword As String, Optional ByVal blnWindowsAuth As Boolean = False) As Boolean
    On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function ConnectToSQLServer"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    On Error Resume Next
    
    If blnWindowsAuth = False Then
        mobjSQLServer.Connect strServerName, strUserName, strpassword
    Else
        mobjSQLServer.LoginSecure = True
        mobjSQLServer.Connect strServerName
    End If
    
    If Err.Number <> 0 Then
        ConnectToSQLServer = False
        mblnConnected = False
        strErrorDescription = Err.Description
        On Error GoTo Handle_Error
        Err.Raise G_ERR_APP_DEFINED_ERROR, , strErrorDescription
    Else
        On Error GoTo Handle_Error
        mobjSQLServer.ApplicationName = "Kamau Database Manager " & App.Major & "." & App.Minor
        ConnectToSQLServer = True
        mblnConnected = True
    End If
    
     On Error GoTo Handle_Error
    
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next
    
Out:
    If Not blnInError Then
        Exit Function
    Else
        On Error GoTo 0
        
        Err.Raise lngErrorNumber, , strErrorDescription
    End If

Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
    End If

    Resume CleanUp
    
End Function

Private Sub Class_Initialize()
   On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Sub Class_Initialize"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    Set mobjSQLServer = New SQLDMO.SQLServer
    
    mblnConnected = False
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next

Out:
    If Not blnInError Then
        Exit Sub
    Else
        On Error GoTo 0
        
        Err.Raise lngErrorNumber, , strErrorDescription
    End If

Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
    End If

    Resume CleanUp
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    mobjSQLServer.DisConnect
    Set mobjBackup = Nothing
    Set mobjRestore = Nothing
    Set mobjSQLServer = Nothing
End Sub
Private Function FormatDateToSQLDate(ByVal strDate) As String
   On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function FormatDateToSQLDate"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    Dim strSQLDate As String
    Dim strYear As String
    Dim strDay As String
    Dim strMonth As String
    
    strYear = DatePart("yyyy", strDate)
    strDay = DatePart("d", strDate)
    strMonth = DatePart("m", strDate)
    
    If Len(strDay) < 2 Then
        strDay = "0" & strDay
    End If
    
    If Len(strMonth) < 2 Then
        strMonth = "0" & strMonth
    End If
    
    strSQLDate = strYear & strMonth & strDay
    
    FormatDateToSQLDate = strSQLDate
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next

Out:
    If Not blnInError Then
        Exit Function
    Else
        On Error GoTo 0
        
        Err.Raise lngErrorNumber, , strErrorDescription
    End If

Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
    End If

    Resume CleanUp
    
End Function

Public Function FormatSQLDayOfWeekToNormalDay(ByVal intDay) As String
    On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function FormatSQLDayOfWeekToNormalDay"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    Select Case intDay
        Case SQLDMOMonth_Friday
            FormatSQLDayOfWeekToNormalDay = "Friday"
            
        Case SQLDMOMonth_Monday
            FormatSQLDayOfWeekToNormalDay = "Monday"
            
        Case SQLDMOMonth_Tuesday
            FormatSQLDayOfWeekToNormalDay = "Tuesday"
            
        Case SQLDMOMonth_Wednesday
            FormatSQLDayOfWeekToNormalDay = "Wednesday"
            
        Case SQLDMOMonth_Thursday
            FormatSQLDayOfWeekToNormalDay = "Thursday"
            
        Case SQLDMOMonth_Saturday
            FormatSQLDayOfWeekToNormalDay = "Saturday"
            
        Case SQLDMOMonth_Sunday
            FormatSQLDayOfWeekToNormalDay = "Sunday"
    End Select
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next

Out:
    If Not blnInError Then
        Exit Function
    Else
        On Error GoTo 0
        
        Err.Raise lngErrorNumber, , strErrorDescription
    End If

Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
    End If

    Resume CleanUp
    
End Function

Private Function FormatDayOfWeek(ByVal strDay) As Long
   On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function FormatDayOfWeek"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    Select Case UCase$(strDay)
        Case "MONDAY"
            FormatDayOfWeek = SQLDMOMonth_Monday
            
        Case "TUESDAY"
            FormatDayOfWeek = SQLDMOMonth_Tuesday
            
        Case "WEDNESDAY"
            FormatDayOfWeek = SQLDMOMonth_Wednesday
        
        Case "THURSDAY"
            FormatDayOfWeek = SQLDMOMonth_Thursday
            
        Case "FRIDAY"
            FormatDayOfWeek = SQLDMOMonth_Friday
            
        Case "SATURDAY"
            FormatDayOfWeek = SQLDMOMonth_Saturday
            
        Case "SUNDAY"
            FormatDayOfWeek = SQLDMOMonth_Sunday
            
        Case Else
            FormatDayOfWeek = SQLDMOMonth_Unknown
    End Select
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next

Out:
    If Not blnInError Then
        Exit Function
    Else
        On Error GoTo 0
        
        Err.Raise lngErrorNumber, , strErrorDescription
    End If

Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
    End If

    Resume CleanUp
    
End Function

Public Function FormatSQLTimeToNormalTime(ByVal strTime) As String
On Error Resume Next

    Dim hr As String, min As String, sec As String, strNewTime
    hr = Left(strTime, Len(strTime) - 4)
    min = Left(Right(strTime, 4), 2)
    sec = Right(strTime, 2)
        
    If Len(hr) < 2 Then
        hr = "0" & hr
    End If
    
    If Len(min) < 2 Then
        min = "0" & min
    End If
    
    If Len(sec) < 2 Then
        sec = "0" & sec
    End If
    
    strNewTime = hr & ":" & min & ":" & sec
    
    FormatSQLTimeToNormalTime = strNewTime
    
End Function


Private Function FormatTimeToSQLTime(ByVal strTime) As String
   On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function FormatTimeToSQLTime"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
        Dim hr As String, min As String, sec As String
        
        hr = Hour(strTime)
        min = Minute(strTime)
        sec = Second(strTime)
        
        strTime = CStr(strTime)
        
        If Len(min) < 2 Then
            min = "0" & min
        End If
        
        If Len(hr) < 2 Then
            hr = "0" & hr
        End If
        
        If Len(sec) < 2 Then
            sec = "0" & sec
        End If
        
        FormatTimeToSQLTime = hr & min & sec
        
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next

Out:
    If Not blnInError Then
        Exit Function
    Else
        On Error GoTo 0
        
        Err.Raise lngErrorNumber, , strErrorDescription
    End If

Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
    End If

    Resume CleanUp
    
End Function

Public Function SetSchedule(ByRef Frequency As enFrequencyTypes, ByRef dtStartTime, ByRef strDayOfWeek, ByRef startDate)
   On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function SetSchedule"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    '============================ MAIN BODY ===========================
        Dim strJobName As String
        Dim objJob As SQLDMO.Job
        Dim objJobSchedule As SQLDMO.JobSchedule
        Dim objJobStep As SQLDMO.JobStep
        Dim blnJobExistedBefore As Boolean
        
        strJobName = UCase$(G_JOBNAME & mstrDatabaseName)
        
        On Error Resume Next
        
        'Check if Job Existed Before
        Set objJob = mobjSQLServer.JobServer.Jobs(strJobName)
        
        If Err.Number <> 0 Then
            Set objJob = New SQLDMO.Job
            Set objJobSchedule = New JobSchedule
            Set objJobStep = New JobStep
            blnJobExistedBefore = False
        Else
            Set objJobSchedule = objJob.JobSchedules(G_SCHEDULE_NAME)
            Set objJobStep = objJob.JobSteps(G_BACKUP_STEP_NAME)
            blnJobExistedBefore = True
            objJob.BeginAlter
        End If
        
        On Error GoTo Handle_Error
        
        objJob.Name = strJobName
        objJob.Description = G_JOBDESCRIPTION
        objJob.Category = G_JOBCATEGORY
        
        If blnJobExistedBefore = False Then
            mobjSQLServer.JobServer.Jobs.Add objJob
        Else
            objJobStep.BeginAlter
        End If
        ''''''''''''''
        'Set the one and only Job Step
        ''''''''''''''
        With objJobStep
            .Name = G_BACKUP_STEP_NAME 'SEt the name of the job
            .StepID = 1
            .SubSystem = "TSQL"
            .DatabaseName = mstrDatabaseName
            .Command = GenerateBackupScript()
            .OnFailAction = SQLDMOJobStepAction_QuitWithFailure 'Quit if execution of the job failes
            .OSRunPriority = SQLDMORunPri_Normal
        End With
        
        If blnJobExistedBefore = False Then
            objJob.AddStepToJob objJobStep
            objJob.StartStepID = 1
            objJob.ApplyToTargetServer mobjSQLServer.Name
            objJob.BeginAlter
        Else
            objJobStep.DoAlter
            objJobSchedule.BeginAlter
        End If
        
        
        '''''''''''''''''''''''''''''''''''''''
        'Set the One and Only Schedule
        '''''''''''''''''''''''''''''''''''''''
        
        With objJobSchedule
            .Name = G_SCHEDULE_NAME
            .Schedule.FrequencyType = Frequency
            If Frequency = BackupDaily Then
                'The Schedule shall recur once a day
                .Schedule.FrequencyInterval = 1
            Else
                If Frequency = BackupWeekly Then
                    'The Schedule shall recur once a week
                    .Schedule.FrequencyRecurrenceFactor = 1
                End If
                'Get the day of the week the job should execute
                .Schedule.FrequencyRecurrenceFactor = 1
                .Schedule.FrequencyInterval = FormatDayOfWeek(strDayOfWeek)
            End If
            .Schedule.ActiveStartDate = FormatDateToSQLDate(startDate) 'Date the job starts working
            .Schedule.ActiveStartTimeOfDay = FormatTimeToSQLTime(dtStartTime) 'Time of day the job shall begin
            .Schedule.ActiveEndDate = SQLDMO_NOENDDATE 'This job shall never end
            .Schedule.ActiveEndTimeOfDay = SQLDMO_NOENDTIME
        End With
        
        objJob.BeginAlter
        
        If blnJobExistedBefore = False Then
            objJob.JobSchedules.Add objJobSchedule
        Else
            objJobSchedule.DoAlter
        End If
        
        objJob.DoAlter
        
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next
        Set objJob = Nothing
        Set objJobSchedule = Nothing
        Set objJobStep = Nothing
Out:
    If Not blnInError Then
        Exit Function
    Else
        On Error GoTo 0
        
        Err.Raise lngErrorNumber, , strErrorDescription
    End If

Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
    End If

    Resume CleanUp
    
End Function


Public Function DropSchedule() As Boolean
   On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function DropSchedule()"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    Dim objJob As SQLDMO.Job
    On Error Resume Next
    
    Set objJob = mobjSQLServer.JobServer.Jobs(UCase$(G_JOBNAME & mstrDatabaseName))
    
    If Err.Number <> 0 Then
        DropSchedule = False
        GoTo CleanUp
    End If
    
    On Error GoTo Handle_Error
    
    'Attempt removal
    mobjSQLServer.JobServer.RemoveJobByID objJob.JobID
    DropSchedule = True
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next
        Set objJob = Nothing
Out:
    If Not blnInError Then
        Exit Function
    Else
        On Error GoTo 0
        
        Err.Raise lngErrorNumber, , strErrorDescription
    End If

Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
    End If

    Resume CleanUp
    
End Function

Public Function GetSchedule(ByRef Frequency As enFrequencyTypes, ByRef dtStartTime, ByRef strDayOfWeek, ByRef startDate)
    On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function GetSchedule()"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    Dim objJob As SQLDMO.Job
    Dim objJobSchedule As SQLDMO.JobSchedule
    Dim objJobStep As SQLDMO.JobStep
    
    On Error Resume Next
    
    Set objJob = mobjSQLServer.JobServer.Jobs(UCase$(G_JOBNAME & mstrDatabaseName))
    
    If Err.Number <> 0 Then
        'Exit if job is not found
        Frequency = BackupMonthly
        dtStartTime = vbNullString
        strDayOfWeek = "Monday"
        startDate = Now()
        GoTo CleanUp
    End If

    On Error GoTo Handle_Error
    
    Set objJobSchedule = objJob.JobSchedules(G_SCHEDULE_NAME)
    Set objJobStep = objJob.JobSteps.Item(G_BACKUP_STEP_NAME)
    
    Frequency = objJobSchedule.Schedule.FrequencyType
    dtStartTime = FormatSQLTimeToNormalTime(objJobSchedule.Schedule.ActiveStartTimeOfDay)
    strDayOfWeek = FormatSQLDayOfWeekToNormalDay(objJobSchedule.Schedule.FrequencyInterval)
    startDate = objJobSchedule.Schedule.ActiveStartDate
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next
        Set objJob = Nothing
        Set objJobSchedule = Nothing
        Set objJobStep = Nothing
Out:
    If Not blnInError Then
        Exit Function
    Else
        On Error GoTo 0
        
        Err.Raise lngErrorNumber, , strErrorDescription
    End If

Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
    End If

    Resume CleanUp
    
End Function

'Private Sub mobjBackup_PercentComplete(ByVal Message As String, ByVal Percent As Long)
'    RaiseEvent PercentageCompleted(Message, Percent)
'End Sub
'Private Sub mobjRestore_PercentComplete(ByVal Message As String, ByVal Percent As Long)
'    RaiseEvent RestorePercentageCompleted(Message, Percent)
'End Sub
