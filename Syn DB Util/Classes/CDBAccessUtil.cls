VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDBAccessUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const M_MODULENAME As String = "CDBAccessUtil"
'''''''''''''''''''''''
Private mstrBackupFileName As String
Private mstrBackupFileFolderLocation As String
Private mstrDatabaseLocation As String
'''''''''''''''''''''''''
Private Const M_BACKUPFILE_EXTENSION As String = "BAK"
'''''''''''''''''''''''''
Public Property Let BackupFileName(ByVal strBackupFileName As String)
    mstrBackupFileName = strBackupFileName
End Property

Public Property Let BackupFolderLocation(ByVal strBackupFileFolderLocation As String)
    mstrBackupFileFolderLocation = strBackupFileFolderLocation
End Property

Public Property Let DatabaseLocation(ByVal strDatabaseLocation As String)
    mstrDatabaseLocation = strDatabaseLocation
End Property

Public Function CompactAndRepair()
    On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function CompactAndRepair()"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    Dim strDBName As String
    Dim strNewCompactedDatabaseName As String
    Dim strDBFolder As String
    Dim objFso As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    
    Set objFso = New FileSystemObject
    
    strDBName = objFso.GetFileName(mstrDatabaseLocation)
    strNewCompactedDatabaseName = "CMPT-" & strDBName
    strDBFolder = Left$(mstrDatabaseLocation, InStrRev(mstrDatabaseLocation, "\"))
    
    If objFso.FileExists(strDBFolder & strNewCompactedDatabaseName) Then
        Kill strDBFolder & strNewCompactedDatabaseName
    End If
    
    DBEngine.CompactDatabase mstrDatabaseLocation, strDBFolder & strNewCompactedDatabaseName
    
    Set objFile = objFso.GetFile(strDBFolder & strNewCompactedDatabaseName)
    
    Kill mstrDatabaseLocation
    
    objFile.Name = strDBName
    
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next
    Set objFso = Nothing
    Set objFile = Nothing
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

Public Function RestoreDatabase()
    On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Function RestoreDatabase()"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    Dim objFso As Scripting.FileSystemObject
    Dim strFileLocation As String
    Dim strMasterDatabaseName As String
    Dim strMasterDatabaseFolder As String
    Dim objFile As Scripting.File
    
    strFileLocation = mstrBackupFileFolderLocation & mstrBackupFileName
    
    Set objFso = New Scripting.FileSystemObject
    
    If objFso.FileExists(mstrDatabaseLocation) = False Then
        Err.Raise G_ERR_APP_DEFINED_ERROR, , "Master Database File Cannot be found!"
    End If
    
    If objFso.FolderExists(mstrBackupFileFolderLocation) = False Then
        Err.Raise G_ERR_APP_DEFINED_ERROR, , "Backup Folder Cannot Be found!"
    End If
    
    strMasterDatabaseName = objFso.GetFileName(mstrDatabaseLocation)
    strMasterDatabaseFolder = Left$(mstrDatabaseLocation, InStrRev(mstrDatabaseLocation, "\"))
    
    If Right$(strMasterDatabaseFolder, 1) <> "\" Then
        strMasterDatabaseFolder = strMasterDatabaseFolder & "\"
    End If
    
    objFso.CopyFile strFileLocation, mstrDatabaseLocation, True
        
    Set objFile = objFso.GetFile(strMasterDatabaseFolder & mstrBackupFileName)
    objFile.Name = strMasterDatabaseName
    
    'Do a backup
    mstrBackupFileName = objFso.GetBaseName(mstrDatabaseLocation) & ".bak"
    mstrBackupFileFolderLocation = strMasterDatabaseFolder
    
    BackupDatabase
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next
    Set objFso = Nothing
    Set objFile = Nothing
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
    
    Const PROC_NAME As String = "Function BackupDatabase()"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    Dim objFso As Scripting.FileSystemObject
    Dim strBackupFileLocation As String
    Dim objFile As Scripting.File
    
    strBackupFileLocation = mstrBackupFileFolderLocation & mstrBackupFileName
    
    Set objFso = New Scripting.FileSystemObject
    
    If objFso.FileExists(mstrDatabaseLocation) = False Then
        Err.Raise G_ERR_APP_DEFINED_ERROR, , "Master Database Location must be specified!"
    End If
    
    If objFso.FolderExists(mstrBackupFileFolderLocation) = False Then
        Err.Raise G_ERR_APP_DEFINED_ERROR, , "Specified Backup Folder Location does not exist!"
    End If
    
    If objFso.FileExists(strBackupFileLocation) Then
        'Check if readonly
        Set objFile = objFso.GetFile(strBackupFileLocation)
        objFile.Attributes = Normal
    End If
    
    objFso.CopyFile mstrDatabaseLocation, mstrBackupFileFolderLocation & mstrBackupFileName, True
    
    'Change Extension to .BAK
    If UCase$(M_BACKUPFILE_EXTENSION) <> UCase$(objFso.GetExtensionName(strBackupFileLocation)) Then
        Set objFile = objFso.GetFile(strBackupFileLocation)
        objFile.Name = Left$(objFile.Name, InStrRev(objFile.Name, ".")) & "bak"
    End If
    
    ' ============================ MAIN BODY ===========================
   
CleanUp:
    On Error Resume Next
        Set objFso = Nothing
        Set objFile = Nothing
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
