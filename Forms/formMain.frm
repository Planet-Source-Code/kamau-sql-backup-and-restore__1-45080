VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form formMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup and Restore"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "formMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   75
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   13
      Top             =   7080
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14711
            Text            =   "Backup and Restore Utility"
            TextSave        =   "Backup and Restore Utility"
            Key             =   "pnlMain"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      FillColor       =   &H00E0E0E0&
      Height          =   870
      Left            =   -30
      ScaleHeight     =   810
      ScaleWidth      =   8370
      TabIndex        =   2
      Top             =   -30
      Width           =   8430
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "formMain.frx":0442
         Top             =   150
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"formMain.frx":0884
         Height          =   450
         Left            =   780
         TabIndex        =   4
         Top             =   330
         Width           =   7545
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Database Backup and Restore"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   810
         TabIndex        =   3
         Top             =   60
         Width           =   3945
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5715
      Left            =   15
      TabIndex        =   1
      Top             =   885
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   10081
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      BackColor       =   -2147483648
      TabCaption(0)   =   "Microsoft® S&QL Server™"
      TabPicture(0)   =   "formMain.frx":0922
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cboDatabase"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdConnect"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdViewDatabaseInfo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraSQLRestore"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraSQLBackup"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "optConnType(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "optConnType(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtServerName"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.Frame Frame2 
         Caption         =   "Operation"
         ForeColor       =   &H8000000D&
         Height          =   690
         Left            =   4125
         TabIndex        =   41
         Top             =   1605
         Width           =   4110
         Begin VB.OptionButton optOperation 
            Caption         =   "&Backup"
            Height          =   195
            Index           =   0
            Left            =   450
            TabIndex        =   43
            Top             =   300
            Width           =   990
         End
         Begin VB.OptionButton optOperation 
            Caption         =   "&Restore"
            Height          =   195
            Index           =   1
            Left            =   1890
            TabIndex        =   42
            Top             =   300
            Width           =   1605
         End
      End
      Begin VB.TextBox txtServerName 
         Height          =   300
         Left            =   795
         TabIndex        =   39
         Top             =   540
         Width           =   2700
      End
      Begin VB.OptionButton optConnType 
         Caption         =   "&SQL Server Authentication"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   33
         Top             =   1260
         Value           =   -1  'True
         Width           =   2235
      End
      Begin VB.Frame Frame1 
         Height          =   1050
         Left            =   90
         TabIndex        =   34
         Top             =   1245
         Width           =   3945
         Begin VB.TextBox txtUserName 
            Height          =   285
            Left            =   1095
            TabIndex        =   36
            Top             =   270
            Width           =   2715
         End
         Begin VB.TextBox txtPassword 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1095
            PasswordChar    =   "*"
            TabIndex        =   35
            Top             =   585
            Width           =   1530
         End
         Begin VB.Label Label5 
            Caption         =   "Login:"
            Height          =   255
            Left            =   135
            TabIndex        =   38
            Top             =   345
            Width           =   465
         End
         Begin VB.Label Label6 
            Caption         =   "Password:"
            Height          =   225
            Left            =   120
            TabIndex        =   37
            Top             =   645
            Width           =   795
         End
      End
      Begin VB.OptionButton optConnType 
         Caption         =   "&Windows Authentication"
         Height          =   210
         Index           =   0
         Left            =   210
         TabIndex        =   32
         Top             =   930
         Width           =   2070
      End
      Begin VB.Frame fraSQLBackup 
         Caption         =   "Complete Backup of Database"
         ForeColor       =   &H00008000&
         Height          =   2895
         Left            =   60
         TabIndex        =   5
         Top             =   2730
         Visible         =   0   'False
         Width           =   8175
         Begin VB.CommandButton cmdAutoFill 
            Height          =   315
            Left            =   7365
            Picture         =   "formMain.frx":093E
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Automatically autofill the backup location based off the connection information"
            Top             =   630
            Width           =   375
         End
         Begin VB.CommandButton cmdScheduleBackup 
            Caption         =   "&Schedule"
            Enabled         =   0   'False
            Height          =   345
            Left            =   1620
            TabIndex        =   19
            Top             =   2475
            Width           =   1425
         End
         Begin VB.CommandButton cmdVerify 
            Caption         =   "&Verify"
            Enabled         =   0   'False
            Height          =   345
            Left            =   3060
            TabIndex        =   18
            Top             =   2475
            Width           =   1245
         End
         Begin VB.TextBox txtBackupSetName 
            Height          =   285
            Left            =   1515
            TabIndex        =   17
            Top             =   345
            Width           =   2940
         End
         Begin VB.TextBox txtSQLBackupFileName 
            Height          =   285
            Left            =   1515
            TabIndex        =   15
            Top             =   975
            Width           =   6195
         End
         Begin VB.CommandButton cmdBackup 
            Caption         =   "&Backup Now!"
            Enabled         =   0   'False
            Height          =   345
            Left            =   60
            TabIndex        =   12
            Top             =   2475
            Width           =   1545
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   330
            Left            =   7755
            TabIndex        =   7
            Top             =   615
            Width           =   330
         End
         Begin VB.TextBox txtBackupFolder 
            Height          =   285
            Left            =   1515
            TabIndex        =   6
            Top             =   660
            Width           =   5820
         End
         Begin VB.Label Label11 
            Caption         =   $"formMain.frx":0A88
            Height          =   690
            Left            =   120
            TabIndex        =   30
            Top             =   1575
            Width           =   7710
         End
         Begin VB.Label lblBackupName 
            Caption         =   "&Backup Name"
            Height          =   210
            Left            =   120
            TabIndex        =   16
            Top             =   375
            Width           =   1185
         End
         Begin VB.Label Label10 
            Caption         =   "&Backup File Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1035
            Width           =   1515
         End
         Begin VB.Label Label3 
            Caption         =   "&Location:"
            Height          =   210
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   1110
         End
      End
      Begin VB.Frame fraSQLRestore 
         Caption         =   "Restoration of Database"
         ForeColor       =   &H000000C0&
         Height          =   2925
         Left            =   60
         TabIndex        =   20
         Top             =   2685
         Visible         =   0   'False
         Width           =   8175
         Begin VB.CheckBox chkReplace 
            Caption         =   "Completely &Replace Database"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1515
            TabIndex        =   28
            Top             =   2640
            Value           =   1  'Checked
            Width           =   3945
         End
         Begin VB.ComboBox cboData 
            Height          =   315
            Left            =   1020
            TabIndex        =   26
            Top             =   600
            Width           =   7080
         End
         Begin VB.CommandButton cmdViewSets 
            Caption         =   "Refresh Data List"
            Height          =   300
            Left            =   1020
            TabIndex        =   25
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton cmdRestore 
            Caption         =   "&Restore Now!"
            Enabled         =   0   'False
            Height          =   330
            Left            =   75
            TabIndex        =   24
            Top             =   2535
            Width           =   1380
         End
         Begin VB.CommandButton cmdRestoreBrowse 
            Caption         =   "..."
            Height          =   270
            Left            =   7740
            TabIndex        =   23
            Top             =   285
            Width           =   345
         End
         Begin VB.TextBox txtSQLRestorationFileLocation 
            Height          =   285
            Left            =   1020
            TabIndex        =   22
            Top             =   285
            Width           =   6675
         End
         Begin VB.Label Label9 
            Caption         =   "Media Data:"
            Height          =   255
            Left            =   105
            TabIndex        =   27
            Top             =   675
            Width           =   975
         End
         Begin VB.Label lblBackupFile 
            Caption         =   "&Backup File:"
            Height          =   225
            Left            =   105
            TabIndex        =   21
            Top             =   330
            Width           =   1005
         End
      End
      Begin VB.CommandButton cmdViewDatabaseInfo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7785
         Picture         =   "formMain.frx":0B74
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "View More Information about this database."
         Top             =   1005
         Width           =   330
      End
      Begin VB.CommandButton cmdConnect 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Co&nnect"
         Height          =   315
         Left            =   3525
         TabIndex        =   11
         Top             =   540
         Width           =   1155
      End
      Begin VB.ComboBox cboDatabase 
         Height          =   315
         Left            =   5040
         TabIndex        =   10
         Top             =   1005
         Width           =   2700
      End
      Begin VB.Label Label4 
         Caption         =   "Server:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   585
         Width           =   675
      End
      Begin VB.Label Label7 
         Caption         =   "Database:"
         Height          =   240
         Left            =   4185
         TabIndex        =   9
         Top             =   1215
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   345
      Left            =   7140
      TabIndex        =   0
      Top             =   6645
      Width           =   1215
   End
End
Attribute VB_Name = "formMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''
'Project: Backup and Restore
'Author: John Kamau
'Email: john_ndungu@hotmail.com
'Comments:
'   For comments on this prj. See the CDBSQLUtil class in the SynDBUtil dll
'
'''''''''''''''''''''''''''''''''
Public mobjCDBSQL As SynDBUtil.CDBSQLUtil
Private mobjCDBAccess As SynDBUtil.CDBAccessUtil

Private Const M_MODULENAME As String = "formMain"
Private mblnConnected As Boolean

Private Sub cboDatabase_Change()
    SetDefaultBackupValues
    fraSQLBackup.Caption = "&Backup of Database (" & cboDatabase.Text & ")"
    fraSQLRestore.Caption = "&Complete Restoration of Database(" & cboDatabase.Text & ")"
End Sub

Private Sub cboDatabase_Click()
    SetDefaultBackupValues
    fraSQLBackup.Caption = "&Backup of Database (" & cboDatabase.Text & ")"
    fraSQLRestore.Caption = "&Complete Restoration of Database(" & cboDatabase.Text & ")"
End Sub

Private Sub cmdAutoFill_Click()
    On Error Resume Next
    txtBackupFolder.Text = mobjCDBSQL.GetDefaultBackupLocation
End Sub

Private Sub cmdBackup_Click()
    On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Sub cmdBackup_Click()"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    '================================ MAIN BODY =============================================
    Dim strFileName As String
    Dim strFolderLocation As String
    
    If cboDatabase = vbNullString Then
        MsgBox "You Must Connect to SQL Server and Select a database.", vbInformation, G_APPTITLE
        Exit Sub
    End If
    
    If txtBackupFolder = vbNullString Or txtBackupSetName = vbNullString Or txtSQLBackupFileName = vbNullString Then
        MsgBox "You Must select a Backup folder location, Enter a Backup name and a Backup FILE name. " & vbCrLf & _
                "Note: Remember that the location is in reference to the location of your " & _
                "SQL Server, ensure that the folder(s) exist, otherwise you shall recieve an error.", vbInformation, G_APPTITLE
        Exit Sub
    End If
    
    strFolderLocation = txtBackupFolder
    strFileName = txtSQLBackupFileName
    
    'strFileName = Right(strFolderLocation, (Len(strFolderLocation) - InStrRev(txtBackupFolder.Text, "\")))
    With mobjCDBSQL
        .BackupFileLocation = strFolderLocation
        .BackupFileName = strFileName
        .DatabaseName = cboDatabase.Text
        .BackupSetName = txtBackupSetName
        StatusMessage "Backing Up " & cboDatabase.Text & " Completely to file " & strFileName & "...."
        mobjCDBSQL.BackupDatabase
        StatusMessage "Successfully Backed Up " & cboDatabase.Text & "."
    End With
    '================================ MAIN BODY =============================================
    
    
CleanUp:
    On Error Resume Next
    
Out:
    
    Exit Sub
    
Handle_Error:
    blnInError = True
    StatusMessage "Backup Incomplete due to Errors!"
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
        ReportUnexpectedError lngErrorNumber, strErrorDescription
        ShowErrorMessage "Error Number: " & lngErrorNumber & vbCrLf & "Error: " & strErrorDescription
    Else
        ShowErrorMessage strErrorDescription
    End If
    
    Resume CleanUp
End Sub

Private Sub SetDefaultBackupValues()
    txtSQLBackupFileName = "BK-" & cboDatabase.Text & "-" & ReturnMonthString(Month(Date)) & ".BAK"
    txtBackupSetName = cboDatabase.Text & " Backup"
End Sub

Private Sub cmdBrowse_Click()
    If cboDatabase.Text = vbNullString Then
        MsgBox "Please Select a Database!", vbInformation, G_APPTITLE
        cboDatabase.SetFocus
    Else
        txtBackupFolder.Text = SynBrowseForFolder(Me.hwnd, "Select your backup folder location") & "\"
        SetDefaultBackupValues
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub DoDisconnectAndClearLists()
On Error Resume Next
    mobjCDBSQL.DisConnectFromSQLServer
    cboDatabase.Clear
    optConnType(0).Enabled = True
    optConnType(1).Enabled = True
        
    If optConnType(0).Value Then
        EnableTextBox txtUserName
        EnableTextBox txtServerName
        EnableTextBox txtPassword
    Else
        DisableTextBox txtUserName
        DisableTextBox txtServerName
        DisableTextBox txtPassword
    End If
    
    txtBackupFolder = vbNullString
    txtBackupSetName = vbNullString
    txtSQLBackupFileName = vbNullString
    mblnConnected = False
    cmdConnect.Caption = "&Connect"
    cmdRestore.Enabled = False
    cmdViewDatabaseInfo.Enabled = False
    cmdBackup.Enabled = False
    cmdVerify.Enabled = False
    cmdScheduleBackup.Enabled = False
End Sub

Private Sub cmdConnect_Click()
    On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Sub cmdConnect_Click()"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    '================================ MAIN BODY =============================================
        If txtServerName = vbNullString Then
            MsgBox "Please Provide a Valid SQL Server Name!", vbExclamation, G_APPTITLE
            Exit Sub
        End If
        
        If optConnType(1).Value = True Then
            If txtUserName = vbNullString Then
                MsgBox "Please provide a valid SQL Server Login Name. The User must have Backup rights.", vbExclamation, G_APPTITLE
                Exit Sub
            End If
        End If
        
        If mblnConnected = False Then
            DoConnectAndFetchList
        Else
            DoDisconnectAndClearLists
        End If
    '================================ MAIN BODY =============================================
    
    
CleanUp:
    On Error Resume Next
    
Out:
    
    Exit Sub
    
Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
        ReportUnexpectedError lngErrorNumber, strErrorDescription
        ShowErrorMessage strErrorDescription
    Else
        ShowErrorMessage strErrorDescription
    End If
    
    Resume CleanUp
   
End Sub

Private Sub DoConnectAndFetchList()
   On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Sub DoConnectAndFetchList()"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    ' ============================ MAIN BODY ===========================
    Dim strList As Variant
    Dim strListDevices As Variant
    Dim i
    Dim blnWindowsAuth As Boolean
    
    If mobjCDBSQL Is Nothing Then
        Set mobjCDBSQL = New CDBSQLUtil
    End If
    
    cboDatabase.Clear
    
    StatusMessage "Connecting to SQL Server and Fetching List of Databases..."
    blnWindowsAuth = optConnType(0).Value
    
    If mobjCDBSQL.ConnectToSQLServer(txtServerName, txtUserName, txtPassword, blnWindowsAuth) = True Then
        strList = mobjCDBSQL.GetListOfDatabases(False)
        For i = 0 To UBound(strList)
            cboDatabase.AddItem strList(i)
        Next i
    End If
    
    cboDatabase.ListIndex = cboDatabase.ListCount - 1
    
    optConnType(0).Enabled = False
    optConnType(1).Enabled = False
    DisableTextBox txtServerName
    DisableTextBox txtUserName
    DisableTextBox txtPassword
    cmdConnect.Caption = "&Disconnect"
    cmdViewDatabaseInfo.Enabled = True
    cmdRestore.Enabled = True
    cmdBackup.Enabled = True
    cmdVerify.Enabled = True
    cmdScheduleBackup.Enabled = True
    mblnConnected = True
    StatusMessage "Done"
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

Private Sub cmdRestore_Click()
    On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Sub cmdRestore_Click()"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    '================================ MAIN BODY =============================================
    Dim blnReplace As Boolean: blnReplace = False
    
    
    If cboDatabase = vbNullString Then
        MsgBox "You Must Connect to SQL Server and Select a database.", vbInformation, G_APPTITLE
        Exit Sub
    End If
    
    If txtSQLRestorationFileLocation = vbNullString Or cboData.Text = vbNullString Then
        MsgBox "You Must select a Previous Backup File and Select a Data Item. " & vbCrLf & _
                "Note: The Restoration Process shall completely overwrite your database. So please select " & _
                "a data item of type 'Database'. Selecting a data item type of 'Differential' shall cause the " & _
                "the resoration process to fail.", vbInformation, G_APPTITLE
        Exit Sub
    End If
    
    
    If chkReplace.Value = vbChecked Then
        blnReplace = True
    End If
    
    StatusMessage "Restoring Database " & cboDatabase.Text & " from " & txtSQLRestorationFileLocation
    
    With mobjCDBSQL
        .BackupFileLocation = Left$(txtSQLRestorationFileLocation.Text, (InStrRev(txtSQLRestorationFileLocation, "\")))
        .BackupFileName = Right$(txtSQLRestorationFileLocation, (Len(txtSQLRestorationFileLocation) - InStrRev(txtSQLRestorationFileLocation, "\")))
        .DatabaseName = cboDatabase.Text
        .RestoreDatabase blnReplace, Left$(cboData.Text, 1)
    End With
    
    StatusMessage "Successfully Restored Database!"
    '================================ MAIN BODY =============================================
    
    
CleanUp:
    On Error Resume Next
Out:
    
    Exit Sub
    
Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    StatusMessage "Restore Unsuccessfull due to errors!"
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
        ReportUnexpectedError lngErrorNumber, strErrorDescription
        ShowErrorMessage "Error Number: " & lngErrorNumber & vbCrLf & " Message: " & strErrorDescription
    Else
        ShowErrorMessage strErrorDescription
    End If
    
    Resume CleanUp


End Sub

Private Sub cmdRestoreBrowse_Click()
  On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "cmdRestoreBrowse_Click()"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    '================================ MAIN BODY =============================================
    With dlgCommon
        .Filter = "BAK Files(*.bak)|*.bak|All Files(*.*)|*.*"
        .DialogTitle = "Select Previous Backup File Location"
        .ShowOpen
        If .FileName <> vbNullString Then
            txtSQLRestorationFileLocation.Text = .FileName
        End If
    End With
    '================================ MAIN BODY =============================================
    
    
CleanUp:
    On Error Resume Next
    
Out:
    
    Exit Sub
    
Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
        ReportUnexpectedError lngErrorNumber, strErrorDescription
    Else
        ShowErrorMessage strErrorDescription
    End If
    
    Resume CleanUp
End Sub

Private Sub cmdScheduleBackup_Click()
    On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "cmdScheduleBackup_Click()"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    '================================ MAIN BODY =============================================
    If txtBackupFolder.Text = vbNullString Or txtSQLBackupFileName = vbNullString Then
        MsgBox "You Must Select a Valid Backup Folder, and Backup file name. Note that the folder location is relative to the SQL Server™", vbInformation, G_APPTITLE
        Exit Sub
    End If
    With formSchedule
        .Show vbModal
    End With
    '================================ MAIN BODY =============================================
    
    
CleanUp:
    On Error Resume Next
    
Out:
    
    Exit Sub
    
Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
        ReportUnexpectedError lngErrorNumber, strErrorDescription
        ShowErrorMessage "Error: " & strErrorDescription
    Else
        ShowErrorMessage strErrorDescription
    End If
    
    Resume CleanUp


End Sub

Private Sub cmdVerify_Click()
    On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Sub cmdVerify_Click()"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    '================================ MAIN BODY =============================================
    Dim strFileName As String
    Dim strFolderLocation As String
    
    If cboDatabase = vbNullString Then
        MsgBox "You Must Connect to SQL Server and Select a database.", vbInformation, G_APPTITLE
        Exit Sub
    End If
    
    strFolderLocation = txtBackupFolder
    strFileName = txtSQLBackupFileName
    StatusMessage "Verifying Backup File..."
    mobjCDBSQL.VerifyBackupFile strFolderLocation & strFileName, cboDatabase.Text
    StatusMessage "Vefication Complete! No Errors Found."
    '================================ MAIN BODY =============================================
    
    
CleanUp:
    On Error Resume Next
    
Out:
    
    Exit Sub
    
Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
        ReportUnexpectedError lngErrorNumber, strErrorDescription
        MsgBox "Error Number: " & lngErrorNumber & " Message: " & strErrorDescription, vbCritical, G_APPTITLE
    Else
        ShowErrorMessage strErrorDescription
    End If
    
    Resume CleanUp
End Sub

Private Sub cmdViewDatabaseInfo_Click()
    On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Sub cmdViewDatabaseInfo_Click()"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    '================================ MAIN BODY =============================================
        Dim strCurrentSize, strDBName, strDBOwner, strDataSpace, strIndexSpace, strPrimaryFilePath, dtCreateDate, strMemoryUse, arrDbFiles
        Dim i As Integer
        
        If cboDatabase.Text = vbNullString Then
            MsgBox "You Must Select a Database!", vbInformation, G_APPTITLE
        End If
        
        StatusMessage "Retrieving More Database Information..."
        mobjCDBSQL.DatabaseName = cboDatabase.Text
        mobjCDBSQL.GetDatabaseSummary strDBOwner, strCurrentSize, strDataSpace, strIndexSpace, strPrimaryFilePath, dtCreateDate, strMemoryUse, arrDbFiles
        
        With formSqlDBInfo
            .txtCurrentSize.Text = strCurrentSize
            .txtDataSpaceUsed.Text = CDbl(strDataSpace)
            .txtDBName.Text = cboDatabase.Text
            .txtDBOwner.Text = strDBOwner
            .txtIndexSpaceUsed.Text = strIndexSpace
            .txtMemoryUse.Text = strMemoryUse
            .txtPrimaryFilePath.Text = strPrimaryFilePath
            .txtCreateDate = dtCreateDate
            If Not IsEmpty(arrDbFiles) Then
                For i = 0 To UBound(arrDbFiles)
                    If Not Trim(arrDbFiles(i)) = vbNullString And Not IsEmpty(arrDbFiles(i)) Then
                        .txtFiles.Text = .txtFiles.Text & vbCrLf & Trim$(arrDbFiles(i))
                    End If
                Next i
            End If
            .Show
            StatusMessage "Done"
        End With
    '================================ MAIN BODY =============================================
    
    
CleanUp:
    On Error Resume Next
    
Out:
    
    Exit Sub
    
Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    StatusMessage "Error Attempting to Retrieve Information. Error: " & strErrorDescription
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
        ReportUnexpectedError lngErrorNumber, strErrorDescription
    Else
        ShowErrorMessage strErrorDescription
    End If
    
    Resume CleanUp

End Sub

Private Sub cmdViewSets_Click()
  On Error GoTo Handle_Error
    
    Const PROC_NAME As String = "Sub cmdViewSets_Click()"
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim blnInError As Boolean
    blnInError = False
    
    '================================ MAIN BODY =============================================
    Dim i As Integer
    Dim strList As Variant
    
    cboData.Clear
    
    If txtSQLRestorationFileLocation <> vbNullString Then
        strList = mobjCDBSQL.GetListOfBackupSetsInFile(txtSQLRestorationFileLocation)
        For i = 0 To UBound(strList)
            cboData.AddItem strList(i)
        Next
    Else
        MsgBox "Please Select a valid Backup File(.bak) Location.", vbInformation, G_APPTITLE
    End If
    cboData.ListIndex = cboData.ListCount - 1
    '================================ MAIN BODY =============================================
    
    
CleanUp:
    On Error Resume Next
    
Out:
    
    Exit Sub
    
Handle_Error:
    blnInError = True
    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description
    
    If lngErrorNumber <> G_ERR_APP_DEFINED_ERROR Then
        strErrorDescription = FormatErrorDescription(M_MODULENAME, PROC_NAME, strErrorDescription)
        ReportUnexpectedError lngErrorNumber, strErrorDescription
    Else
        ShowErrorMessage strErrorDescription
    End If
    
    Resume CleanUp
End Sub

Private Sub Form_Load()
    mblnConnected = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjCDBSQL = Nothing
    Set mobjCDBAccess = Nothing
End Sub



Private Sub optConnType_Click(Index As Integer)
    If Index = 0 Then 'windows
        
        DisableTextBox txtPassword
        DisableTextBox txtUserName
    Else
        
        EnableTextBox txtPassword
        EnableTextBox txtUserName
    End If
End Sub

Private Sub optOperation_Click(Index As Integer)
    Select Case Index
        Case 0
            fraSQLBackup.Visible = True
            fraSQLRestore.Visible = False
        Case 1
            fraSQLBackup.Visible = False
            fraSQLRestore.Visible = True
    End Select
End Sub

Private Sub txtBackupFolder_Change()
    SetDefaultBackupValues
End Sub

Private Sub txtSQLBackupFileName_Change()
    txtSQLBackupFileName = RemoveWhiteSpace(txtSQLBackupFileName)
End Sub

Private Sub txtSQLRestorationFileLocation_Change()
    cboData.Clear
End Sub

