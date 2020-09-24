VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form fCrew 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Service Crew Information"
   ClientHeight    =   6315
   ClientLeft      =   1095
   ClientTop       =   405
   ClientWidth     =   8910
   Icon            =   "FrmCrew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   6300
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   11113
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   192
      TabCaption(0)   =   "&Main Entry"
      TabPicture(0)   =   "FrmCrew.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblStatus"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "picButtons"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdLast"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdNext"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdPrevious"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdFirst"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   120
         MouseIcon       =   "FrmCrew.frx":0E5E
         MousePointer    =   99  'Custom
         Picture         =   "FrmCrew.frx":1168
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   465
         MouseIcon       =   "FrmCrew.frx":14AA
         MousePointer    =   99  'Custom
         Picture         =   "FrmCrew.frx":17B4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   8050
         MouseIcon       =   "FrmCrew.frx":1AF6
         MousePointer    =   99  'Custom
         Picture         =   "FrmCrew.frx":1E00
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   8395
         MouseIcon       =   "FrmCrew.frx":2142
         MousePointer    =   99  'Custom
         Picture         =   "FrmCrew.frx":244C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Frame Frame1 
         Height          =   4995
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   8650
         Begin VB.TextBox txtFields 
            DataField       =   "Schedule"
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   2400
            TabIndex        =   3
            Top             =   3000
            Width           =   2055
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Designation"
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2400
            TabIndex        =   2
            Top             =   2640
            Width           =   3975
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Name"
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2400
            TabIndex        =   1
            Top             =   2280
            Width           =   5535
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Crew_Code"
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2400
            TabIndex        =   0
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Schedule:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   1440
            TabIndex        =   19
            Top             =   3000
            Width           =   870
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Designation:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   1230
            TabIndex        =   16
            Top             =   2640
            Width           =   1080
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1755
            TabIndex        =   15
            Top             =   2280
            Width           =   555
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Service Crew Code:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   615
            TabIndex        =   14
            Top             =   1935
            Width           =   1695
         End
      End
      Begin VB.PictureBox picButtons 
         Height          =   465
         Left            =   120
         ScaleHeight     =   405
         ScaleWidth      =   8595
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   5720
         Width           =   8650
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3600
            TabIndex        =   9
            Top             =   50
            Width           =   1335
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5040
            TabIndex        =   10
            Top             =   50
            Width           =   1335
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2160
            TabIndex        =   8
            Top             =   50
            Width           =   1335
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2880
            TabIndex        =   17
            Top             =   50
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Undo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4320
            TabIndex        =   18
            Top             =   50
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   810
         TabIndex        =   20
         Top             =   5400
         Width           =   7220
      End
   End
End
Attribute VB_Name = "fCrew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MOfDULE/FORM: SERVICE CREW INFORMATION
' VERSION: VB6

' SERVICE CREW VARIABLE SETTINGS
Dim strDB As String
Dim strSQL As String
Dim strSQL2 As String
Dim strSQL3 As String
Dim WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS2 As ADODB.Recordset
Attribute adoPrimaryRS2.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS3 As ADODB.Recordset
Attribute adoPrimaryRS3.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Form_Load()
    ' STARTUP SERVICE CREW DATABASE CONNECTIONS
    Reload_PrimaryRS
    Me.Height = 6720
    Me.Width = 9030
End Sub
    
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' PRESS ESCAPE TO CLOSE
    If mbEditFlag Or mbAddNewFlag Then Exit Sub
    Select Case KeyCode
        Case vbKeyEscape
        cmdClose_Click
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' MOUSE POINTER SET DEFAULT
    Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    ' PARAMETERIZED VALIDATIONS NOT ENABLED BUT CAN BE ADDED
    Dim bCancel As Boolean
        Select Case adReason
        Case adRsnAddNew
        Case adRsnClose
        Case adRsnDelete
        Case adRsnFirstChange
        Case adRsnMove
        Case adRsnRequery
        Case adRsnResynch
        Case adRsnUndoAddNew
        Case adRsnUndoDelete
        Case adRsnUndoUpdate
    Case adRsnUpdate
    End Select
    If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
    ' ADD BUTTON CLICK WITH AUTO-NUMBER ACCORDING TO RECORD COUNT
    On Error GoTo AddErr
    Dim i, nCount
    If Rights1_Add = 1 Then
        With adoPrimaryRS
            If Not (.BOF And .EOF) Then
                mvBookMark = .Bookmark
            End If
            .AddNew
            mbAddNewFlag = True
            SetButtons False
        End With
        For i = 0 To 3
            txtFields(i).Enabled = True
        Next i
        If adoPrimaryRS.RecordCount = 0 Then
            txtFields(0) = 1
        Else
            nCount = (adoPrimaryRS.RecordCount - 1) + 1
            txtFields(0) = nCount
        End If
        txtFields(1).SetFocus
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
    Exit Sub
AddErr:
  MsgBox Err.Description, vbOKOnly + vbCritical, " Warning:End-User" + UserName
End Sub

Private Sub cmdEdit_Click()
    ' EDIT BUTTON CLICK WITH SEARCH SERVICE CREW CODE METHOD
    On Error GoTo EditErr
    Dim xCode, i
    If Rights1_Edit = 1 Then
        If adoPrimaryRS.RecordCount <> 0 Then
            xCode = InputBox("Please Enter Service Crew Code:", " Service Crew Information - Edit Mode")
            If xCode <> "" Then
                adoPrimaryRS.MoveFirst
                Do While adoPrimaryRS.Fields("Crew_Code") <> Trim(xCode)
                    adoPrimaryRS.MoveNext
                Loop
                mbEditFlag = True
                SetButtons False
                For i = 1 To 3
                    txtFields(i).Enabled = True
                Next i
                txtFields(1).SetFocus
            Else
                Beep
            End If
        End If
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
    Exit Sub
EditErr:
  MsgBox "Service Crew Code does not exist!!", vbOKOnly + vbCritical, " Warning:End-User" + UserName
End Sub

Private Sub cmdCancel_Click()
    ' UNDO BUTTON CLICK/UNDO CHANGES
    On Error Resume Next
    Dim i
    SetButtons True
    mbEditFlag = False
    mbAddNewFlag = False
    adoPrimaryRS.CancelUpdate
    If mvBookMark > 0 Then
        adoPrimaryRS.Bookmark = mvBookMark
    Else
        adoPrimaryRS.MoveFirst
    End If
    mbDataChanged = False
    For i = 0 To 3
        txtFields(i).Enabled = False
    Next i
End Sub

Private Sub cmdUpdate_Click()
    ' SAVE BUTTON CLICK/SAVE CHANGES
    On Error GoTo UpdateErr
    Dim ActiveBlankFields As String
    If txtFields(0) = "" Then
        ActiveBlankFields = ActiveBlankFields + "Service Crew Code"
    Else
        ActiveBlankFields = ""
    End If
    If txtFields(1) = "" Then
        If txtFields(0) = "" Then
            ActiveBlankFields = ActiveBlankFields + ",Name"
        Else
            ActiveBlankFields = ActiveBlankFields + "Name"
        End If
    Else
        ActiveBlankFields = ""
    End If
    If txtFields(2) = "" Then
        If txtFields(1) = "" Then
            ActiveBlankFields = ActiveBlankFields + ",Designation"
        Else
            ActiveBlankFields = ActiveBlankFields + "Designation"
        End If
    Else
        ActiveBlankFields = ""
    End If
    If txtFields(3) = "" Then
        If txtFields(2) = "" Then
            ActiveBlankFields = ActiveBlankFields + ",Schedule"
        Else
            ActiveBlankFields = ActiveBlankFields + "Schedule"
        End If
    Else
        ActiveBlankFields = ""
    End If
    If ActiveBlankFields = "" Then
        adoPrimaryRS.UpdateBatch adAffectAll
        If mbAddNewFlag Then
            adoPrimaryRS.MoveLast
        End If
        mbEditFlag = False
        mbAddNewFlag = False
        SetButtons True
        mbDataChanged = False
        For i = 0 To 3
            txtFields(i).Enabled = False
        Next i
    Else
        MsgBox ActiveBlankFields & " is empty!!", vbOKOnly + vbCritical, " Warning:End-User" + UserName
    End If
    Exit Sub
UpdateErr:
  MsgBox Err.Description, vbOKOnly + vbCritical, " Warning:End-User" + UserName
End Sub

Private Sub cmdClose_Click()
    ' CLOSE BUTTON CLICK
    Unload Me
End Sub

Private Sub SetButtons(bVal As Boolean)
    ' CONTROL SET BUTTONS VISIBILITY
    cmdAdd.Visible = bVal
    cmdEdit.Visible = bVal
    cmdUpdate.Visible = Not bVal
    cmdCancel.Visible = Not bVal
    cmdClose.Visible = bVal
    cmdNext.Enabled = bVal
    cmdFirst.Enabled = bVal
    cmdLast.Enabled = bVal
    cmdPrevious.Enabled = bVal
End Sub

Public Sub Database_Refresh(xMode As Integer)
    ' PRE-DATABASE CONNECTION WITH PARAMETERIZED SQL VARIABLES ATTACHED IN EVERY MODE
    On Error Resume Next
    Set db = New Connection
        db.CursorLocation = adUseClient
        db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDB
    If xMode = 0 Then
        Set adoPrimaryRS = New ADODB.Recordset
        adoPrimaryRS.Open strSQL, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 1 Then
        Set adoPrimaryRS2 = New ADODB.Recordset
        adoPrimaryRS2.Open strSQL2, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 2 Then
        Set adoPrimaryRS3 = New ADODB.Recordset
        adoPrimaryRS3.Open strSQL3, db, adOpenStatic, adLockOptimistic
    End If
End Sub

Private Sub Reload_PrimaryRS()
    ' RELOADING DATA OBJECTS AND DATABASE CONNECTIONS
    On Error Resume Next
    Dim oText As TextBox, i
    strDB = App.Path + "\DATABASE\MFSC.MDB;Jet OLEDB:Database Password=MFSC;"
    strSQL = "SELECT ServiceCrew_Code AS Crew_Code,ServiceCrew_Name AS Name,ServiceCrew_Designation AS Designation,ServiceCrew_Schedule AS Schedule FROM Service_Crew ORDER BY ServiceCrew_Code"
    Database_Refresh 0
    For Each oText In Me.txtFields
        Set oText.DataSource = adoPrimaryRS
    Next
    If adoPrimaryRS.RecordCount <> 0 Then
        adoPrimaryRS.MoveFirst
        Set txtCombo.DataSource = adoPrimaryRS
        mbDataChanged = False
    End If
End Sub

Private Sub txtCombo_KeyPress(KeyAscii As Integer)
    ' DISABLING THE ALPHA/NUMERIC KEYASCII FUNCTIONS FOR txtCombo COMBO BOX
    If Index = 0 Then
        Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")
                        KeyAscii = IIf(Index = 5 Or Index = 0, 0, KeyAscii)
                Case KeyAscii = vbKeyBack
                Case Else
                     KeyAscii = 0
        End Select
  End If
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
    ' TXTFIELDS VALIDATIONS
    On Error Resume Next
    Dim Msg
    If Index = 0 Then
        If Get_ServiceCrew_Code Then
                Msg = MsgBox("Service Crew Code already exist!!", vbOKOnly + vbCritical, "Warning:End-User:")
                txtFields(0) = ""
                txtFields(0).SetFocus
        ElseIf txtFields(0) = "" Then
                Msg = MsgBox("Service Crew Code cannot be empty!!", vbOKOnly + vbCritical, "Warning:End-User:")
                cmdCancel_Click
        End If
    End If
End Sub

Function Get_ServiceCrew_Code() As Boolean
    ' SERVICE CREW CODE DUPLICATE FINDER
    strSQL2 = "SELECT ServiceCrew_Code FROM Service_Crew WHERE ServiceCrew_Code = '" & txtFields(0) & "'"
    Database_Refresh 1
    If adoPrimaryRS2.AbsolutePosition <> -1 Then
        Get_ServiceCrew_Code = True
    Else
        Get_ServiceCrew_Code = False
    End If
End Function

Private Sub cmdFirst_Click()
    ' SCROLL BUTTON TOP RECORD
    On Error GoTo GoFirstError
    adoPrimaryRS.MoveFirst
    mbDataChanged = False
    Exit Sub
GoFirstError:
  MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub cmdLast_Click()
    ' SCROLL BUTTON LAST RECORD
    On Error GoTo GoLastError
    adoPrimaryRS.MoveLast
    mbDataChanged = False
    Exit Sub
GoLastError:
  MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub cmdNext_Click()
    ' SCROLL BUTTON NEXT RECORD
    On Error GoTo GoNextError
    If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
    If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
        Beep
        adoPrimaryRS.MoveLast
    End If
    mbDataChanged = False
    Exit Sub
GoNextError:
  MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub cmdPrevious_Click()
    ' SCROLL BUTTON PREVIOUS RECORD
    On Error GoTo GoPrevError
    If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
    If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
        Beep
        adoPrimaryRS.MoveFirst
    End If
    mbDataChanged = False
    Exit Sub
GoPrevError:
  MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    ' RECORD NUMBER STATUS
    lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

