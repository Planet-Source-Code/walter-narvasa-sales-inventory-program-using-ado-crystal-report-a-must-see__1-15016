VERSION 5.00
Begin VB.Form fHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Program Guide"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
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
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   8655
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.TextBox txtFields 
         DataField       =   "Program_Guide"
         Height          =   5535
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   8415
      End
   End
End
Attribute VB_Name = "fHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MODULE/FORM: HELP/PROGRAM GUIDE
' VERSION: VB6

' HELP/PROGRAM GUIDE VARIABLE SETTINGS
Dim strDB As String
Dim db As ADODB.Connection
Dim strSQL As String
Dim WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1

Private Sub Form_Load()
    ' STARTUP HELP/PROGRAM GUIDE DATABASE CONNECTIONS
    strDB = App.Path + "\DATABASE\MFSC.MDB;Jet OLEDB:Database Password=MFSC;"
    strSQL = "SELECT Program_Guide FROM Help_Guide"
    Database_Refresh 0
    For Each oText In Me.txtFields
        Set oText.DataSource = adoPrimaryRS
    Next
End Sub

Private Sub cmdOk_Click()
    ' CANCEL POSTING THEN EXIT
    Unload Me
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
    End If
End Sub
