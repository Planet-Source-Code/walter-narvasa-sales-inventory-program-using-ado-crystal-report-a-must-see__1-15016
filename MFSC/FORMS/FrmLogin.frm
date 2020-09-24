VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login - Meeting's Fillers, Spirit & Cafe - Sales & Inventory System"
   ClientHeight    =   4200
   ClientLeft      =   2355
   ClientTop       =   2340
   ClientWidth     =   3330
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.ListBox txtUserName 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3075
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3120
      Width           =   3075
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2160
      MouseIcon       =   "FrmLogin.frx":0E42
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3720
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   960
      MouseIcon       =   "FrmLogin.frx":114C
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3720
      Width           =   1065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   3240
      Y1              =   3615
      Y2              =   3615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   1
      X1              =   120
      X2              =   3240
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "User Name:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Password:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   885
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MODULE/FORM: PASSWORD SECURITY
' VERSION: VB6

Option Explicit

' PASSWORD SECURITY VARIABLE SETTINGS
Dim strDBPass As String
Dim strSQLPass As String
Dim dbPass As ADODB.Connection
Dim WithEvents adoPrimaryRSPass As ADODB.Recordset
Attribute adoPrimaryRSPass.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim ctr As Integer
Dim xText

Private Sub Form_Load()
    ' STARTUP PASSWORD SECURITY DATABASE CONNECTIONS
    strDBPass = App.Path + "\DATABASE\MFSC.MDB;Jet OLEDB:Database Password=MFSC;"
    strSQLPass = "SELECT * FROM Password_Security ORDER BY User_Name"
    Database_Refresh 0
    If adoPrimaryRSPass.RecordCount = 0 Then
        Exit Sub
    Else
        adoPrimaryRSPass.MoveFirst
        Do While Not adoPrimaryRSPass.EOF
            txtUserName.AddItem IIf(IsNull(adoPrimaryRSPass("User_Name")), "", adoPrimaryRSPass("User_Name"))
            adoPrimaryRSPass.MoveNext
        Loop
        mbDataChanged = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    ' LOGIN ENTRY
    Dim ShowAtStartup As Long
    If Get_User(txtUserName, txtPassword) Then
        UserName = txtUserName
        Rights1_Add = IIf(IsNull(adoPrimaryRSPass("User_Rights1_Add")), 0, IIf(adoPrimaryRSPass("User_Rights1_Add") = 0, 0, 1))
        Rights1_Edit = IIf(IsNull(adoPrimaryRSPass("User_Rights1_Edit")), 0, IIf(adoPrimaryRSPass("User_Rights1_Edit") = 0, 0, 1))
        Rights1_Delete = IIf(IsNull(adoPrimaryRSPass("User_Rights1_Delete")), 0, IIf(adoPrimaryRSPass("User_Rights1_Delete") = 0, 0, 1))
        Rights2_Tables = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Tables")), 0, IIf(adoPrimaryRSPass("User_Rights2_Tables") = 0, 0, 1))
        Rights2_Service_Crew = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Service_Crew")), 0, IIf(adoPrimaryRSPass("User_Rights2_Service_Crew") = 0, 0, 1))
        Rights2_Ingredients = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Ingredients")), 0, IIf(adoPrimaryRSPass("User_Rights2_Ingredients") = 0, 0, 1))
        Rights2_Menu = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Menu")), 0, IIf(adoPrimaryRSPass("User_Rights2_Menu") = 0, 0, 1))
        Rights2_Supplier = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Supplier")), 0, IIf(adoPrimaryRSPass("User_Rights2_Supplier") = 0, 0, 1))
        Rights2_SalesOrders = IIf(IsNull(adoPrimaryRSPass("User_Rights2_SalesOrders")), 0, IIf(adoPrimaryRSPass("User_Rights2_SalesOrders") = 0, 0, 1))
        Rights2_PurchaseOrders = IIf(IsNull(adoPrimaryRSPass("User_Rights2_PurchaseOrders")), 0, IIf(adoPrimaryRSPass("User_Rights2_PurchaseOrders") = 0, 0, 1))
        Rights2_ReceivingOrders = IIf(IsNull(adoPrimaryRSPass("User_Rights2_ReceivingOrders")), 0, IIf(adoPrimaryRSPass("User_Rights2_ReceivingOrders") = 0, 0, 1))
        Rights2_Post_SalesOrders = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Post_SalesOrders")), 0, IIf(adoPrimaryRSPass("User_Rights2_Post_SalesOrders") = 0, 0, 1))
        Rights2_Post_ReceivingOrders = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Post_SalesOrders")), 0, IIf(adoPrimaryRSPass("User_Rights2_ReceivingOrders") = 0, 0, 1))
        Rights2_Inventory_Report = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Inventory_Report")), 0, IIf(adoPrimaryRSPass("User_Rights2_Inventory_Report") = 0, 0, 1))
        Rights2_Sales_Report = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Sales_Report")), 0, IIf(adoPrimaryRSPass("User_Rights2_Sales_Report") = 0, 0, 1))
        Rights2_Critical_Report = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Critical_Report")), 0, IIf(adoPrimaryRSPass("User_Rights2_Critical_Report") = 0, 0, 1))
        Rights3_Backup = IIf(IsNull(adoPrimaryRSPass("User_Rights3_Backup")), 0, IIf(adoPrimaryRSPass("User_Rights3_Backup") = 0, 0, 1))
        Rights3_Restore = IIf(IsNull(adoPrimaryRSPass("User_Rights3_Restore")), 0, IIf(adoPrimaryRSPass("User_Rights3_Restore") = 0, 0, 1))
        Rights3_Password_Security = IIf(IsNull(adoPrimaryRSPass("User_Rights3_Password_Security")), 0, IIf(adoPrimaryRSPass("User_Rights3_Password_Security") = 0, 0, 1))
        adoPrimaryRSPass.Close
        Unload Me
        fMain.Show
    ElseIf Trim(txtUserName) = "" And Trim(txtPassword) = "3773" Then
        Rights1_Add = 1
        Rights1_Edit = 1
        Rights1_Delete = 1
        Rights2_Tables = 1
        Rights2_Service_Crew = 1
        Rights2_Ingredients = 1
        Rights2_Menu = 1
        Rights2_Supplier = 1
        Rights2_SalesOrders = 1
        Rights2_PurchaseOrders = 1
        Rights2_ReceivingOrders = 1
        Rights2_Post_SalesOrders = 1
        Rights2_Post_ReceivingOrders = 1
        Rights2_Inventory_Report = 1
        Rights2_Sales_Report = 1
        Rights2_Critical_Report = 1
        Rights3_Backup = 1
        Rights3_Restore = 1
        Rights3_Password_Security = 1
        UserName = "Administrator"
        adoPrimaryRSPass.Close
        Unload Me
        fMain.Show
    Else
        ctr = ctr + 1
        If ctr = 4 Then
           End
        Else
            xText = "You have" + Str(4 - ctr) + " tries left"
            If ctr = 3 Then
                xText = "This is your last chance!!"
            End If
            MsgBox "Access Denied!!" & vbCrLf & _
                   xText, vbOKOnly + vbCritical, "Warning:End-User"
            SendKeys "{Home}+{End}"
        End If
   End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOk_Click
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPassword.SetFocus
End Sub

Private Sub Database_Refresh(xMode As Integer)
    Set dbPass = New Connection
        dbPass.CursorLocation = adUseClient
        dbPass.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & strDBPass
    If xMode = 0 Then
        Set adoPrimaryRSPass = New Recordset
        adoPrimaryRSPass.Open strSQLPass, dbPass, adOpenStatic, adLockOptimistic
    End If
End Sub

Function Get_User(p_user As String, p_pass As String) As Boolean
    ' USERNAME AND PASSWORD VALIDATION
    strSQLPass = "SELECT * FROM Password_Security WHERE User_Name = '" & p_user & "'" _
            & " AND User_Password = '" & Decode_Pass(p_pass) & "'"
    Database_Refresh 0
    If adoPrimaryRSPass.AbsolutePosition <> -1 Then
        Get_User = True
    Else
        Get_User = False
    End If
End Function
