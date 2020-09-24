VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fPasswordSecurity 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Password Security"
   ClientHeight    =   6315
   ClientLeft      =   690
   ClientTop       =   780
   ClientWidth     =   8910
   Icon            =   "FrmPasswordSecurity.frx":0000
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
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   11113
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   192
      TabCaption(0)   =   "Password Security Folder"
      TabPicture(0)   =   "FrmPasswordSecurity.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblStatus"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CDlgExcel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "picButtons"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdFirst"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdPrevious"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdNext"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdLast"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   8400
         MouseIcon       =   "FrmPasswordSecurity.frx":0028
         MousePointer    =   99  'Custom
         Picture         =   "FrmPasswordSecurity.frx":0332
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   8055
         MouseIcon       =   "FrmPasswordSecurity.frx":0674
         MousePointer    =   99  'Custom
         Picture         =   "FrmPasswordSecurity.frx":097E
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   465
         MouseIcon       =   "FrmPasswordSecurity.frx":0CC0
         MousePointer    =   99  'Custom
         Picture         =   "FrmPasswordSecurity.frx":0FCA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   120
         MouseIcon       =   "FrmPasswordSecurity.frx":130C
         MousePointer    =   99  'Custom
         Picture         =   "FrmPasswordSecurity.frx":1616
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.PictureBox picButtons 
         Height          =   465
         Left            =   120
         ScaleHeight     =   405
         ScaleWidth      =   8535
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   5720
         Width           =   8595
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
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
            TabIndex        =   29
            Top             =   50
            Width           =   1335
         End
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
            Left            =   2880
            TabIndex        =   28
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
            Left            =   5760
            TabIndex        =   30
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
            Left            =   1440
            TabIndex        =   27
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
            TabIndex        =   39
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
            TabIndex        =   40
            Top             =   50
            Visible         =   0   'False
            Width           =   1335
         End
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
         Height          =   2175
         Left            =   240
         TabIndex        =   37
         Top             =   3120
         Width           =   8610
         Begin VB.CheckBox chkFields 
            Caption         =   "Password Security"
            DataField       =   "User_Rights3_Password_Security"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   18
            Left            =   6480
            TabIndex        =   22
            Top             =   1320
            Width           =   2055
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Restore"
            DataField       =   "User_Rights3_Restore"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   17
            Left            =   6480
            TabIndex        =   21
            Top             =   960
            Width           =   1095
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Inventory Report"
            DataField       =   "User_Rights2_Inventory_Report"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   13
            Left            =   4080
            TabIndex        =   17
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Backup"
            DataField       =   "User_Rights3_Backup"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   16
            Left            =   6480
            TabIndex        =   20
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Critical Report"
            DataField       =   "User_Rights2_Critical_Report"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   15
            Left            =   6480
            TabIndex        =   19
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Sales Report"
            DataField       =   "User_Rights2_Sales_Report"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   14
            Left            =   4080
            TabIndex        =   18
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Post Receiving Orders"
            DataField       =   "User_Rights2_Post_ReceivingOrders"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   12
            Left            =   4080
            TabIndex        =   16
            Top             =   960
            Width           =   2295
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Post Sales Orders"
            DataField       =   "User_Rights2_Post_SalesOrders"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   11
            Left            =   4080
            TabIndex        =   15
            Top             =   600
            Width           =   1935
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Receiving Orders"
            DataField       =   "User_Rights2_ReceivingOrders"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   4080
            TabIndex        =   14
            Top             =   240
            Width           =   2175
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Purchase Orders"
            DataField       =   "User_Rights2_PurchaseOrders"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   2160
            TabIndex        =   13
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Sales Orders"
            DataField       =   "User_Rights2_SalesOrders"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   2160
            TabIndex        =   12
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Supplier Info."
            DataField       =   "User_Rights2_Supplier"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   2160
            TabIndex        =   11
            Top             =   960
            Width           =   2055
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Menu Info."
            DataField       =   "User_Rights2_Menu"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   2160
            TabIndex        =   10
            Top             =   600
            Width           =   1935
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Ingredients Info."
            DataField       =   "User_Rights2_Ingredients"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   2160
            TabIndex        =   9
            Top             =   290
            Width           =   2415
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Service Crew Info."
            DataField       =   "User_Rights2_Service_Crew"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   8
            Top             =   1680
            Width           =   2535
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Tables Info."
            DataField       =   "User_Rights2_Tables"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   7
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Delete Record"
            DataField       =   "User_Rights1_Delete"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   6
            Top             =   960
            Width           =   1575
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Edit Record"
            DataField       =   "User_Rights1_Edit"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   5
            Top             =   600
            Width           =   1455
         End
         Begin VB.CheckBox chkFields 
            Caption         =   "Add Record"
            DataField       =   "User_Rights1_Add"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   290
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
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
         Height          =   2775
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   8610
         Begin VB.TextBox txtFields 
            DataField       =   "User_Description"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "M/d/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1245
            Index           =   3
            Left            =   2160
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1320
            Width           =   6255
         End
         Begin VB.TextBox txtFields 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "M/d/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   2160
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txtFields 
            DataField       =   "User_Password"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "M/d/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   2160
            PasswordChar    =   "*"
            TabIndex        =   1
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtFields 
            DataField       =   "User_Name"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "M/d/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   2160
            TabIndex        =   0
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Description"
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
            Left            =   1020
            TabIndex        =   36
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Re-Enter Password"
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
            Left            =   360
            TabIndex        =   35
            Top             =   960
            Width           =   1635
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Password"
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
            Left            =   1170
            TabIndex        =   34
            Top             =   600
            Width           =   825
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "User Name"
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
            Left            =   1050
            TabIndex        =   33
            Top             =   255
            Width           =   945
         End
      End
      Begin MSComDlg.CommonDialog CDlgExcel 
         Left            =   1920
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   810
         TabIndex        =   41
         Top             =   5400
         Width           =   7215
      End
   End
End
Attribute VB_Name = "fPasswordSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MODULE/FORM: PASSWORD SECURITY UTILITIES
' VERSION: VB6

Dim strDB As String
Dim strSQL As String
Dim db As ADODB.Connection
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Public strSQL2 As String
Public WithEvents adoPrimaryRS2 As Recordset
Attribute adoPrimaryRS2.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim oText As TextBox
Dim xDeleteLogic As Boolean
Dim intColIdx As Integer
Dim blnListShow As Boolean
Dim intKeyCode As Integer
Dim xButton As Integer

Private Sub Form_Load()
    ' STARTUP MODULE FOR PASSWORD SECURITY
    Dim oText As TextBox, oCheckBox As CheckBox
    blnListShow = False
    strDB = App.Path + "\DATABASE\MFSC.MDB;Jet OLEDB:Database Password=MFSC;"
    strSQL = "SELECT * FROM Password_Security ORDER BY User_Name"
    Database_Refresh 0
    For Each oText In Me.txtFields
        Set oText.DataSource = adoPrimaryRS
    Next
    For Each oCheckBox In Me.chkFields
        Set oCheckBox.DataSource = adoPrimaryRS
    Next
    Me.Height = 6720
    Me.Width = 9030
    Call DisplayRestrictions
    SetButtons True
    xDeleteLogic = False
    EditClicked = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Controlling the Buttons for adoPrimaryRS Recordset.
    If mbEditFlag Or mbAddNewFlag Then Exit Sub
    Select Case KeyCode
            Case vbKeyEscape
                    cmdClose_Click
            Case vbKeyEnd
                    cmdLast_Click
            Case vbKeyHome
                    cmdFirst_Click
            Case vbKeyUp, vbKeyPageUp
                    If Shift = vbCtrlMask Then
                        cmdFirst_Click
                    Else
                        cmdPrevious_Click
                    End If
            Case vbKeyDown, vbKeyPageDown
                    If Shift = vbCtrlMask Then
                        cmdLast_Click
                    Else
                        cmdNext_Click
                    End If
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' MOUSE NATURE DEFAULT
    Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    ' adoPrimaryRS RECORD NUMBER
    On Error Resume Next
    lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    ' ADDITIONAL BUT OPTIONAL VALIDATIONS
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
    ' ADD/NEW BUTTON
    On Error GoTo AddErr
    If Rights1_Add = 1 Then
        With adoPrimaryRS
            If Not (.BOF And .EOF) Then
                mvBookMark = .Bookmark
            End If
            .AddNew
            mbAddNewFlag = True
            SetButtons False
        End With
        xDeleteLogic = True
        EditClicked = True
        For i = 0 To 3
            txtFields(i).Enabled = True
        Next i
        txtFields(0).SetFocus
        For i = 0 To 18
            chkFields(i).Enabled = True
        Next i
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
    Exit Sub
AddErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub cmdEdit_Click()
    ' EDIT BUTTON
    On Error GoTo EditErr
    If Rights1_Edit = 1 Then
        mbEditFlag = True
        SetButtons False
        xDeleteLogic = False
        EditClicked = True
        txtFields(1) = UnCode_Pass(txtFields(1))
        For i = 1 To 3
            txtFields(i).Enabled = True
        Next i
        txtFields(1).SetFocus
        For i = 0 To 18
            chkFields(i).Enabled = True
        Next i
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
    Exit Sub
EditErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub cmdDelete_Click()
    ' DELETE BUTTON
    On Error GoTo DeleteErr
    If Rights1_Delete = 1 Then
        Msg = MsgBox("Do you want to delete User Name " & txtFields(0) & "?", vbYesNo + vbExclamation + vbDefaultButton2, _
                      "Warning:End-User:" + UserName)
        If Msg = vbYes Then
            With adoPrimaryRS
                .Delete
                .MoveNext
                If .EOF Then .MoveLast
            End With
        End If
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
    Exit Sub
DeleteErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub cmdRefresh_Click()
    ' REFRESH BUTTON
    On Error GoTo RefreshErr
    adoPrimaryRS.Requery
    Call DisplayRestrictions
    Exit Sub
RefreshErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub cmdCancel_Click()
    ' UNDO BUTTON
    On Error Resume Next
    SetButtons True
    mbEditFlag = False
    mbAddNewFlag = False
    adoPrimaryRS.CancelUpdate
    If mvBookMark > 0 Then
        adoPrimaryRS.Bookmark = mvBookMark
    Else
        adoPrimaryRS.MoveFirst
    End If
    For i = 0 To 3
        txtFields(i).Enabled = False
    Next i
    For i = 0 To 18
        chkFields(i).Enabled = False
    Next i
    mbDataChanged = False
    EditClicked = False
End Sub

Private Sub cmdUpdate_Click()
    ' SAVE BUTTON
    'On Error GoTo UpdateErr
    Dim ActiveBlankFields As String
    If txtFields(0) = "" Then
        ActiveBlankFields = ActiveBlankFields + "User Name"
    Else
        ActiveBlankFields = ""
    End If
    If txtFields(1) = "" Then
        If txtFields(0) = "" Then
            ActiveBlankFields = ActiveBlankFields + ",Password"
        Else
            ActiveBlankFields = ActiveBlankFields + "Password"
        End If
    Else
        ActiveBlankFields = ""
    End If
    If txtFields(3) = "" Then
        If txtFields(1) = "" Then
            ActiveBlankFields = ActiveBlankFields + ",Description"
        Else
            ActiveBlankFields = ActiveBlankFields + "Description"
        End If
    Else
        ActiveBlankFields = ""
    End If
    If ActiveBlankFields = "" Then
        If EditClicked = True Then
            txtFields(1) = Decode_Pass(txtFields(1))
        End If
        adoPrimaryRS.UpdateBatch adAffectAll
        If mbAddNewFlag Then
            adoPrimaryRS.MoveLast              'move to the new record
        End If
        mbEditFlag = False
        mbAddNewFlag = False
        SetButtons True
        mbDataChanged = False
        If xGridLogic = True Then
            cmdRefresh_Click
            xGridLogic = False
        End If
        For i = 0 To 3
            txtFields(i).Enabled = False
        Next i
        For i = 0 To 18
            chkFields(i).Enabled = False
        Next i
        EditClicked = False
    Else
        MsgBox ActiveBlankFields & " is empty!!", vbOKOnly + vbCritical, " Warning:End-User" + UserName
    End If
    Exit Sub
UpdateErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub cmdClose_Click()
    ' CLOSE BUTTON - EXIT
    Unload Me
End Sub

Private Sub cmdFirst_Click()
    ' TOP/FIRST BUTTON
    On Error Resume Next
    Dim Msg
    adoPrimaryRS.MoveFirst
    mbDataChanged = False
    Exit Sub
GoFirstError:
    Msg = MsgBox(Err.Description, vbOKOnly, "Validation:End-User")
End Sub

Private Sub cmdLast_Click()
    ' BOTTOM/LAST BUTTON
    On Error Resume Next
    Dim Msg
    adoPrimaryRS.MoveLast
    mbDataChanged = False
    Exit Sub
GoLastError:
    Msg = MsgBox(Err.Description, vbOKOnly, "Validation:End-User")
End Sub

Private Sub cmdNext_Click()
    ' NEXT BUTTON
    On Error Resume Next
    Dim Msg
    If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
    If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
        Beep
        adoPrimaryRS.MoveLast
    End If
    mbDataChanged = False
    Exit Sub
GoNextError:
    Msg = MsgBox(Err.Description, vbOKOnly, "Validation:End-User")
End Sub

Private Sub cmdPrevious_Click()
    ' PREVIOUS BUTTON
    On Error Resume Next
    Dim Msg
    If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
    If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
        Beep
        adoPrimaryRS.MoveFirst
    End If
    mbDataChanged = False
    Exit Sub
GoPrevError:
    Msg = MsgBox(Err.Description, vbOKOnly, "Validation:End-User")
End Sub

Private Sub SetButtons(bVal As Boolean)
    ' COMMAND BUTTONS ENABLED MODES
    On Error GoTo ErrorSetButtons
    cmdAdd.Visible = bVal
    cmdEdit.Visible = bVal
    cmdUpdate.Visible = Not bVal
    cmdCancel.Visible = Not bVal
    cmdDelete.Visible = bVal
    cmdClose.Visible = bVal
    cmdNext.Visible = bVal
    cmdFirst.Visible = bVal
    cmdLast.Visible = bVal
    cmdPrevious.Visible = bVal
    Exit Sub
ErrorSetButtons:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Public Sub Database_Refresh(xMode As Integer)
    ' DATABASE CONNECTIVITY SETTINGS
    On Error Resume Next
    Set db = New Connection
        db.CursorLocation = adUseClient
        db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & strDB
    If xMode = 0 Then
        Set adoPrimaryRS = New Recordset
        adoPrimaryRS.Open strSQL, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 1 Then
        Set adoPrimaryRS2 = New Recordset
        adoPrimaryRS2.Open strSQL2, db, adOpenStatic, adLockOptimistic
    End If
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
    ' txtFields(0) VALIDATION KEY INPUTTED
    On Error GoTo ErrorTxtFieldsFocus
    If Index = 0 Then
        If Get_User_Name Then
                Msg = MsgBox("User Name already exist!!", vbOKOnly + vbCritical, "Warning:End-User:" + UserName)
                txtFields(0) = ""
                txtFields(0).SetFocus
        ElseIf txtFields(0) = "" Then
                Msg = MsgBox("User Name cannot be empty!!", vbOKOnly + vbCritical, "Warning:End-User:" + UserName)
                cmdCancel_Click
        End If
    ElseIf Index = 2 Then
        If txtFields(1) <> txtFields(2) Then
            Msg = MsgBox("Password does not match!!" & vbCrLf & "Please Re-Enter your Password.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName)
            txtFields(1) = ""
            txtFields(2) = ""
            txtFields(1).SetFocus
        ElseIf txtFields(1) = "" Then
            Msg = MsgBox("Password cannot be empty!!", vbOKOnly + vbCritical, "Warning:End-User:" + UserName)
            txtFields(1).SetFocus
        ElseIf txtFields(2) = "" Then
            Msg = MsgBox("Re-Entered Password cannot be empty!!", vbOKOnly + vbCritical, "Warning:End-User:" + UserName)
            txtFields(2).SetFocus
        End If
    End If
    Exit Sub
ErrorTxtFieldsFocus:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Function Get_User_Name() As Boolean
    ' USER NAME VALIDATION ON txtFields(0)
    On Error Resume Next
    strSQL2 = "SELECT * FROM Password_Security WHERE User_Name = '" & txtFields(0) & "'"
    Database_Refresh 1
    If adoPrimaryRS2.AbsolutePosition <> -1 Then
        Get_User_Name = True
    Else
        Get_User_Name = False
    End If
End Function

Function FindTag_PasswordSecurity()
    ' FOR FIND BUTTON FUNCTION USE
    On Error Resume Next
    Dim oText As TextBox, oCheckBox As CheckBox
    Do While adoPrimaryRS.Fields("User_Name") <> Trim(Finder.txtWord)
        adoPrimaryRS.MoveNext
    Loop
    For Each oText In Me.txtFields
        Set oText.DataSource = adoPrimaryRS
    Next
        For Each oCheckBox In Me.chkFields
        Set oCheckBox.DataSource = adoPrimaryRS
    Next
    Call DisplayRestrictions
End Function

Function DisplayRestrictions()
    ' CHECKBOX DISPLAY RESTRICTIONS/VALUES
    If adoPrimaryRS.RecordCount <> 0 Then
        chkFields(0) = IIf(IsNull(adoPrimaryRS("User_Rights1_Add")), 0, IIf(adoPrimaryRS("User_Rights1_Add") = 0, 0, 1))
        chkFields(1) = IIf(IsNull(adoPrimaryRS("User_Rights1_Edit")), 0, IIf(adoPrimaryRS("User_Rights1_Edit") = 0, 0, 1))
        chkFields(2) = IIf(IsNull(adoPrimaryRS("User_Rights1_Delete")), 0, IIf(adoPrimaryRS("User_Rights1_Delete") = 0, 0, 1))
        chkFields(3) = IIf(IsNull(adoPrimaryRS("User_Rights2_Tables")), 0, IIf(adoPrimaryRS("User_Rights2_Tables") = 0, 0, 1))
        chkFields(4) = IIf(IsNull(adoPrimaryRS("User_Rights2_Service_Crew")), 0, IIf(adoPrimaryRS("User_Rights2_Service_Crew") = 0, 0, 1))
        chkFields(5) = IIf(IsNull(adoPrimaryRS("User_Rights2_Ingredients")), 0, IIf(adoPrimaryRS("User_Rights2_Ingredients") = 0, 0, 1))
        chkFields(6) = IIf(IsNull(adoPrimaryRS("User_Rights2_Menu")), 0, IIf(adoPrimaryRS("User_Rights2_Menu") = 0, 0, 1))
        chkFields(7) = IIf(IsNull(adoPrimaryRS("User_Rights2_Supplier")), 0, IIf(adoPrimaryRS("User_Rights2_Supplier") = 0, 0, 1))
        chkFields(8) = IIf(IsNull(adoPrimaryRS("User_Rights2_SalesOrders")), 0, IIf(adoPrimaryRS("User_Rights2_SalesOrders") = 0, 0, 1))
        chkFields(9) = IIf(IsNull(adoPrimaryRS("User_Rights2_PurchaseOrders")), 0, IIf(adoPrimaryRS("User_Rights2_PurchaseOrders") = 0, 0, 1))
        chkFields(10) = IIf(IsNull(adoPrimaryRS("User_Rights2_ReceivingOrders")), 0, IIf(adoPrimaryRS("User_Rights2_ReceivingOrders") = 0, 0, 1))
        chkFields(11) = IIf(IsNull(adoPrimaryRS("User_Rights2_Post_SalesOrders")), 0, IIf(adoPrimaryRS("User_Rights2_Post_SalesOrders") = 0, 0, 1))
        chkFields(12) = IIf(IsNull(adoPrimaryRS("User_Rights2_Post_SalesOrders")), 0, IIf(adoPrimaryRS("User_Rights2_ReceivingOrders") = 0, 0, 1))
        chkFields(13) = IIf(IsNull(adoPrimaryRS("User_Rights2_Inventory_Report")), 0, IIf(adoPrimaryRS("User_Rights2_Inventory_Report") = 0, 0, 1))
        chkFields(14) = IIf(IsNull(adoPrimaryRS("User_Rights2_Sales_Report")), 0, IIf(adoPrimaryRS("User_Rights2_Sales_Report") = 0, 0, 1))
        chkFields(15) = IIf(IsNull(adoPrimaryRS("User_Rights2_Critical_Report")), 0, IIf(adoPrimaryRS("User_Rights2_Critical_Report") = 0, 0, 1))
        chkFields(16) = IIf(IsNull(adoPrimaryRS("User_Rights3_Backup")), 0, IIf(adoPrimaryRS("User_Rights3_Backup") = 0, 0, 1))
        chkFields(17) = IIf(IsNull(adoPrimaryRS("User_Rights3_Restore")), 0, IIf(adoPrimaryRS("User_Rights3_Restore") = 0, 0, 1))
        chkFields(18) = IIf(IsNull(adoPrimaryRS("User_Rights3_Password_Security")), 0, IIf(adoPrimaryRS("User_Rights3_Password_Security") = 0, 0, 1))
    End If
End Function
