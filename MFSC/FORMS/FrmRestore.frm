VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fRestore 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Restore"
   ClientHeight    =   840
   ClientLeft      =   3180
   ClientTop       =   2175
   ClientWidth     =   5055
   FillColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdCancel 
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
         Height          =   300
         Left            =   2450
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
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
         Top             =   240
         Width           =   2175
      End
      Begin MSComDlg.CommonDialog Dialog 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
   End
End
Attribute VB_Name = "fRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    ' RESTORE TO HARDRIVE STARTS
    On Error GoTo RestoreError
    Dim Ans
    Dialog.Filter = "Backup files (*.bck) |*.bck|"
    Dialog.ShowOpen
    If Dialog.FileName <> "" Then
        Ans = MsgBox("Are you sure you want to restore the database", vbExclamation + vbYesNo, " Warning:End-User" + UserName)
        If Ans = vbYes Then
            Unload fCrew
            Unload fIngredients
            Unload fMenu
            Unload fPurchaseOrders
            Unload fReceivingOrders
            Unload fReceivingOrdersPosting
            Unload fSalesOrders
            Unload fSalesOrdersPosting
            Unload fSuppliers
            Unload fTables
            FileCopy Dialog.FileName, App.Path + "\DATABASE\MFSC.MDB"
            MsgBox "Database has been fully restored.", vbOKOnly + vbCritical, " Warning:End-User" + UserName
        End If
    End If
    Exit Sub
RestoreError:
  MsgBox Err.Description, vbOKOnly + vbCritical, " Warning:End-User" + UserName
End Sub

Private Sub cmdCancel_Click()
    ' EXIT MODULE
    Unload Me
End Sub

