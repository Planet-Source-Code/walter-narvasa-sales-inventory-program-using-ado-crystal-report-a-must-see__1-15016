VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form fIngredients 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingredients Information"
   ClientHeight    =   6315
   ClientLeft      =   1095
   ClientTop       =   405
   ClientWidth     =   8910
   Icon            =   "FrmIngredients.frx":0000
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
      TabIndex        =   15
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
      TabPicture(0)   =   "FrmIngredients.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblStatus"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "picButtons"
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
         MouseIcon       =   "FrmIngredients.frx":0E5E
         MousePointer    =   99  'Custom
         Picture         =   "FrmIngredients.frx":1168
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   465
         MouseIcon       =   "FrmIngredients.frx":14AA
         MousePointer    =   99  'Custom
         Picture         =   "FrmIngredients.frx":17B4
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   8055
         MouseIcon       =   "FrmIngredients.frx":1AF6
         MousePointer    =   99  'Custom
         Picture         =   "FrmIngredients.frx":1E00
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   8400
         MouseIcon       =   "FrmIngredients.frx":2142
         MousePointer    =   99  'Custom
         Picture         =   "FrmIngredients.frx":244C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.PictureBox picButtons 
         Height          =   465
         Left            =   120
         ScaleHeight     =   405
         ScaleWidth      =   8595
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   5720
         Width           =   8650
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
            TabIndex        =   12
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
            TabIndex        =   14
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
            Left            =   3600
            TabIndex        =   13
            Top             =   50
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
            TabIndex        =   28
            Top             =   50
            Visible         =   0   'False
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
            TabIndex        =   27
            Top             =   50
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4995
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   8650
         Begin VB.Frame Frame2 
            Caption         =   "Inventory && Costing:"
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
            Height          =   1200
            Left            =   120
            TabIndex        =   22
            Top             =   3720
            Width           =   8390
            Begin VB.TextBox txtFields 
               Alignment       =   1  'Right Justify
               DataField       =   "Costs"
               Enabled         =   0   'False
               Height          =   285
               Index           =   5
               Left            =   6000
               TabIndex        =   7
               Top             =   360
               Width           =   1695
            End
            Begin VB.TextBox txtFields 
               Alignment       =   1  'Right Justify
               DataField       =   "On_Hand"
               Enabled         =   0   'False
               Height          =   285
               Index           =   3
               Left            =   2280
               TabIndex        =   5
               Top             =   360
               Width           =   1695
            End
            Begin VB.TextBox txtFields 
               Alignment       =   1  'Right Justify
               DataField       =   "Reorder"
               Enabled         =   0   'False
               Height          =   285
               Index           =   4
               Left            =   2280
               TabIndex        =   6
               Top             =   795
               Width           =   1695
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Item Costs:"
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
               Index           =   8
               Left            =   4920
               TabIndex        =   25
               Top             =   360
               Width           =   960
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Quantity On-Hand:"
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
               Index           =   7
               Left            =   600
               TabIndex        =   24
               Top             =   360
               Width           =   1590
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Quantity Reorder:"
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
               Index           =   6
               Left            =   675
               TabIndex        =   23
               Top             =   830
               Width           =   1515
            End
         End
         Begin VB.ComboBox txtCombo 
            DataField       =   "Type"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   2280
            TabIndex        =   4
            Top             =   3390
            Width           =   2535
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Name"
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2280
            TabIndex        =   1
            Top             =   840
            Width           =   5535
         End
         Begin VB.ComboBox txtCombo 
            DataField       =   "Units"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   2280
            TabIndex        =   3
            Top             =   3000
            Width           =   1935
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Description"
            Enabled         =   0   'False
            Height          =   1725
            Index           =   2
            Left            =   2280
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   1200
            Width           =   5535
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Code"
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2280
            TabIndex        =   0
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Type:"
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
            Index           =   4
            Left            =   1680
            TabIndex        =   21
            Top             =   3450
            Width           =   495
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Units:"
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
            Left            =   1665
            TabIndex        =   20
            Top             =   3060
            Width           =   510
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Description:"
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
            Left            =   1140
            TabIndex        =   19
            Top             =   1200
            Width           =   1035
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
            Left            =   1620
            TabIndex        =   18
            Top             =   885
            Width           =   555
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Ingredient Code:"
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
            Left            =   750
            TabIndex        =   17
            Top             =   525
            Width           =   1425
         End
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   810
         TabIndex        =   29
         Top             =   5400
         Width           =   7215
      End
   End
End
Attribute VB_Name = "fIngredients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MODULE/FORM: INGREDIENTS INFORMATION
' VERSION: VB6

' INGREDIENTS VARIABLE SETTINGS
Dim strDB As String
Dim strSQL As String
Dim strSQL2 As String
Dim strSQL3 As String
Dim strSQL4 As String
Dim WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS2 As ADODB.Recordset
Attribute adoPrimaryRS2.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS3 As ADODB.Recordset
Attribute adoPrimaryRS3.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS4 As ADODB.Recordset
Attribute adoPrimaryRS4.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Form_Load()
    ' STARTUP INGREDIENTS DATABASE CONNECTIONS
    Reload_PrimaryRS
    strSQL3 = "SELECT Units FROM Units_of_Measure ORDER BY Units"
    Database_Refresh 2
    If adoPrimaryRS3.RecordCount <> 0 Then
        adoPrimaryRS3.MoveFirst
        Do While Not adoPrimaryRS3.EOF
            txtCombo(0).AddItem IIf(IsNull(adoPrimaryRS3("Units")), "", adoPrimaryRS3("Units"))
            adoPrimaryRS3.MoveNext
        Loop
    End If
    strSQL4 = "SELECT Ingredients_Type FROM Type_of_Ingredients ORDER BY Ingredients_Type"
    Database_Refresh 3
    If adoPrimaryRS4.RecordCount <> 0 Then
        adoPrimaryRS4.MoveFirst
        Do While Not adoPrimaryRS4.EOF
            txtCombo(1).AddItem IIf(IsNull(adoPrimaryRS4("Ingredients_Type")), "", adoPrimaryRS4("Ingredients_Type"))
            adoPrimaryRS4.MoveNext
        Loop
    End If
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
    If OrderEntryOpen = True Then
        If OrderEntryModule = "fMenu" Then
            fMenu.grdDetails.Columns(1) = fIngredients.txtFields(0)
            fMenu.grdDetails.Columns(2) = fIngredients.txtFields(1)
            fMenu.grdDetails.Columns(3) = fIngredients.txtFields(5)
        ElseIf OrderEntryModule = "fPurchaseOrders" Then
            fPurchaseOrders.grdDetails.Columns(1) = fIngredients.txtFields(0)
            fPurchaseOrders.grdDetails.Columns(2) = fIngredients.txtFields(1)
            fPurchaseOrders.grdDetails.Columns(3) = fIngredients.txtFields(5)
        ElseIf OrderEntryModule = "fReceivingOrders" Then
            fReceivingOrders.grdDetails.Columns(1) = fIngredients.txtFields(0)
            fReceivingOrders.grdDetails.Columns(2) = fIngredients.txtFields(1)
            fReceivingOrders.grdDetails.Columns(3) = fIngredients.txtFields(5)
        End If
    End If
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

Public Sub cmdAdd_Click()
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
        For i = 0 To 5
            txtFields(i).Enabled = True
        Next i
        For i = 0 To 1
            txtCombo(i).Enabled = True
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

Public Sub cmdEdit_Click()
    ' EDIT BUTTON CLICK WITH SEARCH INGREDIENT CODE METHOD
    On Error GoTo EditErr
    Dim xCode, i
    If Rights1_Edit = 1 Then
        If adoPrimaryRS.RecordCount <> 0 Then
            xCode = InputBox("Please Enter Ingredient Code:", " Ingredients Information - Edit Mode")
            If xCode <> "" Then
                adoPrimaryRS.MoveFirst
                Do While adoPrimaryRS.Fields("Code") <> Trim(xCode)
                    adoPrimaryRS.MoveNext
                Loop
                mbEditFlag = True
                SetButtons False
                For i = 1 To 5
                    txtFields(i).Enabled = True
                Next i
                For i = 0 To 1
                    txtCombo(i).Enabled = True
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
  MsgBox "Ingredient Code does not exist!!", vbOKOnly + vbCritical, " Warning:End-User" + UserName
End Sub

Public Sub cmdCancel_Click()
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
    For i = 0 To 5
        txtFields(i).Enabled = False
    Next i
    For i = 0 To 1
        txtCombo(i).Enabled = False
    Next i
End Sub

Public Sub cmdUpdate_Click()
    ' SAVE BUTTON CLICK/SAVE CHANGES
    On Error GoTo UpdateErr
    Dim ActiveBlankFields As String
    If txtFields(0) = "" Then
        ActiveBlankFields = ActiveBlankFields + "Ingredient Code"
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
            ActiveBlankFields = ActiveBlankFields + ",Description"
        Else
            ActiveBlankFields = ActiveBlankFields + "Description"
        End If
    Else
        ActiveBlankFields = ""
    End If
    If txtCombo(0) = "" Then
        If txtFields(2) = "" Then
            ActiveBlankFields = ActiveBlankFields + ",Units"
        Else
            ActiveBlankFields = ActiveBlankFields + "Units"
        End If
    Else
        ActiveBlankFields = ""
    End If
    If txtCombo(1) = "" Then
        If txtCombo(0) = "" Then
            ActiveBlankFields = ActiveBlankFields + ",Type"
        Else
            ActiveBlankFields = ActiveBlankFields + "Type"
        End If
    Else
        ActiveBlankFields = ""
    End If
    If txtFields(4) = "" Then
        If txtCombo(1) = "" Then
            ActiveBlankFields = ActiveBlankFields + ",Quantity Reorder"
        Else
            ActiveBlankFields = ActiveBlankFields + "Quantity Reorder"
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
        For i = 0 To 5
            txtFields(i).Enabled = False
        Next i
        For i = 0 To 1
            txtCombo(i).Enabled = False
        Next i
        If OrderEntryOpen = True Then
            If OrderEntryModule = "fMenu" Then
                fMenu.Ingredient_Initialization
            ElseIf OrderEntryModule = "fPurchaseOrders" Then
                fPurchaseOrders.Ingredient_Initialization
            ElseIf OrderEntryModule = "fReceivingOrders" Then
                fReceivingOrders.Ingredient_Initialization
            End If
        End If
    Else
        MsgBox ActiveBlankFields & " is empty!!", vbOKOnly + vbCritical, " Warning:End-User" + UserName
    End If
    Exit Sub
UpdateErr:
  MsgBox Err.Description, vbOKOnly + vbCritical, " Warning:End-User" + UserName
End Sub

Public Sub cmdClose_Click()
    ' CLOSE BUTTON CLICK WITH DATA TRANSFER TO grdDetails
    If OrderEntryOpen = True Then
        If OrderEntryModule = "fMenu" Then
            fMenu.grdDetails.Columns(1) = fIngredients.txtFields(0)
            fMenu.grdDetails.Columns(2) = fIngredients.txtFields(1)
            fMenu.grdDetails.Columns(3) = fIngredients.txtFields(5)
        ElseIf OrderEntryModule = "fPurchaseOrders" Then
            fPurchaseOrders.grdDetails.Columns(1) = fIngredients.txtFields(0)
            fPurchaseOrders.grdDetails.Columns(2) = fIngredients.txtFields(1)
            fPurchaseOrders.grdDetails.Columns(3) = fIngredients.txtFields(5)
        ElseIf OrderEntryModule = "fReceivingOrders" Then
            fReceivingOrders.grdDetails.Columns(1) = fIngredients.txtFields(0)
            fReceivingOrders.grdDetails.Columns(2) = fIngredients.txtFields(1)
            fReceivingOrders.grdDetails.Columns(3) = fIngredients.txtFields(5)
        End If
    End If
    Unload Me
End Sub

Public Sub SetButtons(bVal As Boolean)
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
    ElseIf xMode = 3 Then
        Set adoPrimaryRS4 = New ADODB.Recordset
        adoPrimaryRS4.Open strSQL4, db, adOpenStatic, adLockOptimistic
    End If
End Sub

Public Sub Reload_PrimaryRS()
    ' RELOADING DATA OBJECTS AND DATABASE CONNECTIONS
    On Error Resume Next
    Dim oText As TextBox, i
    strDB = App.Path + "\DATABASE\MFSC.MDB;Jet OLEDB:Database Password=MFSC;"
    strSQL = "SELECT Ingredient_Code AS Code,Ingredient_Name AS Name,Ingredient_Description AS Description," & _
             "Ingredient_Units AS Units,Ingredient_Type AS Type,Ingredient_QtyRecvd AS Recvd,Ingredient_QtyReorder AS Reorder," & _
             "Ingredient_QtyOnHand AS On_Hand,Ingredient_Costs AS Costs FROM Ingredients ORDER BY Ingredient_Code"
    Database_Refresh 0
    For Each oText In Me.txtFields
        Set oText.DataSource = adoPrimaryRS
    Next
    If adoPrimaryRS.RecordCount <> 0 Then
        adoPrimaryRS.MoveFirst
        For i = 0 To 1
            Set txtCombo(i).DataSource = adoPrimaryRS
        Next i
        mbDataChanged = False
    End If
End Sub

Public Sub txtCombo_KeyPress(Index As Integer, KeyAscii As Integer)
    ' DISABLING THE ALPHA/NUMERIC KEYASCII FUNCTIONS FOR txtCombo COMBO BOX
    If Index = 0 Or Index = 1 Then
        Select Case KeyAscii
                Case KeyAscii = vbKeyBack
                Case Else
                     KeyAscii = 0
        End Select
  End If
End Sub

Public Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
    ' DISABLING THE ALPHABET KEYASCII FUNCTIONS FOR TXTFIELDS 3 TO 5
    If Index = 3 Or Index = 4 Or Index = 5 Then
        Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")
                        KeyAscii = IIf(Index = 3 Or Index = 4, 0, KeyAscii)
                Case Else
                     KeyAscii = 0
        End Select
  End If
End Sub

Public Sub txtFields_LostFocus(Index As Integer)
    ' TXTFIELDS VALIDATIONS
    On Error Resume Next
    Dim Msg
    If Index = 0 Then
        If Get_Ingredient_Code Then
                Msg = MsgBox("Ingredient Code already exist!!", vbOKOnly + vbCritical, "Warning:End-User:")
                txtFields(0) = ""
                txtFields(0).SetFocus
        ElseIf txtFields(0) = "" Then
                Msg = MsgBox("Ingredient Code cannot be empty!!", vbOKOnly + vbCritical, "Warning:End-User:")
                cmdCancel_Click
        End If
    End If
End Sub

Function Get_Ingredient_Code() As Boolean
    ' INGREDIENT CODE DUPLICATE FINDER
    strSQL2 = "SELECT Ingredient_Code FROM Ingredients WHERE Ingredient_Code = '" & txtFields(0) & "'"
    Database_Refresh 1
    If adoPrimaryRS2.AbsolutePosition <> -1 Then
        Get_Ingredient_Code = True
    Else
        Get_Ingredient_Code = False
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
