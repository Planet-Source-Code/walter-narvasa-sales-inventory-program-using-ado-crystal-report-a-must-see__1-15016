VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form fPurchaseOrders 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Orders Transaction"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "FrmPurchaseOrders.frx":0000
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
      Top             =   0
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   11113
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   192
      TabCaption(0)   =   "&Main Header Entry"
      TabPicture(0)   =   "FrmPurchaseOrders.frx":0CCA
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
         MouseIcon       =   "FrmPurchaseOrders.frx":0CE6
         MousePointer    =   99  'Custom
         Picture         =   "FrmPurchaseOrders.frx":0FF0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   465
         MouseIcon       =   "FrmPurchaseOrders.frx":1332
         MousePointer    =   99  'Custom
         Picture         =   "FrmPurchaseOrders.frx":163C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   8055
         MouseIcon       =   "FrmPurchaseOrders.frx":197E
         MousePointer    =   99  'Custom
         Picture         =   "FrmPurchaseOrders.frx":1C88
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   8400
         MouseIcon       =   "FrmPurchaseOrders.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "FrmPurchaseOrders.frx":22D4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.PictureBox picButtons 
         Height          =   465
         Left            =   120
         ScaleHeight     =   405
         ScaleWidth      =   8595
         TabIndex        =   22
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
            TabIndex        =   24
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
            TabIndex        =   23
            Top             =   50
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4995
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   8650
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            DataField       =   "PurchaseOrder_GTotals"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   1560
            TabIndex        =   3
            Top             =   4560
            Width           =   1695
         End
         Begin VB.Frame Frame2 
            Caption         =   "Purchase Orders Ingredients Details:"
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
            Height          =   3495
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   8415
            Begin VB.Timer xTransfer 
               Interval        =   1
               Left            =   4320
               Top             =   960
            End
            Begin VB.ListBox cboSearchIngredient 
               DragIcon        =   "FrmPurchaseOrders.frx":2616
               Height          =   1425
               Left            =   240
               Sorted          =   -1  'True
               TabIndex        =   21
               ToolTipText     =   " Select a city "
               Top             =   600
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.ListBox lstSearchIngredient 
               DragIcon        =   "FrmPurchaseOrders.frx":2A58
               Height          =   1425
               Left            =   2280
               Sorted          =   -1  'True
               TabIndex        =   20
               ToolTipText     =   " Select a city "
               Top             =   600
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.TextBox txtWords 
               Height          =   285
               Left            =   4320
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   600
               Visible         =   0   'False
               Width           =   2055
            End
            Begin MSDataGridLib.DataGrid grdDetails 
               Height          =   3180
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Width           =   8145
               _ExtentX        =   14367
               _ExtentY        =   5609
               _Version        =   393216
               AllowUpdate     =   -1  'True
               Enabled         =   0   'False
               HeadLines       =   1
               RowHeight       =   15
               TabAction       =   1
               RowDividerStyle =   0
               AllowAddNew     =   -1  'True
               AllowDelete     =   -1  'True
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  ScrollBars      =   2
                  AllowRowSizing  =   0   'False
                  AllowSizing     =   0   'False
                  RecordSelectors =   0   'False
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
         End
         Begin VB.ComboBox txtCombo 
            DataField       =   "PurchaseOrder_Supplier_Code"
            Enabled         =   0   'False
            Height          =   315
            Left            =   1680
            TabIndex        =   2
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtFields 
            DataField       =   "PurchaseOrder_Date"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
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
            Index           =   1
            Left            =   4560
            TabIndex        =   1
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            DataField       =   "PurchaseOrder_Number"
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   0
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Total PO Costs:"
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
            Left            =   120
            TabIndex        =   18
            Top             =   4575
            Width           =   1350
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "PO Date:"
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
            Left            =   3600
            TabIndex        =   15
            Top             =   255
            Width           =   795
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Supplier Code:"
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
            Left            =   255
            TabIndex        =   14
            Top             =   615
            Width           =   1260
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "PO Number:"
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
            Left            =   480
            TabIndex        =   13
            Top             =   285
            Width           =   1035
         End
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   810
         TabIndex        =   25
         Top             =   5400
         Width           =   7215
      End
   End
End
Attribute VB_Name = "fPurchaseOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MODULE/FORM: PURCHASE ORDERS TRANSACTION
' VERSION: VB6

' PURCHASE ORDERS VARIABLE SETTINGS
Dim strDB As String
Dim strSQL As String
Dim db As ADODB.Connection
Dim WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Public strSQL2 As String
Public WithEvents adoPrimaryRS2 As ADODB.Recordset
Attribute adoPrimaryRS2.VB_VarHelpID = -1
Public strSQL3 As String
Public WithEvents adoPrimaryRS3 As ADODB.Recordset
Attribute adoPrimaryRS3.VB_VarHelpID = -1
Public strSQL4 As String
Public WithEvents adoPrimaryRS4 As ADODB.Recordset
Attribute adoPrimaryRS4.VB_VarHelpID = -1
Public strSQL5 As String
Public WithEvents adoPrimaryRS5 As ADODB.Recordset
Attribute adoPrimaryRS5.VB_VarHelpID = -1
Public strSQL6 As String
Public WithEvents adoPrimaryRS6 As ADODB.Recordset
Attribute adoPrimaryRS6.VB_VarHelpID = -1
Public strSQL7 As String
Public WithEvents adoPrimaryRS7 As ADODB.Recordset
Attribute adoPrimaryRS7.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim oText As TextBox
Dim xGridLogic As Boolean
Dim intHeight As Integer
Dim intCol As Integer
Dim intRow As Integer

Private Sub Form_Load()
    ' STARTUP PURCHASE ORDERS TRANSACTION DATABASE CONNECTIONS/SETTINGS
    blnListShow = False
    strDB = App.Path + "\DATABASE\MFSC.MDB;Jet OLEDB:Database Password=MFSC;"
    strSQL = "SHAPE {SELECT PurchaseOrder_Number,PurchaseOrder_Date,PurchaseOrder_Supplier_Code,PurchaseOrder_GTotals FROM [PurchaseOrders_Header] ORDER BY PurchaseOrder_Number} AS ParentCMD" & _
             " APPEND ({SELECT PurchaseOrder_Number,PurchaseOrder_Ingredient_Code,PurchaseOrder_Ingredient_Name,PurchaseOrder_Ingredient_Costs,PurchaseOrder_QtyOrdered,PurchaseOrder_STotalCosts FROM [PurchaseOrders_Detail] ORDER BY PurchaseOrder_Ingredient_Code } AS ChildCMD" & _
             " RELATE PurchaseOrder_Number TO PurchaseOrder_Number) AS ChildCMD"
    Database_Refresh 0
    For Each oText In Me.txtFields
        Set oText.DataSource = adoPrimaryRS
    Next
    If adoPrimaryRS.RecordCount <> 0 Then
        adoPrimaryRS.MoveFirst
        Set grdDetails.DataSource = adoPrimaryRS("ChildCMD").UnderlyingValue
        Set txtCombo.DataSource = adoPrimaryRS
        mbDataChanged = False
        xGridLogic = False
    Else
        xGridLogic = True
    End If
    strSQL7 = "SELECT Supplier_Code FROM Suppliers ORDER BY Supplier_Code"
    Database_Refresh 6
    If adoPrimaryRS7.RecordCount <> 0 Then
        adoPrimaryRS7.MoveFirst
        Do While Not adoPrimaryRS7.EOF
            txtCombo.AddItem IIf(IsNull(adoPrimaryRS7("Supplier_Code")), "", adoPrimaryRS7("Supplier_Code"))
            adoPrimaryRS7.MoveNext
        Loop
    End If
    Call Ingredient_Initialization
    Me.Height = 6720
    Me.Width = 9030
    Call HideColumns
    Call Recalculate_Grand_Totals
    SetButtons True
    OrderEntryOpen = True
End Sub

Private Sub Form_Click()
    ' FORM WHEN CLICK cboSearchIngredient LISTBOX DISAPPEARS
    Ingredient_Validation
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' CONTROLLING THE BUTTONS FOR adoPrimaryRS ADODB.recordset
    If mbEditFlag Or mbAddNewFlag Then Exit Sub
    Select Case KeyCode
            Case vbKeyEscape
                    cmdClose_Click
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' MOUSE DEFAULT STATUS
    Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimadryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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
    ' ADD BUTTON
    On Error GoTo AddErr
    Dim nCount, i
    If Rights1_Add = 1 Then
        With adoPrimaryRS
            If Not (.BOF And .EOF) Then
                mvBookMark = .Bookmark
            End If
            .AddNew
            mbAddNewFlag = True
            SetButtons False
        End With
        For i = 0 To 2
            txtFields(i).Enabled = True
        Next i
        txtCombo.Enabled = True
        grdDetails.Enabled = True
        If adoPrimaryRS.RecordCount = 0 Then
            txtFields(0) = 1
        Else
            nCount = (adoPrimaryRS.RecordCount - 1) + 1
            txtFields(0) = nCount
        End If
        txtFields(1) = Date
        txtCombo.SetFocus
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
    Exit Sub
AddErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Public Sub cmdEdit_Click()
    ' EDIT BUTTON
    On Error GoTo EditErr
    Dim xCode, i
    If Rights1_Edit = 1 Then
        If adoPrimaryRS.RecordCount <> 0 Then
            xCode = InputBox("Please Enter PO Number:", "Purchase Orders Transaction - Edit Mode")
            If xCode <> "" Then
                adoPrimaryRS.MoveFirst
                Do While adoPrimaryRS.Fields("PurchaseOrder_Number") <> Trim(xCode)
                    adoPrimaryRS.MoveNext
                Loop
                mbEditFlag = True
                SetButtons False
                grdDetails.Enabled = True
                For i = 1 To 2
                    txtFields(i).Enabled = True
                Next i
                txtCombo.Enabled = True
                txtFields(1).SetFocus
            End If
        Else
            Beep
        End If
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
    Exit Sub
EditErr:
  MsgBox "PO Number does not exist!!", vbOKOnly + vbCritical, " Warning:End-User"
End Sub

Public Sub cmdRefresh_Click()
    ' INVISIBLE REFRESH BUTTON
    On Error GoTo RefreshErr
        Set grdDetails.DataSource = Nothing
        adoPrimaryRS.Requery
        Set grdDetails.DataSource = adoPrimaryRS("ChildCMD").UnderlyingValue
        Call HideColumns
        Call Recalculate_Grand_Totals
    Exit Sub
RefreshErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Public Sub cmdCancel_Click()
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
    For i = 0 To 2
        txtFields(i).Enabled = False
    Next i
    txtCombo.Enabled = False
    grdDetails.Enabled = False
    mbDataChanged = False
End Sub

Public Sub cmdUpdate_Click()
    ' SAVE BUTTON
    On Error GoTo UpdateErr
    Dim ActiveBlankFields As String
    If txtFields(0) = "" Then
        ActiveBlankFields = ActiveBlankFields + "PO Number"
    Else
        ActiveBlankFields = ""
    End If
    If txtFields(1) = "" Then
        If txtFields(0) = "" Then
            ActiveBlankFields = ActiveBlankFields + ",PO Date"
        Else
            ActiveBlankFields = ActiveBlankFields + "PO Date"
        End If
    Else
        ActiveBlankFields = ""
    End If
    If txtCombo = "" Then
        If txtFields(1) = "" Then
            ActiveBlankFields = ActiveBlankFields + ",Supplier Code"
        Else
            ActiveBlankFields = ActiveBlankFields + "Supplier Code"
        End If
    Else
        ActiveBlankFields = ""
    End If
    If ActiveBlankFields = "" Then
        adoPrimaryRS.UpdateBatch adAffectAll
        If mbAddNewFlag Then
            adoPrimaryRS.MoveLast              'move to the new record
        End If
        mbEditFlag = False
        mbAddNewFlag = False
        SetButtons True
        mbDataChanged = False
        For i = 0 To 2
            txtFields(i).Enabled = False
        Next i
        txtCombo.Enabled = False
        grdDetails.Enabled = False
        If xGridLogic = True Then
            cmdRefresh_Click
            xGridLogic = False
        End If
        Call Recalculate_Grand_Totals
    Else
        MsgBox ActiveBlankFields & " is empty!!", vbOKOnly + vbCritical, " Warning:End-User" + UserName
    End If
    Exit Sub
UpdateErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Public Sub cmdClose_Click()
    ' CLOSE BUTTON
    OrderEntryOpen = False
    Unload Me
End Sub

Public Sub SetButtons(bVal As Boolean)
    ' COMMAND BUTTONS VISIBLE MODE
    On Error GoTo ErrorSetButtons
    cmdAdd.Visible = bVal
    cmdEdit.Visible = bVal
    cmdUpdate.Visible = Not bVal
    cmdCancel.Visible = Not bVal
    cmdClose.Visible = bVal
    cmdNext.Enabled = bVal
    cmdFirst.Enabled = bVal
    cmdLast.Enabled = bVal
    cmdPrevious.Enabled = bVal
    Exit Sub
ErrorSetButtons:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Public Sub Database_Refresh(xMode As Integer)
    ' DATABASE CONNECTIVITY SETTINGS
    On Error Resume Next
    Set db = New Connection
        db.CursorLocation = adUseClient
        db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDB
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
    ElseIf xMode = 4 Then
        Set adoPrimaryRS5 = New ADODB.Recordset
        adoPrimaryRS5.Open strSQL5, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 5 Then
        Set adoPrimaryRS6 = New ADODB.Recordset
        adoPrimaryRS6.Open strSQL6, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 6 Then
        Set adoPrimaryRS7 = New ADODB.Recordset
        adoPrimaryRS7.Open strSQL7, db, adOpenStatic, adLockOptimistic
    End If
End Sub

Function HideColumns()
    ' HIDING OF NECESSARY COLUMNS,SIZES & ALIGNMENTS OF grdDetails DATAGRID.
    On Error Resume Next
    Dim i
    If adoPrimaryRS.RecordCount <> 0 Then
        grdDetails.Columns.Item(0).Visible = False
        grdDetails.Columns.Item(1).Caption = "Code"
        grdDetails.Columns.Item(2).Caption = "Ingredient Name"
        grdDetails.Columns.Item(3).Caption = "Costs"
        grdDetails.Columns.Item(4).Caption = "Qty"
        grdDetails.Columns.Item(5).Caption = "Total Costs"
        grdDetails.Columns.Item(1).Width = 900
        grdDetails.Columns.Item(2).Width = 2700
        grdDetails.Columns.Item(3).Width = 1700
        grdDetails.Columns.Item(4).Width = 900
        grdDetails.Columns.Item(5).Width = 1700
        grdDetails.Columns.Item(1).Button = True
        grdDetails.Columns.Item(3).NumberFormat = "###,###,###.00"
        grdDetails.Columns.Item(4).NumberFormat = "###,###,###"
        grdDetails.Columns.Item(5).NumberFormat = "###,###,###.00"
        For i = 3 To 5
            grdDetails.Columns.Item(i).Alignment = dbgRight
        Next i
        For i = 0 To 5
            grdDetails.Columns.Item(i).AllowSizing = False
        Next i
    End If
End Function

Public Sub grdDetails_ButtonClick(ByVal ColIndex As Integer)
    ' USE TO TRANSFER THE CURRENT COORDINATES OF grdDetails COLUMNS TO cboSearchIngredient LISTBOX
    On Error Resume Next
    Dim strItem As String
    With grdDetails
        strItem = .Text
        Select Case ColIndex
                Case 1
                    cboSearchIngredient.Height = (.Height / .RowHeight - (intRow - 1)) * .RowHeight
                    cboSearchIngredient.Move .Left + .Columns(1).Left, .Top + .RowTop(.Row) + .RowHeight, .Columns(4).Width
                    If Len(strItem) Then
                        cboSearchIngredient = strItem
                    Else
                        cboSearchIngredient.ListIndex = 0
                    End If
                        cboSearchIngredient.Visible = True
                        cboSearchIngredient.SetFocus
        End Select
    End With
End Sub

Public Sub grdDetails_AfterColUpdate(ByVal ColIndex As Integer)
    ' COMPUTES grdDetails.Columns.Item(3) MULTIPLIES WITH grdDetails.Columns.Item(4) = grdDetails.Columns.Item(5) ->Total Costs
    If adoPrimaryRS.RecordCount <> 0 Then
        If grdDetails.Enabled = True Then
            grdDetails.Columns.Item(5) = (IIf(IsNull(grdDetails.Columns.Item(3)), 0#, Val(grdDetails.Columns.Item(3))) * (IIf(IsNull(grdDetails.Columns.Item(4)), 0, Val(grdDetails.Columns.Item(4)))))
        End If
    End If
End Sub

Public Sub grdDetails_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    intCol = grdDetails.Col
    intRow = grdDetails.Row
    Ingredient_Validation
End Sub

Public Sub grdDetails_Scroll(Cancel As Integer)
    Ingredient_Validation
End Sub

Public Sub cboSearchIngredient_KeyDown(KeyCode As Integer, Shift As Integer)
    ' cboSearchIngredient KEYDOWN VALIDATION
    If KeyCode = vbKeyEscape Then
        cboSearchIngredient.Visible = False
    ElseIf KeyCode = vbKeyReturn Then
        grdDetails.Text = cboSearchIngredient.Text
        cboSearchIngredient.Visible = False
    Else
        SendKeys "{ENTER}"
        MsgBox ""
    End If
End Sub

Public Sub cboSearchIngredient_Click()
    ' TRANSFERRING OF cboSearchIngredient LISTBOX DATA TO a to grdDetails.Columns(1) Ingredient Code DETAIL DATA ON CLICK MODE
    On Error Resume Next
    grdDetails.Text = cboSearchIngredient
    cboSearchIngredient.Visible = False
End Sub

Public Sub cboSearchIngredient_LostFocus()
    'cboSearhIngredient LISTBOX DISAPPEARING/INVISIBLE ACT
    cboSearchIngredient.Visible = False
End Sub

Private Sub txtCombo_KeyPress(KeyAscii As Integer)
    ' DISABLING THE ALPHA/NUMERIC KEYASCII FUNCTIONS FOR txtCombo
    Select Case KeyAscii
            Case KeyAscii = vbKeyBack
            Case Else
                KeyAscii = 0
    End Select
End Sub

Public Sub txtFields_LostFocus(Index As Integer)
    ' txtFields VALIDATION KEY INPUTTED
    On Error GoTo ErrorTxtFieldsFocus
    If Index = 0 Then
        If Get_PO_Number Then
                Msg = MsgBox("PO Number already exist!!", vbOKOnly + vbCritical, "Warning:End-User:" + UserName)
                txtFields(0) = ""
                txtFields(0).SetFocus
        ElseIf txtFields(0) = "" Then
                Msg = MsgBox("PO Number cannot be empty!!", vbOKOnly + vbCritical, "Warning:End-User:" + UserName)
                cmdCancel_Click
        End If
    End If
    Exit Sub
ErrorTxtFieldsFocus:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Function Get_PO_Number() As Boolean
    ' PO NUMBER VALIDATION ON txtFields(0) TEXTBOX
    On Error Resume Next
    strSQL2 = "SELECT * FROM [PurchaseOrders_Header] WHERE PurchaseOrder_Number = '" & txtFields(0) & "'"
    Database_Refresh 1
    If adoPrimaryRS2.AbsolutePosition <> -1 Then
        Get_PO_Number = True
    Else
        Get_PO_Number = False
    End If
End Function

Public Sub txtWords_Change()
    ' INGREDIENT'S SENSITIVE KEY FILTER OR BRIDGE WHEN TYPING INTO txtWords TEXTBOX DOING AN SPECIAL ROLE
    Call QueryIngredient(txtWords.Text)
    grdDetails.Columns(1) = txtWords.Text
End Sub

Public Sub xTransfer_Timer()
    ' THIS WILL TRANSFER THE INPUTTED DATA FROM grdDetails DATAGRID TO txtWords TEXT BOX
    If grdDetails.Enabled = True Then
        If grdDetails.Row <> -1 Then
            txtWords.Text = grdDetails.Columns.Item(1)
        End If
    End If
End Sub

Public Sub QueryIngredient(reqText As String)
    ' AUTO-COMPLETE MODULE FOR CONTROLLING lstSearchIngredient ListBox related to grdDetails DataGrid.
    strSQL3 = "SELECT * FROM Ingredients WHERE Left(Ingredient_Code," & Len(reqText) & ")='" & reqText & "';"
    Database_Refresh 2
    lstSearchIngredient.Clear
    If adoPrimaryRS3.RecordCount = 0 Then
        lstSearchIngredient.AddItem "Ingredient Code not found!"
        Call Ingredient_Not_found
        grdDetails.Columns.Item(2) = ""
        grdDetails.Columns.Item(3) = ""
        Exit Sub
    End If
        adoPrimaryRS3.MoveLast: adoPrimaryRS3.MoveFirst
        Do Until adoPrimaryRS3.EOF
           lstSearchIngredient.AddItem adoPrimaryRS3("Ingredient_Code")
           adoPrimaryRS3.MoveNext
        Loop
        If lstSearchIngredient.ListCount = 1 Then
            adoPrimaryRS3.MoveFirst
            grdDetails.Columns.Item(2) = IIf(IsNull(adoPrimaryRS3("Ingredient_Name")), "", adoPrimaryRS3("Ingredient_Name"))
            grdDetails.Columns.Item(3) = IIf(IsNull(adoPrimaryRS3("Ingredient_Costs")), "", adoPrimaryRS3("Ingredient_Costs"))
            txtWords.Text = IIf(IsNull(lstSearchIngredient.List(0)), "", lstSearchIngredient.List(0))
            txtWords.SelLength = Len(txtWords.Text)
        Else
            grdDetails.Columns(2) = ""
            grdDetails.Columns.Item(3) = ""
        End If
End Sub

Public Sub Ingredient_Initialization()
    ' INITIALIZE INGREDIENTS TABLE
    strSQL3 = "SELECT Ingredient_Code FROM Ingredients ORDER BY Ingredient_Code"
    Database_Refresh 2
    cboSearchIngredient.Clear
    If adoPrimaryRS3.RecordCount <> 0 Then
        adoPrimaryRS3.MoveFirst
        Do Until adoPrimaryRS3.EOF
            cboSearchIngredient.AddItem adoPrimaryRS3("Ingredient_Code")
            adoPrimaryRS3.MoveNext
        Loop
    End If
End Sub

Public Sub Ingredient_Not_found()
    ' INGREDIENT CODE BE DIRECTED TO INGREDIENTS INFORMATION DATA ENTRY IF RESPONSE TO YES.
    Dim Msg
    Msg = MsgBox("Ingredient Code not found!!" & vbCrLf & "Do you want add this" & vbCrLf & _
                "New Ingredient to Ingredients Information?", vbYesNo + vbDefaultButton2 + vbExclamation, _
                "Warning:End-User:" + UserName)
        If Msg = vbYes Then
            fIngredients.Show
            fIngredients.SetFocus
            fIngredients.cmdAdd_Click
            fIngredients.txtFields(1).SetFocus
            fIngredients.txtFields(0) = IIf(IsNull(grdDetails.Columns.Item(1)), "", grdDetails.Columns.Item(1))
            fIngredients.txtFields(1) = IIf(IsNull(grdDetails.Columns.Item(2)), "", grdDetails.Columns.Item(2))
            fIngredients.txtFields(6) = IIf(IsNull(grdDetails.Columns.Item(3)), "", grdDetails.Columns.Item(3))
            OrderEntryModule = "fPurchaseOrders"
        End If
End Sub

Public Sub Ingredient_Validation()
    ' INGREDIENTS COMBO SEARCH BOX VISIBILITY VALIDATIONS
    If cboSearchIngredient.Visible Then
        cboSearchIngredient.Visible = False
    End If
End Sub

Public Sub Recalculate_Grand_Totals()
    ' THIS SUM/RECALCULATE GRAND TOTALS
    On Error Resume Next
    If adoPrimaryRS.RecordCount <> 0 Then
        strSQL4 = "SELECT SUM(PurchaseOrder_STotalCosts) FROM [PurchaseOrders_Detail] WHERE PurchaseOrder_Number = '" & adoPrimaryRS("PurchaseOrder_Number") & "'"
        Database_Refresh 3
        txtFields(2) = IIf(IsNull(adoPrimaryRS4(0)), 0#, Format(adoPrimaryRS4(0), "###,###,###.00"))
    End If
End Sub

Public Sub SSTab1_Click(PreviousTab As Integer)
    ' RECALCULATES GRAND TOTALS
    Call Recalculate_Grand_Totals
End Sub

Private Sub cmdFirst_Click()
    ' SCROLL BUTTON TOP RECORD
    On Error GoTo GoFirstError
    adoPrimaryRS.MoveFirst
    mbDataChanged = False
    Call Recalculate_Grand_Totals
    Exit Sub
GoFirstError:
  MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub cmdLast_Click()
    ' SCROLL BUTTON LAST RECORD
    On Error GoTo GoLastError
    adoPrimaryRS.MoveLast
    mbDataChanged = False
    Call Recalculate_Grand_Totals
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
    Call Recalculate_Grand_Totals
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
    Call Recalculate_Grand_Totals
    Exit Sub
GoPrevError:
  MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    ' RECORD NUMBER STATUS
    lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

