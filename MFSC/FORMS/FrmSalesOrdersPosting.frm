VERSION 5.00
Begin VB.Form fSalesOrdersPosting 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Post Sales Orders to Inventory"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Select OR Number to be Posted:"
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
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
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
         Left            =   960
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
      End
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
         Left            =   2520
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox txtCombo 
         DataField       =   "Menu_Group"
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "OR Number:"
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
         Left            =   720
         TabIndex        =   2
         Top             =   660
         Width           =   1050
      End
   End
End
Attribute VB_Name = "fSalesOrdersPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MODULE/FORM: SALES ORDERS POSTING TRANSACTION
' VERSION: VB6

' SALES ORDERS POSTING VARIABLE SETTINGS
Dim strDB As String
Dim db As ADODB.Connection
Dim strSQL As String
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

Private Sub Form_Load()
    ' STARTUP SALES ORDERS POSTING DATABASE CONNECTIONS
    strDB = App.Path + "\DATABASE\MFSC.MDB;Jet OLEDB:Database Password=MFSC;"
    strSQL = "SELECT * FROM SalesOrders_Header ORDER BY SalesOrder_OR_Number"
    Database_Refresh 0
    If adoPrimaryRS.RecordCount <> 0 Then
        adoPrimaryRS.MoveFirst
        Do While Not adoPrimaryRS.EOF
            If adoPrimaryRS("SalesOrder_Posted") = False Then
                txtCombo.AddItem IIf(IsNull(adoPrimaryRS("SalesOrder_OR_Number")), "", adoPrimaryRS("SalesOrder_OR_Number"))
            End If
            adoPrimaryRS.MoveNext
        Loop
    End If
End Sub

Private Sub cmdOk_Click()
    ' POSTING TO INVENTORY STARTS
    adoPrimaryRS.MoveFirst
    Do While adoPrimaryRS.Fields("SalesOrder_OR_Number") <> Trim(txtCombo.Text)
        adoPrimaryRS.MoveNext
    Loop
    mbEditFlag = True
    adoPrimaryRS("SalesOrder_Posted") = True
    adoPrimaryRS.UpdateBatch adAffectAll
    strSQL2 = "SELECT * FROM SalesOrders_Detail WHERE SalesOrder_OR_Number = '" & adoPrimaryRS("SalesOrder_OR_Number") & "'"
    Database_Refresh 1
    Do While Not adoPrimaryRS2.EOF
        strSQL3 = ""
        strSQL3 = "SELECT * FROM Menu_Detail WHERE Menu_Code = '" & adoPrimaryRS2("SalesOrder_Menu_Code") & "'"
        Database_Refresh 2
        Do While Not adoPrimaryRS3.EOF
            strSQL4 = ""
            strSQL4 = "SELECT * FROM Ingredients WHERE Ingredient_Code = '" & adoPrimaryRS3("Menu_Ingredient_Code") & "'"
            Database_Refresh 3
            mbEditFlag = True
            adoPrimaryRS4("Ingredient_QtyOnHand") = (adoPrimaryRS4("Ingredient_QtyOnHand") - (adoPrimaryRS2("SalesOrder_Qty") * adoPrimaryRS3("Menu_QtyConsumed")))
            adoPrimaryRS4.UpdateBatch adAffectAll
            adoPrimaryRS3.MoveNext
        Loop
        adoPrimaryRS2.MoveNext
    Loop
    txtCombo.RemoveItem txtCombo.ListIndex
    MsgBox "Sales Order OR Number " & adoPrimaryRS("SalesOrder_OR_Number") & " is already posted to Inventory!!", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub cmdCancel_Click()
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
