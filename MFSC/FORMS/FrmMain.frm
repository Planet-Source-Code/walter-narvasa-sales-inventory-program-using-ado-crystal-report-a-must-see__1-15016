VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.MDIForm fMain 
   BackColor       =   &H8000000C&
   Caption         =   "Meeting's Fillers, Spirit & Cafe - Sales & Inventory System"
   ClientHeight    =   5730
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5475
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7761
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "2/5/01"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "4:55 AM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":27A2
            Key             =   "Tables"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":44AE
            Key             =   "Crew"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":5302
            Key             =   "Ingredients"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":6156
            Key             =   "Menu"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":7E62
            Key             =   "Suppliers"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":8CB6
            Key             =   "SalesOrders"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":9F3A
            Key             =   "PurchaseOrders"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":AC16
            Key             =   "ReceivingOrders"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1535
      ButtonWidth     =   2514
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Tables"
            Key             =   "Tables"
            ImageKey        =   "Tables"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "k1"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Service Crew"
            Key             =   "Crew"
            ImageKey        =   "Crew"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ingredients"
            Key             =   "Ingredients"
            ImageKey        =   "Ingredients"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Menu"
            Key             =   "Menu"
            ImageKey        =   "Menu"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Suppliers"
            Key             =   "Suppliers"
            ImageKey        =   "Suppliers"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Sales Orders"
            Key             =   "SalesOrders"
            ImageKey        =   "SalesOrders"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Purchase Orders"
            Key             =   "PurchaseOrders"
            ImageKey        =   "PurchaseOrders"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Receiving Orders"
            Key             =   "ReceivingOrders"
            ImageKey        =   "ReceivingOrders"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport PrintIt 
      Left            =   240
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowLeft      =   0
      WindowTop       =   0
      WindowWidth     =   800
      WindowHeight    =   600
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu mnuFileMaintenance 
      Caption         =   "&File Maintenance"
      Begin VB.Menu mnuTables 
         Caption         =   "&Tables.."
      End
      Begin VB.Menu mnuServiceCrew 
         Caption         =   "&Service Crew.."
      End
      Begin VB.Menu mnufBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIngredients 
         Caption         =   "&Ingredients.."
      End
      Begin VB.Menu mnuMenu 
         Caption         =   "&Menu.."
      End
      Begin VB.Menu mnufBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSuppliers 
         Caption         =   "&Suppliers.."
      End
      Begin VB.Menu mnufBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuSalesOrders 
         Caption         =   "&Sales Orders.."
      End
      Begin VB.Menu mnuPurchaseOrders 
         Caption         =   "&Purchase Orders.."
      End
      Begin VB.Menu mnuReceivingOrders 
         Caption         =   "&Receiving Orders.."
      End
      Begin VB.Menu mnutBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPostSalesOrderstoInventory 
         Caption         =   "Post Sales Orders to Inventory.."
      End
      Begin VB.Menu mnuPostReceivingOrderstoInventory 
         Caption         =   "Post Receiving Orders  to Inventory.."
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuInventoryReport 
         Caption         =   "&Inventory Report"
      End
      Begin VB.Menu mnuSalesReport 
         Caption         =   "&Sales Report"
      End
      Begin VB.Menu mnuCriticalReport 
         Caption         =   "&Critical Report"
      End
   End
   Begin VB.Menu mnuUtilities 
      Caption         =   "&Utilities"
      Begin VB.Menu mnuBackup 
         Caption         =   "&Backup.."
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore.."
      End
      Begin VB.Menu mnuuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPasswordSecurity 
         Caption         =   "&Password Security.."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade.."
      End
      Begin VB.Menu mnuTile 
         Caption         =   "Tile.."
      End
      Begin VB.Menu mnuArrangeIcons 
         Caption         =   "Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuProgramGuide 
         Caption         =   "&Program Guide"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About the System"
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MODULE/FORM: MAIN MENU
' VERSION: VB6

Private Sub mnuTables_Click()
    ' LOAD TABLES INFORMATION - FILE MAINTENANCE
    If Rights2_Tables = 1 Then
        fTables.Show
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
End Sub

Private Sub mnuServiceCrew_Click()
    ' LOAD SERVICE CREW INFORMATION - FILE MAINTENANCE
    If Rights2_Service_Crew = 1 Then
        fCrew.Show
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
End Sub

Private Sub mnuIngredients_Click()
    ' LOAD INGREDIENTS INFORMATION - FILE MAINTENANCE
    If Rights2_Ingredients = 1 Then
        fIngredients.Show
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
End Sub

Private Sub mnuMenu_Click()
    ' LOAD MENU INFORMATION - FILE MAINTENANCE
    If Rights2_Menu = 1 Then
        fMenu.Show
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
End Sub

Private Sub mnuSuppliers_Click()
    ' LOAD SUPPLIERS INFORMATION - FILE MAINTENANCE
    If Rights2_Supplier = 1 Then
        fSuppliers.Show
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
End Sub

Private Sub mnuSalesOrders_Click()
    ' LOAD SALES ORDERS - TRANSACTIONS
    If Rights2_SalesOrders = 1 Then
        fSalesOrders.Show
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
End Sub

Private Sub mnuPurchaseOrders_Click()
    ' LOAD PURCHASE ORDERS - TRANSACTIONS
    If Rights2_PurchaseOrders = 1 Then
        fPurchaseOrders.Show
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
End Sub

Private Sub mnuReceivingOrders_Click()
    ' LOAD RECEIVING ORDERS - TRANSACTIONS
    If Rights2_ReceivingOrders = 1 Then
        fReceivingOrders.Show
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
End Sub

Private Sub mnuPostSalesOrderstoInventory_Click()
    ' LOAD POST SALES ORDERS TO INVENTORY
    If Rights2_Post_SalesOrders = 1 Then
        fSalesOrdersPosting.Show
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
End Sub

Private Sub mnuPostReceivingOrderstoInventory_Click()
   ' LOAD POST RECEIVING ORDERS TO INVENTORY
    If Rights2_Post_ReceivingOrders = 1 Then
        fReceivingOrdersPosting.Show
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
End Sub

Private Sub mnuBackup_Click()
    ' LOAD BACKUP
    If Rights3_Backup = 1 Then
        fBackup.Show
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
End Sub

Private Sub mnuRestore_Click()
    ' LOAD RESTORE
    If Rights3_Restore = 1 Then
        fRestore.Show
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
End Sub

Private Sub mnuPasswordSecurity_Click()
    ' LOAD PASSWORD SECURITY
    If Rights3_Password_Security = 1 Then
        fPasswordSecurity.Show
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
End Sub

Private Sub mnuInventoryReport_Click()
    ' PRINT INVENTORY REPORT
    If Rights2_Inventory_Report = 1 Then
        FixPrint
        PrintIt.WindowTitle = "Meeting's Fillers, Spirit & Cafe - Sales & Inventory System - Inventory Report"
        PrintIt.ReportFileName = App.Path & "\REPORTS\INVENTORY.RPT"
        PrintIt.RetrieveDataFiles
        PrintIt.Connect = ";Pwd=" & "MFSC"
        PrintIt.Action = 1
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
End Sub

Private Sub mnuSalesReport_Click()
    ' PRINT SALES REPORT
    If Rights2_Sales_Report = 1 Then
        FixPrint
        PrintIt.WindowTitle = "Meeting's Fillers, Spirit & Cafe - Sales & Inventory System - Sales Report"
        PrintIt.ReportFileName = App.Path & "\REPORTS\SALES.RPT"
        PrintIt.RetrieveDataFiles
        PrintIt.Connect = ";Pwd=" & "MFSC"
        PrintIt.Action = 1
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
End Sub

Private Sub mnuCriticalReport_Click()
    ' PRINT CRITICAL REPORT
    If Rights2_Critical_Report = 1 Then
        FixPrint
        PrintIt.WindowTitle = "Meeting's Fillers, Spirit & Cafe - Sales & Inventory System - Critical Report"
        PrintIt.ReportFileName = App.Path & "\REPORTS\CRITICAL.RPT"
        PrintIt.RetrieveDataFiles
        PrintIt.Connect = ";Pwd=" & "MFSC"
        PrintIt.Action = 1
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
End Sub

Private Sub mnuProgramGuide_Click()
    ' LOAD PROGRAM GUIDE
    fHelp.Show
End Sub

Private Sub mnuAbout_Click()
    ' LOAD ABOUT THE SYSTEM
     MsgBox "Meeting's Fillers, Spirit & Cafe" & vbCrLf & _
            "Sales & Inventory System", vbOKOnly + vbInformation, "About"
End Sub

Private Sub mnuExit_Click()
    ' EXIT PROGRAM
    End
End Sub

Private Sub mnuCascade_Click()
    ' ARRANGE ALL CHILD FORMS TO CASCADE
    fMain.Arrange vbCascade
End Sub

Private Sub mnuTile_Click()
    ' ARRANGE ALL CHILD FORMS TO TILE HORIZONTAL
    fMain.Arrange vbTileHorizontal
End Sub

Private Sub mnuArrangeIcons_Click()
    ' ARRANGE ALL CHILD FORM'S ICONS
    fMain.Arrange vbArrangeIcons
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    ' TOOLBAR FOR MAIN MENU ICON SHORT-CUTS
    On Error Resume Next
    Select Case Button.Key
        Case "Tables"
            If Rights2_Tables = 1 Then
                fTables.Show
                fTables.SetFocus
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "Crew"
            If Rights2_Service_Crew = 1 Then
                fCrew.Show
                fCrew.SetFocus
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "Ingredients"
            If Rights2_Ingredients = 1 Then
                fIngredients.Show
                fIngredients.SetFocus
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "Menu"
            If Rights2_Menu = 1 Then
                fMenu.Show
                fMenu.SetFocus
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "Suppliers"
            If Rights2_Supplier = 1 Then
                fSuppliers.Show
                fSuppliers.SetFocus
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "SalesOrders"
            If Rights2_SalesOrders = 1 Then
                fSalesOrders.Show
                fSalesOrders.SetFocus
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "PurchaseOrders"
            If Rights2_PurchaseOrders = 1 Then
                fPurchaseOrders.Show
                fPurchaseOrders.SetFocus
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "ReceivingOrders"
            If Rights2_ReceivingOrders = 1 Then
                fReceivingOrders.Show
                fReceivingOrders.SetFocus
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
    End Select
End Sub

Private Sub FixPrint()
    Call ReportResolution
    PrintIt.WindowTop = 0
    PrintIt.WindowLeft = 0
    PrintIt.WindowWidth = xWidth
    PrintIt.WindowHeight = xHeight
End Sub
