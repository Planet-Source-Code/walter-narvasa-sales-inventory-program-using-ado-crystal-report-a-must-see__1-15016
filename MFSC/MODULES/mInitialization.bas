Attribute VB_Name = "mInitialization"
Global UserName
Global Rights1_Add
Global Rights1_Edit
Global Rights1_Delete
Global Rights2_Tables
Global Rights2_Service_Crew
Global Rights2_Ingredients
Global Rights2_Menu
Global Rights2_Supplier
Global Rights2_SalesOrders
Global Rights2_PurchaseOrders
Global Rights2_ReceivingOrders
Global Rights2_Post_SalesOrders
Global Rights2_Post_ReceivingOrders
Global Rights2_Inventory_Report
Global Rights2_Sales_Report
Global Rights2_Critical_Report
Global Rights3_Backup
Global Rights3_Restore
Global Rights3_Password_Security
Global OrderEntryOpen As Boolean
Global OrderEntryModule As String
Global EditClicked As Boolean
Global xWidth As Integer
Global xHeight As Integer

' REPORT RESOLUTION FIXER
Function ReportResolution()
    If IsResolution(640, 480) Then
        xWidth = 640
        xHeight = 480
    ElseIf IsResolution(800, 600) Then
        xWidth = 800
        xHeight = 600
    ElseIf IsResolution(1024, 768) Then
        xWidth = 1024
        xHeight = 768
    ElseIf IsResolution(1280, 1024) Then
        xWidth = 1280
        xHeight = 1024
    ElseIf IsResolution(1600, 1200) Then
        xWidth = 1600
        xHeight = 1200
    End If
End Function

' FOR RESOLUTION VERIFIER
Function IsResolution(Width As Integer, Height As Integer) As Boolean
    If (Screen.Width / Screen.TwipsPerPixelX = Width) And (Screen.Height / Screen.TwipsPerPixelY = Height) Then
        IsResolution = True
    Else
        IsResolution = False
    End If
End Function

' DECODE PASSWORD.
Function Decode_Pass(p_str As String) As String
    For i = 1 To Len(p_str) Step 1
        strs = strs + Chr(Asc(Mid(p_str, i, 1)) * 2)
    Next i
        Decode_Pass = strs
End Function

' UNCODE PASSWORD.
Function UnCode_Pass(p_str As String) As String
    For i = 1 To Len(p_str) Step 1
        strs = strs + Chr(Asc(Mid(p_str, i, 1)) / 2)
    Next i
        UnCode_Pass = strs
End Function

