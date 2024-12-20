VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOrderRequest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Order"
   ClientHeight    =   10305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10305
   ScaleWidth      =   7875
   Begin VB.TextBox txtSubTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   9000
      Width           =   2175
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   9000
      Width           =   2175
   End
   Begin VB.TextBox txtTotalTax 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   8640
      Width           =   2175
   End
   Begin VB.TextBox txtFreightCharge 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1320
      TabIndex        =   8
      Top             =   8640
      Width           =   2175
   End
   Begin VB.TextBox txtSalesTax 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1320
      TabIndex        =   7
      Top             =   8280
      Width           =   2175
   End
   Begin VB.TextBox txtEntry 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   25
      Top             =   7920
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid fgProducts 
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BorderStyle     =   0
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   9930
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13838
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtRequired 
      Height          =   300
      Left            =   1800
      TabIndex        =   4
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16449537
      CurrentDate     =   41323
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   9480
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddProducts 
      Caption         =   "..."
      Height          =   315
      Left            =   7320
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   9480
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search customer"
      Height          =   3375
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtContactLastName 
         Height          =   300
         Left            =   5040
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtContactName 
         Height          =   300
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton cmdCustomers 
         Caption         =   "..."
         Height          =   315
         Left            =   6840
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCompanyName 
         Height          =   300
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin MSComctlLib.ListView lvCustomers 
         Height          =   1935
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Customer ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Company Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Contact Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Contact Last Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "City"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "State"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Country"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Contact last name:"
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Company name:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Contact name:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Customer"
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   7575
      Begin VB.TextBox txtCustomerContact 
         BackColor       =   &H80000004&
         Height          =   300
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtCustomerCompany 
         BackColor       =   &H80000004&
         Height          =   300
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Company:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Contact:"
         Height          =   255
         Left            =   3480
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
   End   
   Begin MSComCtl2.DTPicker dtPromised 
      Height          =   300
      Left            =   5280
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16449537
      CurrentDate     =   41323
   End
   Begin VB.Label Label13 
      Caption         =   "Line quantity:"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Freight Charge:"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Total:"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   9000
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Total Tax:"
      Height          =   255
      Left            =   4080
      TabIndex        =   29
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Sub Total:"
      Height          =   255
      Left            =   4080
      TabIndex        =   27
      Top             =   9000
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Sales Tax:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Promised by date:"
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Required by date:"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   4560
      Width           =   1575
   End
End
Attribute VB_Name = "frmOrderRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private currentCompanyName As String
Private currentIdCustomer As Integer
Private currentContactName As String
Private editingData As Boolean

Private currentSubTotal As Double
Private currentTotal As Double
Private currentTax As Double
Private currentFreightCharge As Double
Private currentTotalTax As Double
Private editingQuantity As Boolean


Private Sub cmdAddProducts_Click()
frmAddProductTo.Id = currentIdCustomer
frmAddProductTo.ObjectReferred = "Customer " & txtCustomerCompany & "|" & txtCustomerContact
frmAddProductTo.Table = "ProductsByCustomer"
frmAddProductTo.ColumnName = "CustomerId"
frmAddProductTo.LoadData
frmAddProductTo.Show vbModal
If frmAddProductTo.SavedChanges Then
    LoadProductsById
End If
End Sub

Private Sub txtName_Change()
DoSearchCustomer
End Sub

Private Sub DoSearchCustomer(Optional Id As String)
Dim filter As String
filter = ""
'If Not IsEmpty(Id) Then
If Id <> "" Then
    filter = "CustomerID = " & Id
End If
If txtCompanyName <> Empty Then
    If filter <> Empty Then
        filter = filter & " AND "
    End If
    filter = "CompanyName LIKE '%" & txtCompanyName & "%'"
End If
If txtContactName <> Empty Then
    If filter <> Empty Then
        filter = filter & " AND "
    End If
    filter = filter & "ContactFirstName LIKE '%" & txtContactName & "%'"
End If
If txtContactLastName <> Empty Then
    If filter <> Empty Then
        filter = filter & " AND "
    End If
    filter = filter & "ContactLastName LIKE '%" & txtContactLastName & "%'"
End If

If filter <> Empty Then
    filter = "Where " & filter
End If
ExecuteSql "Select CustomerID, CompanyName, ContactFirstName, ContactLastName, City, StateOrProvince, 'Country/Region' From Customers " & filter
lvCustomers.ListItems.Clear
If rs.RecordCount = 0 Then
    LogStatus "There are no records with the selected criteria", Me
Else
    Dim x As ListItem
    While Not rs.EOF
        Set x = lvCustomers.ListItems.Add(, , rs.Fields(0))
        For i = 1 To (rs.Fields.Count - 1)
            If Not IsEmpty(rs.Fields(i)) Then
                x.SubItems(i) = rs.Fields(i)
            End If
        Next i
        rs.MoveNext
    Wend
    If lvCustomers.ListItems.Count = 1 Then
        lvCustomers.SelectedItem = lvCustomers.ListItems(1)
    End If
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCustomers_Click()
'On Error Resume Next
frmCustomers.Show vbModal
txtCompanyName = ""
txtContactLastName = ""
txtContactName = ""
DoSearchCustomer frmCustomers.CurrentCustomerID
Unload frmCustomers
End Sub

Private Sub cmdSave_Click()
Dim newOrderId As Integer

On Error GoTo HandleError
ExecuteSql "Select * from OrderRequests"
rs.AddNew
rs("CustomerId") = currentIdCustomer
rs("EmployeeId") = UserId
rs("OrderDate") = CStr(Date)
rs("RequiredByDate") = CStr(dtRequired.value)
rs("PromisedByDate") = CStr(dtPromised.value)
rs("FreightCharge") = currentFreightCharge
rs("SalesTaxRate") = currentTax * 0.01
rs("Status") = "REQUESTED"
rs.Update
newOrderId = rs("OrderID")


For i = 1 To fgProducts.Rows - 1
    If fgProducts.TextMatrix(i, 0) <> "0" Then
        ExecuteSql "Insert into OrderRequestDetails (OrderID, ProductID, DateSold, Quantity, UnitPrice, SalePrice, LineTotal) Values (" _
        & newOrderId & ", '" & fgProducts.TextMatrix(i, 1) & "', '" & Format(Date, "mm/dd/yyyy") & "'," & fgProducts.TextMatrix(i, 0) & "," & fgProducts.TextMatrix(i, 3) & "," & fgProducts.TextMatrix(i, 4) & "," & fgProducts.TextMatrix(i, 4) & ")"
    
        ExecuteSql "Update Products Set UnitsOnOrder = UnitsOnOrder + " & fgProducts.TextMatrix(i, 0) & _
        " Where ProductId = '" & fgProducts.TextMatrix(i, 1) & "'"
    
    End If
Next i

editingData = False
If MsgBox("Order added successfully" & vbCrLf & "Would you like to add a new order?", vbYesNo + vbQuestion, "New data") = vbYes Then
    ClearFields
Else
    Unload Me
End If
Exit Sub
HandleError:
MsgBox "An error has occurred adding the data. Error: (" & err.Number & ") " & err.Description, vbCritical, "Error"
End Sub

Private Sub dtPromised_Change()
editingData = True
End Sub

Private Sub dtRequired_Change()
'If dtPromised.value < dtRequired.value Then
'    dtPromised.value = dtRequired.value
'End If
editingData = True
End Sub

Private Sub MakeTextBoxVisible(txtBox As textbox, grid As MSFlexGrid)
With grid
    If .Row < 0 Or .col < 0 Then Exit Sub
    txtBox.text = .TextMatrix(.Row, .col)
    txtBox.Enabled = True
    
    txtBox.SetFocus
    DoEvents
    editingQuantity = True
End With
End Sub

Private Sub fgProducts_Click()
If fgProducts.col <> 0 Then Exit Sub
MakeTextBoxVisible txtEntry, fgProducts
SelectAll txtEntry
End Sub

Private Sub fgProducts_EnterCell()
SaveEdits
End Sub

Private Sub fgProducts_KeyPress(KeyAscii As Integer)
If fgProducts.col <> 0 Then Exit Sub
Select Case KeyAscii
Case 46, 48 To 57
'Case 45, 46, 47, 48 To 59, 65 To 90, 97 To 122
    MakeTextBoxVisible txtEntry, fgProducts
    txtEntry.text = Chr$(KeyAscii)
    txtEntry.SelStart = 1
End Select
End Sub

Private Sub EditKeyCode(grid As MSFlexGrid, txtBox As textbox, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 27 'ESC
    txtBox = ""
    txtBox.Visible = False
    grid.SetFocus
Case 13 'Return
    grid.SetFocus
Case 37 'Left Arrow
    grid.SetFocus
    DoEvents
    If grid.col > grid.FixedCols Then
        grid.col = grid.col - 1
    End If
Case 38 'Up Arrow
    grid.SetFocus
    DoEvents
    If grid.Row > grid.FixedRows Then
        grid.Row = grid.Row - 1
    End If
Case 39 'Right Arrow
    grid.SetFocus
    DoEvents
    If grid.col < grid.Cols - 1 Then
        grid.col = grid.col + 1
    End If
Case 40 'Down Arrow
    grid.SetFocus
    DoEvents
    If grid.Row < grid.Rows - 1 Then
        grid.Row = grid.Row + 1
    End If
End Select
End Sub


Private Sub txtEntry_LostFocus()
SaveEdits
End Sub


Private Sub fgProducts_LeaveCell()
    If (editingQuantity) Then
        SaveEdits
    End If
End Sub

Private Sub SaveEdits()
Dim lineQuantity As Double, lineUnitPrice As Double, linePrice As Double
Dim previousLinePrice As Double
If Not editingQuantity Or _
   Not ValidateTextBoxDouble(txtEntry, Me) Or _
   Not ValidateTextDouble(fgProducts.TextMatrix(fgProducts.Row, 4), Me) Then
   Exit Sub
End If
previousLinePrice = DoubleValue(fgProducts.TextMatrix(fgProducts.Row, 4))
fgProducts.TextMatrix(fgProducts.Row, fgProducts.col) = txtEntry.text
lineQuantity = DoubleValue(txtEntry.text)
lineUnitPrice = DoubleValue(fgProducts.TextMatrix(fgProducts.Row, 3))
previousLinePrice = DoubleValue(fgProducts.TextMatrix(fgProducts.Row, 4))
linePrice = CDbl(lineQuantity * lineUnitPrice)
fgProducts.TextMatrix(fgProducts.Row, 4) = linePrice
ReCalculateTotals previousLinePrice, linePrice
editingQuantity = False
txtEntry.Enabled = False
txtEntry.text = ""

editingData = True
End Sub

Private Sub ReCalculateTotals(previous As Double, current As Double)
currentSubTotal = currentSubTotal - previous + current
currentTotalTax = currentSubTotal * currentTax * 0.01
currentTotal = currentFreightCharge + currentSubTotal + currentTotalTax
txtSubTotal = Format(currentSubTotal, "#,##0.00")
txtTotalTax = Format(currentTotalTax, "#,##0.00")
txtTotal = Format(currentTotal, "#,##0.00")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If editingData Then
    Dim res As VbMsgBoxResult
    res = MsgBox("Do you want to save the edited data?", vbYesNoCancel + vbQuestion, "Save data")
    If res = vbYes Then
        cmdSave_Click
    ElseIf res <> vbNo Then
        Cancel = True
    End If
End If
End Sub

Private Sub Form_Load()
editingData = False
ClearFields
End Sub

Private Sub lvCustomers_ItemClick(ByVal Item As MSComctlLib.ListItem)
RetrieveDataCustomer
End Sub

Private Sub RetrieveDataCustomer()
If editingData Then
    If MsgBox("Do you want to cancel previous edited data?", vbYesNo + vbQuestion, "Data edition") <> vbYes Then
        Exit Sub
    End If
End If

If lvCustomers.SelectedItem <> Empty Then
    With lvCustomers.SelectedItem
        currentIdCustomer = lvCustomers.SelectedItem
        currentCompanyName = .SubItems(1)
        currentContactName = .SubItems(2) & " " & .SubItems(3)
    End With
    txtCustomerCompany = currentCompanyName
    txtCustomerContact = currentContactName
    editingData = False
End If
LoadProductsById
cmdSave.Enabled = True
cmdAddProducts.Enabled = True
dtRequired.Enabled = True
dtPromised.Enabled = True

End Sub

Private Sub LoadProductsById()
Dim Table As String
Dim ColumnName As String
Dim Id As Integer
Table = "ProductsByCustomer"
ColumnName = "CustomerId"
Id = currentIdCustomer

ExecuteSql "Select p.ProductID, p.ProductName, p.UnitPrice, p.UnitsInStock, p.UnitsOnOrder, p.QuantityPerUnit, p.Unit from Products as p, " & Table _
& " as pb Where pb." & ColumnName & " = " & Id & " And pb.ProductId = p.ProductId"

'lvProducts.ListItems.Clear
'If rs.RecordCount > 0 Then
'    With rs
'        While Not .EOF
'            Set x = lvProducts.ListItems.Add(, , 0)
'            For i = 1 To 5
'                If Not IsEmpty(.Fields(i - 1)) Then
'                    x.SubItems(i) = .Fields(i - 1)
'                End If
'            Next i
'            x.SubItems(6) = .Fields(5) & .Fields(6)
'            .MoveNext
'        Wend
'    End With
'End If

Dim lng As Long
Dim intLoopCount As Integer
Const SCROOL_WIDTH = 320
With fgProducts
    .Cols = 8
    .FixedCols = 0
    .Rows = 0
    .AddItem "Quantity" & vbTab & "Code" & vbTab & "Product" & vbTab & "UnitPrice" & vbTab & "Price" & vbTab & "Existence" & vbTab & "Ordered" & vbTab & "Quantity per unit"
    .Rows = rs.RecordCount + 1
    If .Rows = 1 Then .FixedRows = 0 Else .FixedRows = 1
    Dim i As Integer
    Dim j As Integer
    i = 1
    While Not rs.EOF
        .TextMatrix(i, 0) = "0"
        For j = 1 To 6
            If j = 4 Then
                .TextMatrix(i, j) = "0"
            ElseIf j < 4 Then
                .TextMatrix(i, j) = rs.Fields(j - 1)
            Else
                .TextMatrix(i, j) = rs.Fields(j - 2)
            End If
        Next j
        .TextMatrix(i, 7) = rs.Fields(5) & rs.Fields(6)
        rs.MoveNext
        i = i + 1
    Wend
End With
End Sub

Private Sub lvProducts_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Item.Checked Then
    Item.text = "1"
Else
    Item.text = "0"
End If
End Sub

Private Sub txtCompanyName_Change()
DoSearchCustomer
End Sub

Private Sub txtContactLastName_Change()
DoSearchCustomer
End Sub

Private Sub txtContactName_Change()
DoSearchCustomer
End Sub

Private Sub ClearFields()

fgProducts.Rows = 0
fgProducts.Cols = 0

currentSubTotal = 0
currentTotal = 0
currentTax = 0
currentTotalTax = 0
currentFreightCharge = 0

txtSubTotal = ""
txtTotal = ""
txtTotalTax = ""
txtSalesTax = ""
txtFreightCharge = ""

txtCompanyName = ""
txtContactLastName = ""
txtContactName = ""
txtCustomerContact = ""
txtCustomerCompany = ""
cmdSave.Enabled = False
cmdAddProducts.Enabled = False
dtRequired.Enabled = False
dtPromised.Enabled = False
'txtCompanyName.SetFocus
'txtCompanyName.SetFocus
ReCalculateTotals 0, 0
editingData = False
End Sub

Private Sub txtFreightCharge_Change()
currentFreightCharge = DoubleValue(txtFreightCharge.text)
ReCalculateTotals 0, 0
editingData = True
End Sub

Private Sub txtFreightCharge_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case vbKey0 To vbKey9
    Case vbKeyBack, vbKeyClear, vbKeyDelete
    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
    Case Else
        KeyAscii = 0
        Beep
End Select
End Sub

Private Sub txtSalesTax_Change()
currentTax = DoubleValue(txtSalesTax.text)
ReCalculateTotals 0, 0
editingData = True
End Sub

Private Sub txtSalesTax_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case vbKey0 To vbKey9
    Case vbKeyBack, vbKeyClear, vbKeyDelete
    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
    Case Else
        KeyAscii = 0
        Beep
End Select
End Sub
