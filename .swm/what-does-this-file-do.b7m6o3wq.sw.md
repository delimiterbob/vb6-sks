---
title: What does this file do?
---
# Introduction

This document will walk you through the key functions and design decisions in the <SwmPath>[modFunctions.bas](/modFunctions.bas)</SwmPath> file. This module contains utility functions that support various operations in the application.

We will cover:

1. How collections are managed and checked for duplicates.
2. How string values are converted and validated.
3. How user interface elements like textboxes and comboboxes are handled.
4. How records are managed and identified in the database.

# Managing collections

<SwmSnippet path="/modFunctions.bas" line="10">

---

The <SwmToken path="/modFunctions.bas" pos="10:4:4" line-data="Public Function AddToCollection(col As Collection, Item As String) As Boolean">`AddToCollection`</SwmToken> function ensures that items are only added to a collection if they do not already exist. This prevents duplicates and maintains data integrity.

```
Public Function AddToCollection(col As Collection, Item As String) As Boolean
AddToCollection = False
If Not Exists(col, Item) Then
    col.Add Item, Item
    AddToCollection = True
End If
End Function
```

---

</SwmSnippet>

# Checking existence in collections

<SwmSnippet path="/modFunctions.bas" line="18">

---

The <SwmToken path="/modFunctions.bas" pos="18:4:4" line-data="Public Function Exists(col As Collection, Index As String) As Boolean">`Exists`</SwmToken> function checks if an item is present in a collection. It uses error handling to manage cases where the item is not found.

```
Public Function Exists(col As Collection, Index As String) As Boolean
Dim o As Variant
On Error GoTo Error
    o = col(Index)
Error:
   Exists = o <> Empty
End Function
```

---

</SwmSnippet>

# Converting and validating string values

<SwmSnippet path="/modFunctions.bas" line="27">

---

The <SwmToken path="/modFunctions.bas" pos="27:4:4" line-data="Public Function DoubleValue(strValue As String)">`DoubleValue`</SwmToken> function converts a string to a double. It returns 0 if the string is empty, ensuring that operations expecting a numeric value do not fail.

```
Public Function DoubleValue(strValue As String)
If Len(strValue) <> 0 Then
    DoubleValue = CDbl(strValue)
Else
    DoubleValue = 0
End If
End Function
```

---

</SwmSnippet>

# Validating textbox input

<SwmSnippet path="/modFunctions.bas" line="35">

---

The <SwmToken path="/modFunctions.bas" pos="35:4:4" line-data="Public Function ValidateTextBoxDouble(txBox As textbox, parentForm As Form)">`ValidateTextBoxDouble`</SwmToken> function checks if the input in a textbox can be converted to a double. If not, it logs an error and resets the textbox, ensuring only valid numeric input is processed.

```
Public Function ValidateTextBoxDouble(txBox As textbox, parentForm As Form)
On Error GoTo err:
   DoubleValue txBox.text
   ValidateTextBoxDouble = True
   Exit Function
err:
   modMain.LogStatus "The value inserted is not valid", parentForm
   txBox.text = ""
   txBox.SetFocus
   ValidateTextBoxDouble = False
End Function
```

---

</SwmSnippet>

# Handling combobox operations

<SwmSnippet path="/modFunctions.bas" line="67">

---

The <SwmToken path="/modFunctions.bas" pos="71:4:4" line-data="Public Sub LoadCombo(Table As String, combo As ComboBox, _">`LoadCombo`</SwmToken> function populates a combobox with data from a specified table. It supports optional value fields, allowing for more complex data structures.

```
''''''''''''''''''''''''''''''''''
''' Combobox related functions '''
''''''''''''''''''''''''''''''''''

Public Sub LoadCombo(Table As String, combo As ComboBox, _
                    field As String, Optional valueField As String)
ExecuteSql "Select * From " & Table
combo.Clear
If (valueField <> Empty) Then
    While Not rs.EOF
        combo.AddItem (rs.Fields(field))
        combo.ItemData(combo.NewIndex) = rs.Fields(valueField)
        rs.MoveNext
    Wend
Else
    While Not rs.EOF
        combo.AddItem (rs.Fields(field))
        rs.MoveNext
    Wend
End If
'If strDefault <> Empty Then
   ' combo = strDefault
'End If
End Sub
```

---

</SwmSnippet>

# Checking if combobox is empty

<SwmSnippet path="/modFunctions.bas" line="93">

---

The <SwmToken path="/modFunctions.bas" pos="93:4:4" line-data="Public Function ComboEmpty(ByRef combo As ComboBox, _">`ComboEmpty`</SwmToken> function verifies if a combobox has a selected item. It prompts the user if no selection is made, ensuring that a valid choice is required before proceeding.

```
Public Function ComboEmpty(ByRef combo As ComboBox, _
                Optional strip As Variant, _
                Optional Index As Integer) _
                As Boolean
If combo.ListIndex = -1 Then
    ComboEmpty = True
    MsgBox "Please select an option from the list", vbExclamation
    If Index <> Empty Then
        'strip.SelectedItem = strip.Tabs(Index)
    End If
    combo.SetFocus
Else
    ComboEmpty = False
End If
End Function
```

---

</SwmSnippet>

# Managing record identification

<SwmSnippet path="/modFunctions.bas" line="120">

---

The <SwmToken path="/modFunctions.bas" pos="120:4:4" line-data="Public Function RcrdId(Table As String, Optional Identifier As String, Optional FldNo As String) As String">`RcrdId`</SwmToken> function generates a unique record identifier based on the last record number in a table. It appends an optional identifier and the current month, ensuring unique and traceable record IDs.

```
Public Function RcrdId(Table As String, Optional Identifier As String, Optional FldNo As String) As String
Dim RcrdNo As Integer
ExecuteSql "Select * from " & Table & " order by " & FldNo & " ASC"
If rs.EOF = False Then
    rs.MoveLast
    RcrdNo = rs.Fields(FldNo) + 1
Else
    RcrdNo = 1
End If
If Identifier <> Empty Then
    RcrdId = Identifier & RcrdNo & Format(Date, "mm")
Else
    RcrdId = RcrdNo
End If
End Function
```

---

</SwmSnippet>

# Searching and displaying results

<SwmSnippet path="/modFunctions.bas" line="138">

---

The <SwmToken path="/modFunctions.bas" pos="139:4:4" line-data="Public Sub SearchShow(Table As String, fieldToSearch As String, itemToSearch As String)">`SearchShow`</SwmToken> function initiates a search operation and displays the results in a modal form. This encapsulates the search logic and user interface interaction.

```
'''''''''''''''''''''''''''''''''''''''''
Public Sub SearchShow(Table As String, fieldToSearch As String, itemToSearch As String)
With frmSearch
    .Search Table, fieldToSearch, itemToSearch
    .Show vbModal
End With
End Sub
```

---

</SwmSnippet>

# Validating textbox content

<SwmSnippet path="/modFunctions.bas" line="166">

---

The <SwmToken path="/modFunctions.bas" pos="166:4:4" line-data="Public Function TextBoxEmpty(ByRef stext As textbox, Optional TabObject As Variant, Optional TabIndex As Integer) As Boolean">`TextBoxEmpty`</SwmToken> function checks if a textbox is empty or contains a placeholder date. It prompts the user to fill in required fields, ensuring data completeness.

```
Public Function TextBoxEmpty(ByRef stext As textbox, Optional TabObject As Variant, Optional TabIndex As Integer) As Boolean
If Trim(stext) = Empty Or stext.text = "  /  /    " Then
    TextBoxEmpty = True
    MsgBox "You need to fill in all required fields", vbExclamation
    If TabIndex <> Empty Then
        'TabObject.SelectedItem = TabObject.Tabs(TabIndex)
    End If
    stext.SetFocus
Else
    TextBoxEmpty = False
End If
End Function
```

---

</SwmSnippet>

# Ensuring numeric input in textboxes

<SwmSnippet path="/modFunctions.bas" line="179">

---

The <SwmToken path="/modFunctions.bas" pos="179:4:4" line-data="Public Function TextBoxNumberEmpty(ByRef textbox As textbox) As Boolean">`TextBoxNumberEmpty`</SwmToken> function checks if a textbox contains a numeric value. It prompts the user if the input is not numeric, ensuring that numeric fields are correctly filled.

```
Public Function TextBoxNumberEmpty(ByRef textbox As textbox) As Boolean
'if the input is not a numeric then true
If IsNumeric(textbox.text) = False Then
    TextBoxNumberEmpty = True
    MsgBox "The field requires a numeric value.", vbExclamation
    textbox.SetFocus
    SelectAll textbox
Else
    TextBoxNumberEmpty = False
End If
End Function
```

---

</SwmSnippet>

# Saving records to the database

<SwmSnippet path="/modFunctions.bas" line="193">

---

The <SwmToken path="/modFunctions.bas" pos="193:4:4" line-data="Private Sub SaveDetection(Reference As String, Title As String, Description As String, Table As String)">`SaveDetection`</SwmToken> function inserts a new record into a specified table. It uses the <SwmToken path="/modFunctions.bas" pos="120:4:4" line-data="Public Function RcrdId(Table As String, Optional Identifier As String, Optional FldNo As String) As String">`RcrdId`</SwmToken> function to generate a unique record number, ensuring that each entry is distinct and properly indexed.

```
Private Sub SaveDetection(Reference As String, Title As String, Description As String, Table As String)
ExecuteSql2 "Select * from " & Table
rs2.AddNew
rs2.Fields!record_no = Val(RcrdId(Table, , "record_no"))
rs2.Fields!Reference = Reference
rs2.Fields!war_type = Title
rs2.Fields!Description = Description
rs2.Update
End Sub
```

---

</SwmSnippet>

This walkthrough highlights the key functionalities and design choices in <SwmPath>[modFunctions.bas](/modFunctions.bas)</SwmPath>, focusing on data integrity, user input validation, and efficient data management.

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBdmI2LXNrcyUzQSUzQWRlbGltaXRlcmJvYg==" repo-name="vb6-sks"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
