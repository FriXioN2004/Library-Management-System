Option Explicit

Public CurrentRole As String

#NAVIGATION

Sub GoToSheet(sheetName As String)
    Sheets(sheetName).Activate
End Sub

Sub Logout()
    CurrentRole = ""
    frmLogin.Show
End Sub

#ISSUE BOOK

Sub IssueBook()

    Dim bookName As String
    Dim issueDate As Date
    Dim returnDate As Date
    Dim lastRow As Long

    bookName = Sheets("IssueBook").Range("B2").Value
    issueDate = Sheets("IssueBook").Range("B3").Value

    If bookName = "" Then
        MsgBox "Book Name is mandatory!", vbExclamation
        Exit Sub
    End If

    If issueDate < Date Then
        MsgBox "Issue Date cannot be before today!", vbExclamation
        Exit Sub
    End If

    returnDate = issueDate + 7
    Sheets("IssueBook").Range("B4").Value = returnDate

    lastRow = Sheets("Transactions").Cells(Rows.Count, 1).End(xlUp).Row + 1

    With Sheets("Transactions")
        .Cells(lastRow, 1).Value = bookName
        .Cells(lastRow, 2).Value = issueDate
        .Cells(lastRow, 3).Value = returnDate
        .Cells(lastRow, 4).Value = "Issued"
    End With

    MsgBox "Book Issued Successfully!", vbInformation

End Sub

#RETURN BOOK

Sub ReturnBook()

    Dim dueDate As Date
    Dim returnDate As Date
    Dim fine As Double
    Dim fineRate As Double

    fineRate = 10
    dueDate = Sheets("ReturnBook").Range("B3").Value
    returnDate = Date

    Sheets("ReturnBook").Range("B4").Value = returnDate

    If returnDate > dueDate Then
        fine = (returnDate - dueDate) * fineRate
    Else
        fine = 0
    End If

    Sheets("ReturnBook").Range("B5").Value = fine

    MsgBox "Return Processed. Fine = Rs " & fine, vbInformation

End Sub

#ADD MEMBER

Sub AddMember()

    Dim lastRow As Long

    If Sheets("AddMember").Range("B2").Value = "" Then
        MsgBox "All fields mandatory!", vbExclamation
        Exit Sub
    End If

    lastRow = Sheets("Membership").Cells(Rows.Count, 1).End(xlUp).Row + 1

    With Sheets("Membership")
        .Cells(lastRow, 1).Value = Sheets("AddMember").Range("B2").Value
        .Cells(lastRow, 2).Value = Date
        .Cells(lastRow, 3).Value = DateAdd("m", 6, Date)
    End With

    MsgBox "Membership Added Successfully!", vbInformation

End Sub