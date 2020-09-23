VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Loaner"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCheckOut 
      Caption         =   "Check Out"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   6360
      Width           =   10455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Titles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   4200
      TabIndex        =   9
      Top             =   3240
      Width           =   6255
      Begin MSComctlLib.ListView lstTitles 
         Height          =   2415
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Double Click to Remove"
         Top             =   600
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Format"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Title"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.ComboBox cmbFormat 
         Height          =   330
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         ToolTipText     =   "Select a Format and press [Enter] to Save"
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Loanee List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   8
      Top             =   3240
      Width           =   4095
      Begin VB.ListBox lstLoanee 
         Height          =   2370
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Double Click to Remove"
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtLoanee 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Press [Enter] to Save"
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Titles Out on Loan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10455
      Begin MSComctlLib.ListView lstTitlesOut 
         Height          =   2895
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Title"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Loanee"
            Object.Width           =   5644
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''
'  Loaner (c) 2002 by IanThurston.com  '
'     If you decide to use this code   '
'     in your own programs, please     '
'     give credit to IanThurston.com   '
'      Thanks!                         '
''''''''''''''''''''''''''''''''''''''''
Dim db As Database
Dim rs As Recordset

'Once we lose focus, check it against the Database to make sure it's not already in there.  If it isn't, we save it.
Private Sub cmbFormat_LostFocus()
Set rs = db.OpenRecordset("SELECT * From [Formats] Where Format='" & cmbFormat.Text & "';")
    If rs.RecordCount = 0 Then
        cmbFormat.AddItem UCase(cmbFormat.Text)
        With rs
            .AddNew
            !Format = UCase(cmbFormat.Text)
            .Update
        End With
        rs.Close
    End If
End Sub

'Add Selected items as a Checkout and add it to the Database
Private Sub cmdCheckOut_Click()
    If Len(lstLoanee.List(lstLoanee.ListIndex)) > 0 = True And Len(lstTitles.SelectedItem.Text) > 0 Then
        If Not CheckListViewTitlesOutDupe(lstTitlesOut, lstTitles.SelectedItem.SubItems(1) & " " & lstTitles.SelectedItem.Text) Then
            lstTitlesOut.Sorted = False
            lstTitlesOut.ListItems.Add , , lstTitles.SelectedItem.SubItems(1) & " " & lstTitles.SelectedItem.Text
            lstTitlesOut.ListItems(lstTitlesOut.ListItems.Count).SubItems(1) = Format(Date, "MM/DD/YYYY")
            lstTitlesOut.ListItems(lstTitlesOut.ListItems.Count).SubItems(2) = lstLoanee.List(lstLoanee.ListIndex)
            lstTitlesOut.Sorted = True
            Set rs = db.OpenRecordset("LoanList")
                With rs
                    .AddNew
                    !Title = lstTitles.SelectedItem.SubItems(1) & " " & lstTitles.SelectedItem.Text
                    !Date = Format(Date, "MM/DD/YYYY")
                    !Loanee = lstLoanee.List(lstLoanee.ListIndex)
                    .Update
                End With
        Else
        MsgBox "[" & lstTitles.SelectedItem.SubItems(1) & "] is already checked out!"
        End If
    End If
End Sub

Private Sub Form_Load()
' If the Database doesn't already exist, make it
If Not FileExists(App.Path & "\Loaner.mdb") Then
    MsgBox "Default Database not Found. Creating Default."
    CreateDB
    Else
    Set db = OpenDatabase(App.Path & "\Loaner.mdb")
    End If

' Load all Lists, Combos, etc
Call Initialize

End Sub
Private Sub Initialize()
Set rs = db.OpenRecordset("LoaneeList")
If rs.RecordCount > 0 Then
    rs.MoveFirst
        Do Until rs.EOF
        lstLoanee.AddItem rs!Loanee
        rs.MoveNext
        Loop
    rs.Close
    End If
    
Set rs = db.OpenRecordset("Formats")
If rs.RecordCount > 0 Then
    rs.MoveFirst
        Do Until rs.EOF
        cmbFormat.AddItem rs!Format
        rs.MoveNext
        Loop
    rs.Close
    End If
    
Set rs = db.OpenRecordset("Titles")
If rs.RecordCount > 0 Then
    rs.MoveFirst
        Do Until rs.EOF
        lstTitles.ListItems.Add , , rs!Format
        lstTitles.ListItems(lstTitles.ListItems.Count).SubItems(1) = rs!Title
        rs.MoveNext
        Loop
    rs.Close
    End If
    
Set rs = db.OpenRecordset("LoanList")
If rs.RecordCount > 0 Then
    rs.MoveFirst
        Do Until rs.EOF
        lstTitlesOut.ListItems.Add , , rs!Title
        lstTitlesOut.ListItems(lstTitlesOut.ListItems.Count).SubItems(1) = rs!Date
        lstTitlesOut.ListItems(lstTitlesOut.ListItems.Count).SubItems(2) = rs!Loanee
        rs.MoveNext
        Loop
    rs.Close
    End If

' Set the sort order for the ListViews
lstTitles.SortKey = 0
lstTitles.SortOrder = lvwAscending
lstTitles.Sorted = True

lstTitlesOut.SortKey = 1
lstTitlesOut.SortOrder = lvwAscending
lstTitlesOut.Sorted = True
    
' If there are any items in cmbFormat, select the first one
If cmbFormat.ListCount > 0 Then
    cmbFormat.ListIndex = 0
    End If
    
End Sub

Private Sub CreateDB()
Dim td As TableDef
Dim fd As Field

Set db = CreateDatabase(App.Path & "\Loaner.mdb", dbLangGeneral, dbEncrypt)
'Create the table
Set td = db.CreateTableDef("Formats")
'Create the field
    Set fd = td.CreateField("Format", dbText)
    td.Fields.Append fd
'Add the Field
    db.TableDefs.Append td

Set td = db.CreateTableDef("Titles")
    Set fd = td.CreateField("Title", dbText)
    td.Fields.Append fd
  
    Set fd = td.CreateField("Format", dbText)
    td.Fields.Append fd
    db.TableDefs.Append td

Set td = db.CreateTableDef("LoanList")
    Set fd = td.CreateField("Title", dbText)
    td.Fields.Append fd
  
    Set fd = td.CreateField("Date", dbDate)
    td.Fields.Append fd

    Set fd = td.CreateField("Loanee", dbText)
    td.Fields.Append fd
    db.TableDefs.Append td

Set td = db.CreateTableDef("LoaneeList")
    Set fd = td.CreateField("Loanee", dbText)
    td.Fields.Append fd
    db.TableDefs.Append td


db.Close

Set db = OpenDatabase(App.Path & "\Loaner.mdb")


End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
rs.Close
db.Close
End Sub

Private Sub lstLoanee_DblClick()
If lstLoanee.ListCount = 0 Then Exit Sub

' Delete the selected entry from the Loanee list when it's doubleclicked
If MsgBox("Delete [" & lstLoanee.List(lstLoanee.ListIndex) & "] from Database?", vbYesNo) = vbYes Then
    Set rs = db.OpenRecordset("SELECT * From [LoaneeList] WHERE Loanee='" & lstLoanee.List(lstLoanee.ListIndex) & "';")
    If rs.RecordCount > 0 Then
        rs.Delete
        lstLoanee.RemoveItem lstLoanee.ListIndex
        Else
        MsgBox "Record not found in Database"
        End If
    rs.Close
    End If
End Sub

Private Sub lstTitles_DblClick()
If lstTitles.ListItems.Count = 0 Then Exit Sub

' Delete the selected entry from the Titles list when it's doubleclicked

If MsgBox("Delete [" & lstTitles.SelectedItem.SubItems(1) & "] from Database?", vbYesNo) = vbYes Then
    Set rs = db.OpenRecordset("SELECT * From [Titles] WHERE Title='" & lstTitles.SelectedItem.SubItems(1) & "';")
    If rs.RecordCount > 0 Then
        With rs
            Do Until .EOF
            If UCase(!Format) = UCase(lstTitles.SelectedItem.Text) Then
                rs.Delete
                cmbFormat = lstTitles.SelectedItem.Text
                txtTitle = lstTitles.SelectedItem.SubItems(1)
                lstTitles.ListItems.Remove lstTitles.SelectedItem.Index
                End If
            .MoveNext
            Loop
        End With
        Else
        MsgBox "Record not found in Database"
        End If
    rs.Close
    End If
End Sub

Private Sub lstTitlesOut_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

lstTitlesOut.Sorted = False
lstTitlesOut.SortKey = ColumnHeader.Index - 1

If lstTitlesOut.SortOrder = lvwAscending Then
    lstTitlesOut.SortOrder = lvwDescending
    Else
    lstTitlesOut.SortOrder = lvwAscending
    End If
    
lstTitlesOut.Sorted = True


End Sub

Private Sub lstTitlesOut_DblClick()

'Delete the selected entry from the Loaned Out list when it's doubleclicked
If lstTitlesOut.ListItems.Count = 0 Then Exit Sub
    Set rs = db.OpenRecordset("SELECT * From [LoanList] WHERE Title='" & lstTitlesOut.SelectedItem.Text & "';")
    If rs.RecordCount > 0 Then
        With rs
            Do Until .EOF
            If UCase(!Loanee) = UCase(lstTitlesOut.SelectedItem.SubItems(2)) Then
                If MsgBox("Delete [" & !Title & "] checked out by [" & !Loanee & "] from Database?", vbYesNo) = vbYes Then
                    rs.Delete
                    lstTitlesOut.ListItems.Remove lstTitlesOut.SelectedItem.Index
                    End If
                End If
            .MoveNext
            Loop
        End With
        Else
        MsgBox "Record not found in Database"
        End If
    rs.Close

End Sub

Private Sub mnuAbout_Click()
MsgBox "Loaner (c) 2002 by IanThurston.com" & vbCrLf & vbCrLf & "If you feel this program or source code is useful to you," & vbCrLf & "please consider sending $10 via PayPal to AlyssaMT99@hotmail.com", vbInformation, "About"
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub txtLoanee_GotFocus()
SelText txtLoanee
End Sub

Private Sub txtLoanee_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(txtLoanee) > 0 Then
        If Not CheckListBoxDupe(lstLoanee, txtLoanee) Then
            Set rs = db.OpenRecordset("LoaneeList")
            With rs
                .AddNew
                !Loanee = txtLoanee
                .Update
            End With
            rs.Close
            lstLoanee.AddItem txtLoanee
            SelText txtLoanee
        Else
        MsgBox "Duplicate Already Exists"
        End If
    End If
End If
    
End Sub

Private Function CheckListBoxDupe(lst As ListBox, strLoanee As String) As Boolean
Dim i As Integer

For i = 0 To lst.ListCount - 1
    If UCase(lst.List(i)) = UCase(strLoanee) Then CheckListBoxDupe = True
    Next i
End Function

Private Function CheckListViewTitlesDupe(lst As ListView, strFormat As String, strTitle As String) As Boolean
Dim i As Integer
    For i = 1 To lst.ListItems.Count
        If UCase(lst.ListItems(i).Text) = UCase(strFormat) And UCase(lst.ListItems(i).SubItems(1)) = UCase(strTitle) Then
            CheckListViewTitlesDupe = True
            lst.ListItems(i).Selected = True
            lst.SetFocus
            End If
        Next i
End Function

Private Function CheckListViewTitlesOutDupe(lst As ListView, strTitle As String) As Boolean
For i = 1 To lst.ListItems.Count
    If UCase(lst.ListItems(i).Text) = UCase(strTitle) Then
        CheckListViewTitlesOutDupe = True
        lst.ListItems(i).Selected = True
        lst.SetFocus
        End If
    Next i
End Function

Private Sub SelText(txtBox As TextBox)
txtBox.SelStart = 0
txtBox.SelLength = Len(txtBox)
End Sub

Private Sub txtTitle_GotFocus()
SelText txtTitle
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(txtTitle) > 0 And Len(cmbFormat) > 0 Then
        If Not CheckListViewTitlesDupe(lstTitles, cmbFormat, txtTitle) Then
            Set rs = db.OpenRecordset("Titles")
            With rs
                .AddNew
                !Title = txtTitle
                !Format = cmbFormat
                .Update
            End With
            rs.Close
            lstTitles.Sorted = False
            lstTitles.ListItems.Add , , cmbFormat
            lstTitles.ListItems(lstTitles.ListItems.Count).SubItems(1) = txtTitle
            lstTitles.Sorted = True
            SelText txtTitle
            Else
            MsgBox "[" & txtTitle & "] Already Exists!"
        End If
    End If
End If
End Sub

