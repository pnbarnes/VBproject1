' Name:         pb_Assignment5
' Purpose:      To display and sort members based on a text file.
' Programmer:   Paige B

Public Class frmAssignment5
    Private textFile As String = "Youth_Members.txt"

    Private Sub RefreshListBox()
        ' Clears list box and refreshes it with updated file, sets selected index to none, updates count.
        lstMembers.DataSource = Nothing
        lstMembers.DataSource = IO.File.ReadAllLines(textFile)
        txtCount.Text = lstMembers.Items.Count()
    End Sub

    Private Sub frmAssignment5_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Loads Youth_Members.txt file contents into list box.

        ' Determines whether the file exists.
        If IO.File.Exists("Youth_Members.txt") Then
            ' Calls sub to refresh list box.
            RefreshListBox()
        Else
            MessageBox.Show("Cannot find the file.", "Assignment5",
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
    End Sub

    Private Sub mnuFileExit_Click(sender As Object, e As EventArgs) Handles mnuFileExit.Click
        ' Closes application.
        Me.Close()
    End Sub

    Private Sub mnuSortAscend_Click(sender As Object, e As EventArgs) Handles mnuSortAscend.Click
        ' Sorts list box in ascending order.

        ' Clears listbox.
        lstMembers.DataSource = Nothing

        ' Sorts the list box in ascending order.
        Dim qryAscend = From names In IO.File.ReadAllLines(textFile)
                        Order By names Ascending
                        Select names

        lstMembers.DataSource = qryAscend.ToList()
    End Sub

    Private Sub mnuSortDescend_Click(sender As Object, e As EventArgs) Handles mnuSortDescend.Click
        ' Sorts list box in descending order.

        ' Clears list box.
        lstMembers.DataSource = Nothing

        ' Sorts the list box in ascending order.
        Dim qryDescend = From names In IO.File.ReadAllLines(textFile)
                         Order By names Descending
                         Select names

        lstMembers.DataSource = qryDescend.ToList()
    End Sub

    Private Sub mnuEditAdd_Click(sender As Object, e As EventArgs) Handles mnuEditAdd.Click
        ' Adds a member to the open file based on user input.

        Dim memberName As String = InputBox("Please enter the name of the member", "Assignment 5")

        ' Error check selection statement.
        If memberName = "" Then
            MessageBox.Show("Please enter name.", "Assignment 5",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        Else
            ' Writes new member name to file if input is not null.
            Dim sw As IO.StreamWriter
            sw = IO.File.AppendText(textFile)
            sw.WriteLine(memberName)
            sw.Close()
        End If

        ' Calls sub to refresh list box.
        RefreshListBox()
    End Sub

    Private Sub mnuEditDelete_Click(sender As Object, e As EventArgs) Handles mnuEditDelete.Click
        ' Deletes selected member name from file and refreshes list box.
        If lstMembers.SelectedIndex = -1 Then
            MessageBox.Show("Please select a member to delete", "Assignment 5",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        Else
            Dim deleteMember As String = lstMembers.SelectedItem
            Dim queryDelete = From member In IO.File.ReadAllLines(textFile)
                              Where member <> deleteMember
                              Select member
            IO.File.WriteAllLines(textFile, queryDelete)
            ' Calls sub to refresh list box.
            RefreshListBox()
        End If
    End Sub

    Private Sub mnuEditCreate_Click(sender As Object, e As EventArgs) Handles mnuEditCreate.Click
        ' Creates a new file based on user input.
        Dim newFile As String = InputBox("Please enter a filename to create", "Assignment 5")

        ' Error check.
        If newFile = "" Then
            MessageBox.Show("Please enter a valid filename", "Assignment 5",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        ' Updates textFile to new file name.
        textFile = newFile & ".txt"

        If Not IO.File.Exists(newFile) Then
            Dim sw As IO.StreamWriter = IO.File.CreateText(textFile)
            sw.Close()
            ' Calls sub to refresh list box.
            RefreshListBox()
        Else
            MessageBox.Show("This file already exists", "Assignment 5",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
    End Sub
End Class
