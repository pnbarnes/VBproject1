'   Name:       pb_assignment4
'   Purpose:    To calculate and display donations.
'   Programmer: Paige B
Public Class frmDonations
    ' Declares class-level array of 5 donation entry points.
    Private strDonations() As String = {"Main", "Food Bank",
            "Local Missions", "Global Missions", "Planter Church"}

    Private Sub frmDonations_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Clears Donations listbox before it is written to.
        lstDonations.Items.Clear()
        ' Loads the Donations array into listbox.
        For i As Integer = 0 To strDonations.Count - 1
            lstDonations.Items.Add(strDonations(i))
        Next
        lstDonations.SelectedIndex = 0

        ' Displays the current date for donations and programmer's name in Totals listbox.
        lstTotals.Items.Add("Donations Received for " & Today())
        lstTotals.Items.Add("Donations Accepted by Paige B" & ControlChars.NewLine)
        lstTotals.Items.Add("____________________________________" & ControlChars.NewLine)

    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        ' Static array variable for accumulating donation amounts.
        Static decTotals(4) As Decimal

        ' Variables for user input.
        Dim strFullName As String
        Dim decAmount As Double
        Dim strDonation As String
        Dim blnAmountOk As Boolean

        ' Variables declared, error checking.
        strFullName = txtName.Text
        blnAmountOk = Decimal.TryParse(txtAmount.Text, decAmount)
        strDonation = lstDonations.SelectedItem

        ' Error check for Name input
        Dim strProperName
        If strFullName Like "*, *" Then
            strProperName = ProperOrder(strFullName)

        Else
            strProperName = String.Empty
            MessageBox.Show("Please enter name as lastname, firstname", "Donations",
                 MessageBoxButtons.OK, MessageBoxIcon.Information)
            txtName.Focus()
            Return
        End If

        ' Error check for Amount input
        If blnAmountOk = False Then
            MessageBox.Show("Please enter a valid amount.", "Donations",
                MessageBoxButtons.OK, MessageBoxIcon.Information)
            txtAmount.Focus()
            Return
        ElseIf lstDonations.SelectedIndex = -1 Then
            MessageBox.Show("Please select a donation area.", "Donations",
                MessageBoxButtons.OK, MessageBoxIcon.Information)
            lstDonations.Focus()
            Return
        Else
            ' Adds user's amount to the running total of donations for the selected donation area.
            decTotals(lstDonations.SelectedIndex) = decTotals(lstDonations.SelectedIndex) + decAmount
        End If

        ' Displays user's name, donation amount, and donation area.

        lstTotals.Items.Add(strProperName & " - " & "Amount Donated " & decAmount.ToString("C2") & ControlChars.NewLine)
        lstTotals.Items.Add(" ")
        txtAmount.Text = String.Empty
        txtName.Text = String.Empty
        lstDonations.SelectedIndex = -1

        ' Running totals displayed in labels.
        lblMain.Text = decTotals(0).ToString("C2")
        lblBank.Text = decTotals(1).ToString("C2")
        lblLocal.Text = decTotals(2).ToString("C2")
        lblGlobal.Text = decTotals(3).ToString("C2")
        lblChurch.Text = decTotals(4).ToString("C2")

    End Sub

    Private Function ProperOrder(ByVal FullName) As String
        ' Puts name in the order of first space last.
        Dim orderedName As String
        Dim intIndex As Integer = FullName.IndexOf(",")
        Dim strLastName As String = FullName.Substring(0, intIndex)
        Dim strFirstName As String = FullName.Substring(intIndex + 2)
        orderedName = (strFirstName & " " & strLastName)
        Return orderedName
    End Function

    Private Sub radAscending_CheckedChanged(sender As Object, e As EventArgs) Handles radAscending.CheckedChanged
        ' Sorts the Donations listbox in ascending order.
        lstDonations.Items.Clear()
        Array.Sort(strDonations)
        For i As Integer = 0 To strDonations.Count - 1
            lstDonations.Items.Add(strDonations(i))
        Next
        lstDonations.SelectedIndex = 0
    End Sub

    Private Sub radDescending_CheckedChanged(sender As Object, e As EventArgs) Handles radDescending.CheckedChanged
        ' Sorts the Donations listbox in descending order.
        lstDonations.Items.Clear()
        Array.Sort(strDonations)
        Array.Reverse(strDonations)
        For i As Integer = 0 To strDonations.Count - 1
            lstDonations.Items.Add(strDonations(i))
        Next
        lstDonations.SelectedIndex = 0
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        ' Closes the application when the Close button is clicked.
        Me.Close()
    End Sub

End Class
