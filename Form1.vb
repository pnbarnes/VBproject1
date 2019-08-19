'   Program Name:       Youth Camp Calculator
'   Program Purpose:    To calculate and display the cost of youth camp.
'   Programmer Name:    Paige Barnes

Public Class frmYouthCamp

    Private Const FEE As Double = 300
    Private Function getDailyCost(ByVal location As String) As Double
        ' Function to get daily cost based on location selected by user.
        Dim dblDailyCost As Double
        Select Case location
            Case "New York"
                dblDailyCost = 500
            Case "Munising"
                dblDailyCost = 100
            Case "Au Train"
                dblDailyCost = 50
            Case "Hocking Hills"
                dblDailyCost = 150
        End Select
        Return dblDailyCost
    End Function

    Private Sub BusNeed(ByVal CamperNumber)
        ' Sub to determine bus need based on number of campers.
        If CamperNumber > 10 Then
            lblYN.Text = "Y"
        Else
            lblYN.Text = "N"
        End If
    End Sub

    Private Sub btnCalc_Click(sender As Object, e As EventArgs) Handles btnCalc.Click
        ' Variables Defined.
        Dim dblInsurance As Double
        Dim intDays As Integer
        Dim intCampers As Integer
        Dim strName As String
        Dim strLocation As String
        Dim dblCost As Double
        Dim dblTotal As Double
        Dim blnDaysOk As Boolean
        Dim blnCampersOk As Boolean

        ' Error checks.
        If txtName.Text = "" Then
            MessageBox.Show("Please enter group name")
            txtName.Focus()
            Return
        End If
        If txtCampers.Text = "" Then
            MessageBox.Show("Please enter an amount")
            txtCampers.Focus()
            Return
        End If
        If txtDays.Text = "" Then
            MessageBox.Show("Please enter an amount")
            txtDays.Focus()
            Return
        End If
        If cboLocation.SelectedIndex = -1 Then
            MessageBox.Show("Please select a location")
            Return
        End If

        ' Variables declared.
        If chkInsurance.Checked Then
            dblInsurance = 150
        Else
            dblInsurance = 0
        End If

        strName = txtName.Text

        ' Converts user input to integer, stores success/failure in a variable.
        blnDaysOk = Integer.TryParse(txtDays.Text, intDays)
        blnCampersOk = Integer.TryParse(txtCampers.Text, intCampers)

        ' Gets daily cost based on the selected location.
        strLocation = cboLocation.SelectedItem
        dblCost = getDailyCost(strLocation)


        ' Display daily cost.
        lstTotal.Items.Clear()
        lstTotal.Items.Add(strName & " will be travelling to " & strLocation & ControlChars.NewLine)
        lstTotal.Items.Add("They will be going for " & intDays & " days" & ControlChars.NewLine)
        lstTotal.Items.Add("The insurance cost will be " & dblInsurance.ToString("C2") & ControlChars.NewLine)
        lstTotal.Items.Add("The daily cost will be " & dblCost.ToString("C2") & ControlChars.NewLine)
        lstTotal.Items.Add(ControlChars.NewLine)
        For i As Integer = 1 To intCampers
            lstTotal.Items.Add(i & " youth costs " & (dblCost * i).ToString("C2"))
        Next

        lstTotal.Items.Add("-----------------------------")
        lstTotal.Items.Add(ControlChars.NewLine)

        ' Checks for input conversion success/failure and calculates total cost.
        If blnCampersOk = True AndAlso blnDaysOk = True Then
            dblTotal = ((intCampers) * (dblCost) * (intDays)) + dblInsurance + FEE
            ' Display total cost.
            lstTotal.Items.Add("The total price due for the trip is " & dblTotal.ToString("C2"))
        Else
            ' Clears listbox and displays N/A if input conversion fails.
            lstTotal.Items.Clear()
            lstTotal.Items.Add("Values entered were not valid.")
        End If

        ' Calls BusNeed sub.

        BusNeed(intCampers)

    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ' Clears all textboxes, labels, listboxes and checkboxes.
        chkInsurance.Checked = False
        txtName.Text = Nothing
        txtDays.Text = Nothing
        txtCampers.Text = Nothing
        lblYN.Text = Nothing
        cboLocation.Text = Nothing
        lstTotal.Items.Clear()
    End Sub


    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        ' Closes application
        Me.Close()
    End Sub

    Private Sub ClearListBox(sender As Object, e As EventArgs) Handles txtCampers.TextChanged, txtDays.TextChanged, txtName.TextChanged, chkInsurance.CheckedChanged, cboLocation.SelectedValueChanged
        ' Clears listbox and bus label when user input changes.
        lstTotal.Items.Clear()
        lblYN.Text = Nothing
    End Sub
End Class
