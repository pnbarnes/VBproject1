' Name of Program: pb_assignment2.
' Purpose: To calculate costs for Youth Camp.
' Programmer: Paige B
Public Class frmAssignment2

    ' Constant variable for travel fee, $300.
    Private Const dblFEE As Double = 300

    Private Sub btnCalc_Click(sender As Object, e As EventArgs) Handles btnCalc.Click
        ' Radio button variables.
        Dim strLocation As String
        Dim dblDailyCost As Double
        ' Checkbox variable.
        Dim dblInsurance As Double
        ' Variables for user input text box contents.
        Dim intCampers As Integer
        Dim intCampDays As Integer
        ' Variable for running total.
        Dim dblTotal As Double
        ' Variables for error checking.
        Dim blnCampersOk As Boolean
        Dim blnCampDaysOk As Boolean

        ' Error checks.
        If txtCampers.Text = "" Then
            MessageBox.Show("Please enter an amount")
            txtCampers.Focus()
            Return
        End If
        If txtCampDays.Text = "" Then
            MessageBox.Show("Please enter an amount")
            txtCampDays.Focus()
            Return
        End If
        If radAuTrain50.Checked = False And radHocking150.Checked = False And radMunising100.Checked = False And radNewYork500.Checked = False Then
            MessageBox.Show("Please select a location")
            Return
        End If



        ' If statement for defining checkbox variable.
        If chkInsurance.Checked = True Then
            dblInsurance = 150
        Else
            dblInsurance = 0
        End If

        ' Case Statements for radio buttons.
        If radNewYork500.Checked Then
            strLocation = "New York"
        ElseIf radMunising100.Checked Then
            strLocation = "Munising"
        ElseIf radAuTrain50.Checked Then
            strLocation = "Au Train"
        ElseIf radHocking150.Checked Then
            strLocation = "Hocking Hills"
        Else
            MessageBox.Show("Please select a location")
            Return
        End If

        Select Case strLocation
            Case "New York"
                dblDailyCost = 500
            Case "Munising"
                dblDailyCost = 100
            Case "Au Train"
                dblDailyCost = 50
            Case "Hocking Hills"
                dblDailyCost = 150
        End Select

        ' Assigns user input text box contents to variables and checks for errors.
        blnCampDaysOk = Integer.TryParse(txtCampDays.Text, intCampDays)
        blnCampersOk = Integer.TryParse(txtCampers.Text, intCampers)
        If blnCampersOk AndAlso blnCampersOk Then
            ' Calculates total.
            ' ((number of people) * (daily cost of trip) * (number of days)) + insurance amount + fee
            dblTotal = ((intCampers) * (dblDailyCost) * (intCampDays)) + dblInsurance + dblFEE

            ' Displays total in Total textbox as currency.
            txtTotal.Text = dblTotal.ToString("C2")
        Else
            txtTotal.Text = "N/A"
        End If

    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ' Clears all text boxes.
        txtCampDays.Text = Nothing
        txtCampers.Text = Nothing
        txtTotal.Text = Nothing
        chkInsurance.Checked = False
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        ' Closes application.
        Me.Close()
    End Sub

    Private Sub chkInsurance_CheckedChanged(sender As Object, e As EventArgs) Handles chkInsurance.CheckedChanged
        ' Clears Total when the Insurance checked state changes.
        txtTotal.Text = Nothing
    End Sub

    Private Sub txtCampDays_TextChanged(sender As Object, e As EventArgs) Handles txtCampDays.TextChanged
        ' Clears Total when the Days of Camp text changes.
        txtTotal.Text = Nothing
    End Sub

    Private Sub txtCampers_TextChanged(sender As Object, e As EventArgs) Handles txtCampers.TextChanged
        ' Clears Total when the Campers text changes.
        txtTotal.Text = Nothing
    End Sub

    Private Sub txtCampDays_Enter(sender As Object, e As EventArgs) Handles txtCampDays.Enter
        ' Selects all contents when user tabs to Days of Camp.
        txtCampDays.SelectAll()
    End Sub

    Private Sub txtCampers_Enter(sender As Object, e As EventArgs) Handles txtCampers.Enter
        ' Selects all contents when user tabs to Number of Campers.
        txtCampers.SelectAll()
    End Sub
End Class
