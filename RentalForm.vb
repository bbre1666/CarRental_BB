Option Explicit On
Option Strict On
Option Compare Binary
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

'Baden Brenner  
'RCET 0265
'Spring 23
'Car Rental
'

Public Class RentalForm
    Dim totalCustomers As Integer
    Dim totalDistance As Integer
    Dim totalCharges As Integer



    Private Function ValidateInput() As String
        Dim errors As String


        ' Validate odometer readings
        Dim beginOdometer As Double
        Dim endOdometer As Double
        If Not Double.TryParse(BeginOdometerTextBox.Text, beginOdometer) _
            Or Not Double.TryParse(EndOdometerTextBox.Text, endOdometer) Then
            errors &= ("Odometer readings must be numbers." & vbNewLine)
            BeginOdometerTextBox.Focus()
        ElseIf beginOdometer >= endOdometer Then
            errors &= ("Beginning odometer reading must be less than ending odometer reading." & vbNewLine)
            BeginOdometerTextBox.Focus()
        End If

        ' Validate number of days
        Dim numDays As Integer
        If Not Integer.TryParse(DaysTextBox.Text, numDays) Then
            errors &= ("Number of days must be a whole number." & vbNewLine)
            DaysTextBox.Focus()
        ElseIf numDays <= 0 Then
            errors &= ("Number of days must be greater than zero." & vbNewLine)
            DaysTextBox.Focus()
        ElseIf numDays > 45 Then
            errors &= ("Number of days cannot be greater than 45." & vbNewLine)
            DaysTextBox.Focus()
        End If

        ' Validate customer information
        If ZipCodeTextBox.Text = "" Then
            errors &= ("ZipCode field cannot be blank." & vbNewLine)
            ZipCodeTextBox.Focus()
        End If
        If StateTextBox.Text = "" Then
            errors &= ("State field cannot be blank." & vbNewLine)
            StateTextBox.Focus()
        End If
        If CityTextBox.Text = "" Then
            errors &= ("City field cannot be blank." & vbNewLine)
            CityTextBox.Focus()
        End If
        If AddressTextBox.Text = "" Then
            errors &= ("Address field cannot be blank." & vbNewLine)
            AddressTextBox.Focus()
        End If
        If NameTextBox.Text = "" Then
            errors &= ("Name field cannot be blank." & vbNewLine)
            NameTextBox.Focus()
        End If
        'checking to see if the errors string has errors that have been added to or is empty
        If String.IsNullOrEmpty(errors) OrElse String.IsNullOrWhiteSpace(errors) Then
            Calculate()
        Else
            MsgBox(errors)
        End If
    End Function
    Private Sub Clear()
        NameTextBox.Clear()
        AddressTextBox.Clear()
        CityTextBox.Clear()
        StateTextBox.Clear()
        ZipCodeTextBox.Clear()
        BeginOdometerTextBox.Clear()
        EndOdometerTextBox.Clear()
        DaysTextBox.Clear()
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False

        TotalMilesTextBox.Clear()
        MileageChargeTextBox.Clear()
        DayChargeTextBox.Clear()
        TotalDiscountTextBox.Clear()
        TotalChargeTextBox.Clear()
    End Sub
    Private Sub Calculate()

        ' convert textbox values to numeric data types
        Dim BeginOdometer As Integer
        Dim EndOdometer As Integer
        If KilometersradioButton.Checked Then
            BeginOdometer = (Integer.Parse(BeginOdometerTextBox.Text))
            EndOdometer = (Integer.Parse(EndOdometerTextBox.Text))
            BeginOdometer = (CInt(BeginOdometer * 0.62))
            EndOdometer = (CInt(EndOdometer * 0.62))
        Else
            BeginOdometer = Integer.Parse(BeginOdometerTextBox.Text)
            EndOdometer = Integer.Parse(EndOdometerTextBox.Text)

        End If


        ' perform arithmetic operation
        Dim MilesNumber As Integer = EndOdometer - BeginOdometer

        ' convert result back to a string and display it in another textbox
        TotalMilesTextBox.Text = (MilesNumber.ToString())
        TotalMilesTextBox.AppendText("mi")

        Dim milesDriven As String = (MilesNumber.ToString())
        Dim cost As Double

        If CInt(milesDriven) <= 200 Then
            ' first 200 miles are free
            cost = 0
        ElseIf CInt(milesDriven) <= 500 Then
            ' miles between 201 and 500 inclusive are 12 cents per mile
            cost = (CInt(milesDriven) - 200) * 0.12
        Else
            ' miles greater than 500 are charged at 10 cents per mile
            cost = (300 * 0.12) + (CInt(milesDriven) - 500) * 0.1
        End If

        MileageChargeTextBox.Text = cost.ToString("C2")

        'dalily charge calculation
        Dim days As Integer
        Dim TotalChargeValue As Integer
        Dim Totaldiscount As Integer

        days = Integer.Parse(DaysTextBox.Text)

        DayChargeTextBox.Text = ((days * 15).ToString("C"))

        TotalChargeValue = CInt((days * 15) + cost)
        If AAAcheckbox.Checked Then
            TotalChargeValue = CInt(TotalChargeValue * 0.95)
            Totaldiscount = CInt(TotalChargeValue * 0.05)
        End If

        If Seniorcheckbox.Checked Then
            TotalChargeValue = CInt(TotalChargeValue * 0.97)
            Totaldiscount = CInt(Totaldiscount + TotalChargeValue * 0.03)
        End If
        TotalDiscountTextBox.Text = (Totaldiscount.ToString("C"))
        TotalChargeTextBox.Text = CStr(TotalChargeValue.ToString("C"))

        totalCustomers += 1
        totalDistance = (totalDistance + (CInt(milesDriven)))
        totalCharges = (totalCharges + (CInt(TotalChargeValue)))
        SummaryButton.Enabled = True

    End Sub
    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        MessageBox.Show(String.Format("Total customers: {0}{1}Total distance driven: {2} miles{1}Total charges: {3:C}",
                      totalCustomers, Environment.NewLine, totalDistance, totalCharges))
        Clear()
    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to exit?", "Confirm Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            ' User clicked Yes, so exit the application
            Application.Exit()
        End If
    End Sub

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        ValidateInput()
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        Clear()
    End Sub

    Private Sub RentalForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        SummaryButton.Enabled = False
    End Sub
End Class
