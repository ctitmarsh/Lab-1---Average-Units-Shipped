Option Strict On    ' Force compiler to adhere to more strict rules

' Author:         Austin Garrod
' Banner Number:  100607615
' File Name:      MainForm.vb
' Created:        2019/01/14
' Updated:        2019/01/14
' Version:        1.0
' Description:    A VB system to take user inputed units, store them, and return the
'                 average units shipped per day over the period entered

Public Class MainForm

#Region "Variable and constant Declaration"

    ' Declare Constants
    Const MIN_VALUE As Integer = 0      ' Minimum allowed unit volume
    Const MAX_VALUE As Integer = 1000   ' Maximum allowed unit volume
    Const ERROR_MESSAGE As String = "Invalid input, please try again" ' error message to be returned on incorrect data entry

    ' Declare Variables
    Dim units(6) As Integer ' Holds entered units
    Dim day As Integer = 1  ' Holds current day of data entry

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

#End Region

#Region "Functions and Subs"


    ''' <summary>
    '''     Resets form to initial state
    ''' </summary>
    Sub resetForm()
        ' Reset variables
        day = 1
        Array.Clear(units, 0, units.Length)

        ' Reset fields and labels
        txtUnits.Text = ""
        lblCurrentDay.Text = "Day 1"
        txtDataDisplay.Text = ""
        lblOutput.Text = ""

        ' Ensure user's ability to enter data
        btnEnter.Enabled = True
        txtUnits.ReadOnly = False

        ' Set focus back to input box for usability
        txtUnits.Focus()
    End Sub

    ''' <summary>
    '''     Updates data display to contain values currently stored in units array
    ''' </summary>
    Sub updateDataDisplay()
        'txtDataDisplay.Text += newData + vbCrLf

        Dim formattedOutput As String = ""  ' Holds formatted data to display in text box

        ' Example of for loop in VB
        For counter As Integer = 1 To day Step 1    ' Loop through each day's entered data
            formattedOutput += units(counter - 1).ToString + vbCrLf ' Insert each data point into the text box and start a new line
        Next

        txtDataDisplay.Text = formattedOutput ' Output the formatted data in the dataDisplay field
    End Sub

    ''' <summary>
    '''     Validates inputed user data
    ''' </summary>
    ''' <param name="input">User input to be validated</param>
    ''' <returns>Whether the input was valid or invalid as boolean</returns>
    Function validateInput(ByVal input As String) As Boolean
        Dim inputNumber As Integer  ' Holds the user input as an integer
        Dim isValidInput As Boolean = False ' Holds whether the input is valid or not, assuming data is invalid

        ' Example of a try/ catch in VB
        Try
            inputNumber = CInt(input)   ' Try to cast the input as an integer
            If (input.Equals(inputNumber.ToString)) Then   ' Check if the inputed data is a whole number
                If (inputNumber >= MIN_VALUE AndAlso inputNumber <= MAX_VALUE) Then ' Check if the inputed data is within defined bounds
                    isValidInput = True ' All checks passed, data is valid
                End If
            End If
        Catch ex As Exception
            ' input could not be cast as an integer, not a number
        End Try

        Return isValidInput ' Return the success or failure of the validation
    End Function

    ''' <summary>
    '''     Returns the average value of a provided array
    ''' </summary>
    ''' <param name="arrayToAverage">Array to be averaged</param>
    ''' <returns>Average of provided array</returns>
    Function averageArray(ByVal arrayToAverage() As Integer) As Double
        Dim runningTotal As Integer ' Holds a running total of the values in the array

        ' Example of a For Each loop in VB
        For Each dailyTotal In arrayToAverage  ' Loop through the array
            runningTotal += dailyTotal  ' Add the current value to the running total
        Next

        averageArray = Math.Round(runningTotal / arrayToAverage.Length, 2) ' Return the average value of the array, rounded to 2 decimal places
    End Function

#End Region

#Region "Event Handlers"

    ''' <summary>
    '''     Handles reset button
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Call resetForm()    ' Reset the form
    End Sub

    ''' <summary>
    '''     Handles enter button
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnEnter_Click(sender As Object, e As EventArgs) Handles btnEnter.Click
        Dim userInput As String = txtUnits.Text ' Take user data and store it for use

        If (validateInput(userInput)) Then ' Check if data entered by user is valid
            lblOutput.Text = "" ' Reset the data entry field to prep for next data entry

            units(day - 1) = Convert.ToInt32(userInput) ' Add new unit value to array of units
            Call updateDataDisplay()                    ' Update display to show new units

            If (day < 7) Then   ' Check if there is still data to be entered
                day = day + 1                               ' Increment day counter to next day
                lblCurrentDay.Text = "Day " + day.ToString  ' Display updated day count
            Else ' No new data to be entered
                btnEnter.Enabled = False    ' Prevent the user from pressing enter
                txtUnits.ReadOnly = True    ' Prevent the user from entering new text in units box
                lblOutput.Text = "Avg Units/Day: " + averageArray(units).ToString ' Display the rounded average
            End If

        Else
            lblOutput.Text = ERROR_MESSAGE ' Entered data didn't validate, return error
        End If

        txtUnits.Focus()    ' Set focus back to entry box for usability
        txtUnits.Text = ""  ' Clear the last data entry from the text box
    End Sub

    ''' <summary>
    '''     Handles exit button
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnExit_Click(sender As Object, e As EventArgs)
        Application.Exit() ' Gracefully tell the application to exit
    End Sub

    Private Sub btnExit_Click_1(sender As Object, e As EventArgs) Handles btnExit.Click
        Application.Exit()
    End Sub

    Private Sub txtUnits_TextChanged(sender As Object, e As EventArgs) Handles txtUnits.LostFocus

    End Sub

#End Region

End Class
