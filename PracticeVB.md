# college


Public Class frmShopping

    'event handler
    Private Sub btnGO_Click(sender As Object, e As EventArgs) Handles btnGO.Click
        'declare variables (needed to store data from UI and in order to perform calculations
        Dim intItems As Integer
        Dim decPrice As Decimal
        Dim decTotal As Decimal

        'validate the number of items textbox using selection.
        ' this uses and IF statement, which will only execute the code inside it if
        ' the condition (between the words 'If' and 'Then') is true
        If IsNumeric(txtItems.Text) = False Then
            MsgBox("Number of Items must be numeric!") ' display a message to the user
            Exit Sub 'stop the code here so that no further errors can occur.
        End If

        ' ADD other validation???

        'assign the variable with the value from the textbox.
        ' If we get here, we have passed validation, so it is safe to do this.
        intItems = txtItems.Text

        'loop.
        'this is a for loop (repetition) which repeats all code between the 'For' and
        ' the 'next' code words.  The code is repeated as many times are there are items 
        '  (which has been entered by the user).
        For intLoop = 1 To intItems
            'use an inputbox to ask the user for a price, storing it in the appropriate variable
            decPrice = InputBox("Please enter the price:")
            'add the price just entered to the running total
            decTotal = decTotal + decPrice
        Next

        'display in labels
        lblTotal.Text = "£" & decTotal 'display the total
        'call the function (passing it the total as a parameter.
        '  the result that is returned from the function is displayed in the discount label
        lblDiscountTotal.Text = "£" & ApplyDiscount(decTotal)
    End Sub

    ' event handler for clicking on reset button
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Reset() ' call the procedure RESET
    End Sub

    ' A procedure - these are a self-contained section of code that does a particular job
    ' this procedure resets the GUI items back to their starting values
    Private Sub Reset()
        lblTotal.Text = "£0"
        lblDiscountTotal.Text = "£0"
        txtItems.Text = 0
    End Sub

    ' a function - this performs a job, but returns a value back to the code it was called from
    ' this function takes 5% off the total and returns the new (discounted) total.
    'This function accepts a parameter, which is a value the function needs to do its job.
    Function ApplyDiscount(ByVal total As Integer) ' the parameter is total
        Return total * 0.95
    End Function


End Class
