'Title:         Close Program
'Date:          1-22-15
'Author:        Terry Holmes

'Description:   This will offer a choice to the user if they would like to close the program

Option Strict On

Public Class CloseProgram

    Private Sub btnNo_Click(sender As Object, e As EventArgs) Handles btnNo.Click

        'Returns to the form
        Me.Close()

    End Sub

    Private Sub btnYes_Click(sender As Object, e As EventArgs) Handles btnYes.Click

        'This will close the program
        Logon.Close()
        Form1.Close()
        Me.Close()

    End Sub
End Class