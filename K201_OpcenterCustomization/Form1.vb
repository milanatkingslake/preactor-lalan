Public Class DamagePercentageForm
    Public Property isOkClick As Boolean = False
    Public Property damage As Integer
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles DamagePercentage.Click

    End Sub

    Private Sub DamagePercentageForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub OkBtn_Click(sender As Object, e As EventArgs) Handles OkBtn.Click
        isOkClick = True
        Close()
    End Sub
End Class