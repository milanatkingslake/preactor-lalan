Public Class DueDateExceededJobDetails
    Public Property tblDueDateExcJob As DataTable
    Public Property connetionString As String
    Private Sub DueDateExceededJobDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LateJobGridView.DataSource = tblDueDateExcJob
        LateJobGridView.Refresh()
        LateJobGridView.AutoResizeColumns()
    End Sub
End Class