Public Class Report_Form
    Private Sub Report_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        report_viewer.Show()

        report_viewer.RefreshReport()
    End Sub
End Class