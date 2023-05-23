<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Report_Form
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.report_viewer = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.SuspendLayout()
        '
        'report_viewer
        '
        Me.report_viewer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.report_viewer.LocalReport.ReportEmbeddedResource = "OLEDB_MSACCESS_APP.inventory_report.rdlc"
        Me.report_viewer.Location = New System.Drawing.Point(0, 0)
        Me.report_viewer.Name = "report_viewer"
        Me.report_viewer.ServerReport.BearerToken = Nothing
        Me.report_viewer.Size = New System.Drawing.Size(800, 450)
        Me.report_viewer.TabIndex = 0
        '
        'Report_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.report_viewer)
        Me.Name = "Report_Form"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Report"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents report_viewer As Microsoft.Reporting.WinForms.ReportViewer
End Class
