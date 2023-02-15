<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DueDateExceededJobDetails
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
        Me.LateJobGridView = New System.Windows.Forms.DataGridView()
        Me.LabelDelayedOrders = New System.Windows.Forms.Label()
        CType(Me.LateJobGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LateJobGridView
        '
        Me.LateJobGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.LateJobGridView.Location = New System.Drawing.Point(12, 64)
        Me.LateJobGridView.Name = "LateJobGridView"
        Me.LateJobGridView.RowHeadersWidth = 62
        Me.LateJobGridView.RowTemplate.Height = 28
        Me.LateJobGridView.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.LateJobGridView.Size = New System.Drawing.Size(1200, 374)
        Me.LateJobGridView.TabIndex = 0
        '
        'LabelDelayedOrders
        '
        Me.LabelDelayedOrders.AutoSize = True
        Me.LabelDelayedOrders.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelDelayedOrders.Location = New System.Drawing.Point(501, 18)
        Me.LabelDelayedOrders.Name = "LabelDelayedOrders"
        Me.LabelDelayedOrders.Size = New System.Drawing.Size(213, 32)
        Me.LabelDelayedOrders.TabIndex = 1
        Me.LabelDelayedOrders.Text = "Delayed Orders"
        '
        'DueDateExceededJobDetails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1246, 450)
        Me.Controls.Add(Me.LabelDelayedOrders)
        Me.Controls.Add(Me.LateJobGridView)
        Me.Name = "DueDateExceededJobDetails"
        Me.Text = "Delayed Orders"
        CType(Me.LateJobGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LateJobGridView As Windows.Forms.DataGridView
    Friend WithEvents LabelDelayedOrders As Windows.Forms.Label
End Class
