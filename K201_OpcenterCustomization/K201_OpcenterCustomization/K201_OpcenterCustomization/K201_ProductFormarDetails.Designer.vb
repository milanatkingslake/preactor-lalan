<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class K201_ProductFormarDetails
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.DataGridViewCoProduct = New System.Windows.Forms.DataGridView()
        Me.btnConfirmResourceRate = New System.Windows.Forms.Button()
        CType(Me.DataGridViewCoProduct, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridViewCoProduct
        '
        Me.DataGridViewCoProduct.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewCoProduct.Location = New System.Drawing.Point(7, 12)
        Me.DataGridViewCoProduct.Name = "DataGridViewCoProduct"
        Me.DataGridViewCoProduct.RowHeadersWidth = 62
        Me.DataGridViewCoProduct.Size = New System.Drawing.Size(766, 367)
        Me.DataGridViewCoProduct.TabIndex = 0
        '
        'btnConfirmResourceRate
        '
        Me.btnConfirmResourceRate.Location = New System.Drawing.Point(779, 144)
        Me.btnConfirmResourceRate.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnConfirmResourceRate.Name = "btnConfirmResourceRate"
        Me.btnConfirmResourceRate.Size = New System.Drawing.Size(151, 29)
        Me.btnConfirmResourceRate.TabIndex = 1
        Me.btnConfirmResourceRate.Text = "Confirm Resource Rate"
        Me.btnConfirmResourceRate.UseMnemonic = False
        Me.btnConfirmResourceRate.UseVisualStyleBackColor = False
        '
        'K201_ProductFormarDetails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(938, 472)
        Me.Controls.Add(Me.btnConfirmResourceRate)
        Me.Controls.Add(Me.DataGridViewCoProduct)
        Me.Name = "K201_ProductFormarDetails"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Product Former Ratio Calculation"
        CType(Me.DataGridViewCoProduct, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DataGridViewCoProduct As Windows.Forms.DataGridView
    Friend WithEvents btnConfirmResourceRate As Windows.Forms.Button
End Class
