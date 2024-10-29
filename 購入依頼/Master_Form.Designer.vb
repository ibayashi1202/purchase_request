<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Master_Form
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.DGV_マスタ = New System.Windows.Forms.DataGridView()
        Me.cmbマスタ種類 = New System.Windows.Forms.ComboBox()
        CType(Me.DGV_マスタ, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DGV_マスタ
        '
        Me.DGV_マスタ.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV_マスタ.Location = New System.Drawing.Point(12, 31)
        Me.DGV_マスタ.Name = "DGV_マスタ"
        Me.DGV_マスタ.RowTemplate.Height = 21
        Me.DGV_マスタ.Size = New System.Drawing.Size(1035, 568)
        Me.DGV_マスタ.TabIndex = 0
        '
        'cmbマスタ種類
        '
        Me.cmbマスタ種類.FormattingEnabled = True
        Me.cmbマスタ種類.Location = New System.Drawing.Point(12, 5)
        Me.cmbマスタ種類.Name = "cmbマスタ種類"
        Me.cmbマスタ種類.Size = New System.Drawing.Size(121, 20)
        Me.cmbマスタ種類.TabIndex = 1
        '
        'Master_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1059, 611)
        Me.Controls.Add(Me.cmbマスタ種類)
        Me.Controls.Add(Me.DGV_マスタ)
        Me.Name = "Master_Form"
        Me.Text = "マスタ"
        CType(Me.DGV_マスタ, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DGV_マスタ As System.Windows.Forms.DataGridView
    Friend WithEvents cmbマスタ種類 As System.Windows.Forms.ComboBox
End Class
