<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Select_Form
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
        Me.btn部署選択 = New System.Windows.Forms.Button()
        Me.lst部署 = New System.Windows.Forms.ListBox()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.lblcatalog = New System.Windows.Forms.Label()
        Me.btn決裁 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btn部署選択
        '
        Me.btn部署選択.Location = New System.Drawing.Point(12, 198)
        Me.btn部署選択.Name = "btn部署選択"
        Me.btn部署選択.Size = New System.Drawing.Size(228, 29)
        Me.btn部署選択.TabIndex = 6
        Me.btn部署選択.Text = "標準レイアウトで開く"
        Me.btn部署選択.UseVisualStyleBackColor = True
        '
        'lst部署
        '
        Me.lst部署.FormattingEnabled = True
        Me.lst部署.ItemHeight = 12
        Me.lst部署.Location = New System.Drawing.Point(12, 20)
        Me.lst部署.Name = "lst部署"
        Me.lst部署.Size = New System.Drawing.Size(228, 172)
        Me.lst部署.TabIndex = 5
        '
        'lbl1
        '
        Me.lbl1.AutoSize = True
        Me.lbl1.Location = New System.Drawing.Point(12, 5)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(29, 12)
        Me.lbl1.TabIndex = 4
        Me.lbl1.Text = "部署"
        '
        'lblcatalog
        '
        Me.lblcatalog.AutoSize = True
        Me.lblcatalog.ForeColor = System.Drawing.Color.Blue
        Me.lblcatalog.Location = New System.Drawing.Point(10, 295)
        Me.lblcatalog.Name = "lblcatalog"
        Me.lblcatalog.Size = New System.Drawing.Size(54, 12)
        Me.lblcatalog.TabIndex = 7
        Me.lblcatalog.Text = "lblcatalog"
        '
        'btn決裁
        '
        Me.btn決裁.Location = New System.Drawing.Point(12, 233)
        Me.btn決裁.Name = "btn決裁"
        Me.btn決裁.Size = New System.Drawing.Size(228, 28)
        Me.btn決裁.TabIndex = 8
        Me.btn決裁.Text = "承認レイアウトで開く"
        Me.btn決裁.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label1.Location = New System.Drawing.Point(77, 278)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(0, 11)
        Me.Label1.TabIndex = 9
        '
        'Select_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(252, 319)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btn決裁)
        Me.Controls.Add(Me.lblcatalog)
        Me.Controls.Add(Me.btn部署選択)
        Me.Controls.Add(Me.lst部署)
        Me.Controls.Add(Me.lbl1)
        Me.Name = "Select_Form"
        Me.Text = "Select_Form"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btn部署選択 As System.Windows.Forms.Button
    Friend WithEvents lst部署 As System.Windows.Forms.ListBox
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents lblcatalog As System.Windows.Forms.Label
    Friend WithEvents btn決裁 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
