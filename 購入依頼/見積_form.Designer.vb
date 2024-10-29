<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 見積_form
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
        Me.DGV見積 = New System.Windows.Forms.DataGridView()
        Me.btn発注申請 = New System.Windows.Forms.Button()
        Me.btn見積印刷 = New System.Windows.Forms.Button()
        Me.btn見積to製品 = New System.Windows.Forms.Button()
        Me.btn見積_更新 = New System.Windows.Forms.Button()
        Me.btn削除 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.DGV見積, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DGV見積
        '
        Me.DGV見積.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV見積.Location = New System.Drawing.Point(28, 45)
        Me.DGV見積.Name = "DGV見積"
        Me.DGV見積.RowTemplate.Height = 21
        Me.DGV見積.Size = New System.Drawing.Size(1303, 577)
        Me.DGV見積.TabIndex = 0
        '
        'btn発注申請
        '
        Me.btn発注申請.Location = New System.Drawing.Point(1129, 12)
        Me.btn発注申請.Name = "btn発注申請"
        Me.btn発注申請.Size = New System.Drawing.Size(92, 27)
        Me.btn発注申請.TabIndex = 1
        Me.btn発注申請.Text = "発注申請"
        Me.btn発注申請.UseVisualStyleBackColor = True
        '
        'btn見積印刷
        '
        Me.btn見積印刷.Location = New System.Drawing.Point(1227, 12)
        Me.btn見積印刷.Name = "btn見積印刷"
        Me.btn見積印刷.Size = New System.Drawing.Size(104, 27)
        Me.btn見積印刷.TabIndex = 2
        Me.btn見積印刷.Text = "見積依頼書印刷"
        Me.btn見積印刷.UseVisualStyleBackColor = True
        '
        'btn見積to製品
        '
        Me.btn見積to製品.Location = New System.Drawing.Point(28, 12)
        Me.btn見積to製品.Name = "btn見積to製品"
        Me.btn見積to製品.Size = New System.Drawing.Size(119, 27)
        Me.btn見積to製品.TabIndex = 3
        Me.btn見積to製品.Text = "製品へコピー"
        Me.btn見積to製品.UseVisualStyleBackColor = True
        '
        'btn見積_更新
        '
        Me.btn見積_更新.BackColor = System.Drawing.Color.ForestGreen
        Me.btn見積_更新.ForeColor = System.Drawing.Color.Black
        Me.btn見積_更新.Location = New System.Drawing.Point(357, 628)
        Me.btn見積_更新.Name = "btn見積_更新"
        Me.btn見積_更新.Size = New System.Drawing.Size(664, 27)
        Me.btn見積_更新.TabIndex = 4
        Me.btn見積_更新.Text = "更新"
        Me.btn見積_更新.UseVisualStyleBackColor = False
        '
        'btn削除
        '
        Me.btn削除.BackColor = System.Drawing.Color.LightCoral
        Me.btn削除.Location = New System.Drawing.Point(274, 628)
        Me.btn削除.Name = "btn削除"
        Me.btn削除.Size = New System.Drawing.Size(77, 27)
        Me.btn削除.TabIndex = 5
        Me.btn削除.Text = "削除"
        Me.btn削除.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label1.Location = New System.Drawing.Point(787, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(336, 22)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "発注申請：見積状態が「発注希望」の行をメイン画面に移動させます" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "見積依頼書印刷：見積状態が「見積希望」の行の見積依頼書を印刷します"
        '
        '見積_form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1343, 665)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btn削除)
        Me.Controls.Add(Me.btn見積_更新)
        Me.Controls.Add(Me.btn見積to製品)
        Me.Controls.Add(Me.btn見積印刷)
        Me.Controls.Add(Me.btn発注申請)
        Me.Controls.Add(Me.DGV見積)
        Me.Name = "見積_form"
        Me.Text = "見積"
        CType(Me.DGV見積, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DGV見積 As System.Windows.Forms.DataGridView
    Friend WithEvents btn発注申請 As System.Windows.Forms.Button
    Friend WithEvents btn見積印刷 As System.Windows.Forms.Button
    Friend WithEvents btn見積to製品 As System.Windows.Forms.Button
    Friend WithEvents btn見積_更新 As System.Windows.Forms.Button
    Friend WithEvents btn削除 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
