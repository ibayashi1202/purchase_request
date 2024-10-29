<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 製品_Form
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
        Me.DGV製品 = New System.Windows.Forms.DataGridView()
        Me.btn製品to購入 = New System.Windows.Forms.Button()
        Me.btn製品_更新 = New System.Windows.Forms.Button()
        Me.btn製品_見積コピー = New System.Windows.Forms.Button()
        Me.btn削除 = New System.Windows.Forms.Button()
        Me.btn製品検索 = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmb検索項目3 = New System.Windows.Forms.ComboBox()
        Me.cmb記号3 = New System.Windows.Forms.ComboBox()
        Me.txt検索条件3 = New System.Windows.Forms.TextBox()
        Me.cmb検索項目2 = New System.Windows.Forms.ComboBox()
        Me.cmb記号2 = New System.Windows.Forms.ComboBox()
        Me.txt検索条件2 = New System.Windows.Forms.TextBox()
        Me.cmb検索項目1 = New System.Windows.Forms.ComboBox()
        Me.cmb記号1 = New System.Windows.Forms.ComboBox()
        Me.txt検索条件1 = New System.Windows.Forms.TextBox()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.cmb検索条件選択1 = New System.Windows.Forms.ComboBox()
        Me.cmb検索条件選択2 = New System.Windows.Forms.ComboBox()
        Me.cmb検索条件選択3 = New System.Windows.Forms.ComboBox()
        Me.btn貼り付け = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnExcel = New System.Windows.Forms.Button()
        CType(Me.DGV製品, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DGV製品
        '
        Me.DGV製品.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV製品.Location = New System.Drawing.Point(12, 40)
        Me.DGV製品.Name = "DGV製品"
        Me.DGV製品.RowTemplate.Height = 21
        Me.DGV製品.Size = New System.Drawing.Size(1184, 554)
        Me.DGV製品.TabIndex = 0
        '
        'btn製品to購入
        '
        Me.btn製品to購入.Location = New System.Drawing.Point(13, 600)
        Me.btn製品to購入.Name = "btn製品to購入"
        Me.btn製品to購入.Size = New System.Drawing.Size(125, 25)
        Me.btn製品to購入.TabIndex = 1
        Me.btn製品to購入.Text = "購入品入力へコピー"
        Me.btn製品to購入.UseVisualStyleBackColor = True
        '
        'btn製品_更新
        '
        Me.btn製品_更新.BackColor = System.Drawing.Color.ForestGreen
        Me.btn製品_更新.ForeColor = System.Drawing.Color.White
        Me.btn製品_更新.Location = New System.Drawing.Point(440, 600)
        Me.btn製品_更新.Name = "btn製品_更新"
        Me.btn製品_更新.Size = New System.Drawing.Size(260, 25)
        Me.btn製品_更新.TabIndex = 2
        Me.btn製品_更新.Text = "更新"
        Me.btn製品_更新.UseVisualStyleBackColor = False
        '
        'btn製品_見積コピー
        '
        Me.btn製品_見積コピー.Location = New System.Drawing.Point(144, 600)
        Me.btn製品_見積コピー.Name = "btn製品_見積コピー"
        Me.btn製品_見積コピー.Size = New System.Drawing.Size(108, 25)
        Me.btn製品_見積コピー.TabIndex = 3
        Me.btn製品_見積コピー.Text = "見積へコピー"
        Me.btn製品_見積コピー.UseVisualStyleBackColor = True
        '
        'btn削除
        '
        Me.btn削除.BackColor = System.Drawing.Color.LightCoral
        Me.btn削除.Location = New System.Drawing.Point(339, 601)
        Me.btn削除.Name = "btn削除"
        Me.btn削除.Size = New System.Drawing.Size(95, 25)
        Me.btn削除.TabIndex = 5
        Me.btn削除.Text = "削除"
        Me.btn削除.UseVisualStyleBackColor = False
        '
        'btn製品検索
        '
        Me.btn製品検索.Location = New System.Drawing.Point(1067, 15)
        Me.btn製品検索.Name = "btn製品検索"
        Me.btn製品検索.Size = New System.Drawing.Size(60, 20)
        Me.btn製品検索.TabIndex = 6
        Me.btn製品検索.Text = "検索"
        Me.btn製品検索.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(725, 19)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(35, 12)
        Me.Label4.TabIndex = 53
        Me.Label4.Text = "条件3"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(399, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(35, 12)
        Me.Label3.TabIndex = 52
        Me.Label3.Text = "条件2"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(82, 12)
        Me.Label1.TabIndex = 51
        Me.Label1.Text = "絞り込み 条件1"
        '
        'cmb検索項目3
        '
        Me.cmb検索項目3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb検索項目3.FormattingEnabled = True
        Me.cmb検索項目3.Location = New System.Drawing.Point(766, 15)
        Me.cmb検索項目3.Name = "cmb検索項目3"
        Me.cmb検索項目3.Size = New System.Drawing.Size(106, 20)
        Me.cmb検索項目3.TabIndex = 50
        '
        'cmb記号3
        '
        Me.cmb記号3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb記号3.FormattingEnabled = True
        Me.cmb記号3.Location = New System.Drawing.Point(878, 15)
        Me.cmb記号3.Name = "cmb記号3"
        Me.cmb記号3.Size = New System.Drawing.Size(54, 20)
        Me.cmb記号3.TabIndex = 49
        '
        'txt検索条件3
        '
        Me.txt検索条件3.Location = New System.Drawing.Point(938, 16)
        Me.txt検索条件3.Name = "txt検索条件3"
        Me.txt検索条件3.Size = New System.Drawing.Size(104, 19)
        Me.txt検索条件3.TabIndex = 48
        '
        'cmb検索項目2
        '
        Me.cmb検索項目2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb検索項目2.FormattingEnabled = True
        Me.cmb検索項目2.Location = New System.Drawing.Point(440, 15)
        Me.cmb検索項目2.Name = "cmb検索項目2"
        Me.cmb検索項目2.Size = New System.Drawing.Size(106, 20)
        Me.cmb検索項目2.TabIndex = 47
        '
        'cmb記号2
        '
        Me.cmb記号2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb記号2.FormattingEnabled = True
        Me.cmb記号2.Location = New System.Drawing.Point(552, 15)
        Me.cmb記号2.Name = "cmb記号2"
        Me.cmb記号2.Size = New System.Drawing.Size(55, 20)
        Me.cmb記号2.TabIndex = 46
        '
        'txt検索条件2
        '
        Me.txt検索条件2.Location = New System.Drawing.Point(613, 16)
        Me.txt検索条件2.Name = "txt検索条件2"
        Me.txt検索条件2.Size = New System.Drawing.Size(87, 19)
        Me.txt検索条件2.TabIndex = 45
        '
        'cmb検索項目1
        '
        Me.cmb検索項目1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb検索項目1.FormattingEnabled = True
        Me.cmb検索項目1.Location = New System.Drawing.Point(100, 15)
        Me.cmb検索項目1.Name = "cmb検索項目1"
        Me.cmb検索項目1.Size = New System.Drawing.Size(106, 20)
        Me.cmb検索項目1.TabIndex = 44
        '
        'cmb記号1
        '
        Me.cmb記号1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb記号1.FormattingEnabled = True
        Me.cmb記号1.Location = New System.Drawing.Point(212, 15)
        Me.cmb記号1.Name = "cmb記号1"
        Me.cmb記号1.Size = New System.Drawing.Size(54, 20)
        Me.cmb記号1.TabIndex = 43
        '
        'txt検索条件1
        '
        Me.txt検索条件1.Location = New System.Drawing.Point(272, 16)
        Me.txt検索条件1.Name = "txt検索条件1"
        Me.txt検索条件1.Size = New System.Drawing.Size(91, 19)
        Me.txt検索条件1.TabIndex = 42
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(1127, 15)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(69, 20)
        Me.btnClear.TabIndex = 54
        Me.btnClear.Text = "条件クリア"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'cmb検索条件選択1
        '
        Me.cmb検索条件選択1.DropDownHeight = 600
        Me.cmb検索条件選択1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb検索条件選択1.DropDownWidth = 250
        Me.cmb検索条件選択1.FormattingEnabled = True
        Me.cmb検索条件選択1.IntegralHeight = False
        Me.cmb検索条件選択1.Location = New System.Drawing.Point(369, 15)
        Me.cmb検索条件選択1.Name = "cmb検索条件選択1"
        Me.cmb検索条件選択1.Size = New System.Drawing.Size(13, 20)
        Me.cmb検索条件選択1.TabIndex = 56
        '
        'cmb検索条件選択2
        '
        Me.cmb検索条件選択2.DropDownHeight = 600
        Me.cmb検索条件選択2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb検索条件選択2.DropDownWidth = 250
        Me.cmb検索条件選択2.FormattingEnabled = True
        Me.cmb検索条件選択2.IntegralHeight = False
        Me.cmb検索条件選択2.Location = New System.Drawing.Point(706, 15)
        Me.cmb検索条件選択2.Name = "cmb検索条件選択2"
        Me.cmb検索条件選択2.Size = New System.Drawing.Size(13, 20)
        Me.cmb検索条件選択2.TabIndex = 57
        '
        'cmb検索条件選択3
        '
        Me.cmb検索条件選択3.DropDownHeight = 600
        Me.cmb検索条件選択3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb検索条件選択3.DropDownWidth = 250
        Me.cmb検索条件選択3.FormattingEnabled = True
        Me.cmb検索条件選択3.IntegralHeight = False
        Me.cmb検索条件選択3.Location = New System.Drawing.Point(1048, 16)
        Me.cmb検索条件選択3.Name = "cmb検索条件選択3"
        Me.cmb検索条件選択3.Size = New System.Drawing.Size(13, 20)
        Me.cmb検索条件選択3.TabIndex = 58
        '
        'btn貼り付け
        '
        Me.btn貼り付け.Location = New System.Drawing.Point(893, 602)
        Me.btn貼り付け.Name = "btn貼り付け"
        Me.btn貼り付け.Size = New System.Drawing.Size(75, 23)
        Me.btn貼り付け.TabIndex = 59
        Me.btn貼り付け.Text = "貼り付け"
        Me.btn貼り付け.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(974, 606)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(216, 12)
        Me.Label2.TabIndex = 60
        Me.Label2.Text = "…コンボボックスを除いて複数セル貼り付け可"
        '
        'btnExcel
        '
        Me.btnExcel.Location = New System.Drawing.Point(797, 603)
        Me.btnExcel.Name = "btnExcel"
        Me.btnExcel.Size = New System.Drawing.Size(75, 23)
        Me.btnExcel.TabIndex = 61
        Me.btnExcel.Text = "Excel出力"
        Me.btnExcel.UseVisualStyleBackColor = True
        '
        '製品_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1212, 634)
        Me.Controls.Add(Me.btnExcel)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btn貼り付け)
        Me.Controls.Add(Me.cmb検索条件選択3)
        Me.Controls.Add(Me.cmb検索条件選択2)
        Me.Controls.Add(Me.cmb検索条件選択1)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmb検索項目3)
        Me.Controls.Add(Me.cmb記号3)
        Me.Controls.Add(Me.txt検索条件3)
        Me.Controls.Add(Me.cmb検索項目2)
        Me.Controls.Add(Me.cmb記号2)
        Me.Controls.Add(Me.txt検索条件2)
        Me.Controls.Add(Me.cmb検索項目1)
        Me.Controls.Add(Me.cmb記号1)
        Me.Controls.Add(Me.txt検索条件1)
        Me.Controls.Add(Me.btn製品検索)
        Me.Controls.Add(Me.btn削除)
        Me.Controls.Add(Me.btn製品_見積コピー)
        Me.Controls.Add(Me.btn製品_更新)
        Me.Controls.Add(Me.btn製品to購入)
        Me.Controls.Add(Me.DGV製品)
        Me.Name = "製品_Form"
        Me.Text = "製品リスト"
        CType(Me.DGV製品, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DGV製品 As System.Windows.Forms.DataGridView
    Friend WithEvents btn製品to購入 As System.Windows.Forms.Button
    Friend WithEvents btn製品_更新 As System.Windows.Forms.Button
    Friend WithEvents btn製品_見積コピー As System.Windows.Forms.Button
    Friend WithEvents btn削除 As System.Windows.Forms.Button
    Friend WithEvents btn製品検索 As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmb検索項目3 As System.Windows.Forms.ComboBox
    Friend WithEvents cmb記号3 As System.Windows.Forms.ComboBox
    Friend WithEvents txt検索条件3 As System.Windows.Forms.TextBox
    Friend WithEvents cmb検索項目2 As System.Windows.Forms.ComboBox
    Friend WithEvents cmb記号2 As System.Windows.Forms.ComboBox
    Friend WithEvents txt検索条件2 As System.Windows.Forms.TextBox
    Friend WithEvents cmb検索項目1 As System.Windows.Forms.ComboBox
    Friend WithEvents cmb記号1 As System.Windows.Forms.ComboBox
    Friend WithEvents txt検索条件1 As System.Windows.Forms.TextBox
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents cmb検索条件選択1 As System.Windows.Forms.ComboBox
    Friend WithEvents cmb検索条件選択2 As System.Windows.Forms.ComboBox
    Friend WithEvents cmb検索条件選択3 As System.Windows.Forms.ComboBox
    Friend WithEvents btn貼り付け As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnExcel As System.Windows.Forms.Button
End Class
