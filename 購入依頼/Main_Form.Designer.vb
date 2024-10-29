<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main_Form
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
        Me.DGV購入品入力 = New System.Windows.Forms.DataGridView()
        Me.btn更新 = New System.Windows.Forms.Button()
        Me.btn製品 = New System.Windows.Forms.Button()
        Me.btn見積 = New System.Windows.Forms.Button()
        Me.chk未処理 = New System.Windows.Forms.CheckBox()
        Me.txt検索条件1 = New System.Windows.Forms.TextBox()
        Me.cmb記号1 = New System.Windows.Forms.ComboBox()
        Me.cmb検索項目1 = New System.Windows.Forms.ComboBox()
        Me.btn検索 = New System.Windows.Forms.Button()
        Me.dtp開始日 = New System.Windows.Forms.DateTimePicker()
        Me.cmb検索項目日付 = New System.Windows.Forms.ComboBox()
        Me.dtp終了日 = New System.Windows.Forms.DateTimePicker()
        Me.cmb検索項目2 = New System.Windows.Forms.ComboBox()
        Me.cmb記号2 = New System.Windows.Forms.ComboBox()
        Me.txt検索条件2 = New System.Windows.Forms.TextBox()
        Me.cmb検索項目3 = New System.Windows.Forms.ComboBox()
        Me.cmb記号3 = New System.Windows.Forms.ComboBox()
        Me.txt検索条件3 = New System.Windows.Forms.TextBox()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btn送料 = New System.Windows.Forms.Button()
        Me.btn手数料 = New System.Windows.Forms.Button()
        Me.btn値引き = New System.Windows.Forms.Button()
        Me.btnCOPY = New System.Windows.Forms.Button()
        Me.btn一括承認 = New System.Windows.Forms.Button()
        Me.btn削除 = New System.Windows.Forms.Button()
        Me.btn注文書印刷 = New System.Windows.Forms.Button()
        Me.rbn未承認 = New System.Windows.Forms.RadioButton()
        Me.rbn承認済 = New System.Windows.Forms.RadioButton()
        Me.chk未印刷 = New System.Windows.Forms.CheckBox()
        Me.btn未承認印刷 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnマスタ = New System.Windows.Forms.Button()
        Me.btn技術DB = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblcatalog = New System.Windows.Forms.Label()
        Me.btnExcel = New System.Windows.Forms.Button()
        Me.btn貼り付け = New System.Windows.Forms.Button()
        Me.btn一括稟議承認 = New System.Windows.Forms.Button()
        Me.btn出荷データ作成 = New System.Windows.Forms.Button()
        Me.btnManual = New System.Windows.Forms.Button()
        Me.lblver = New System.Windows.Forms.Label()
        Me.lblセル合計 = New System.Windows.Forms.Label()
        Me.btn検索プログラム = New System.Windows.Forms.Button()
        Me.btnスクロール = New System.Windows.Forms.Button()
        Me.cmb検索条件選択1 = New System.Windows.Forms.ComboBox()
        Me.cmb検索条件選択2 = New System.Windows.Forms.ComboBox()
        Me.cmb検索条件選択3 = New System.Windows.Forms.ComboBox()
        CType(Me.DGV購入品入力, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DGV購入品入力
        '
        Me.DGV購入品入力.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV購入品入力.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DGV購入品入力.Location = New System.Drawing.Point(1, 72)
        Me.DGV購入品入力.Name = "DGV購入品入力"
        Me.DGV購入品入力.RowTemplate.Height = 21
        Me.DGV購入品入力.Size = New System.Drawing.Size(1260, 543)
        Me.DGV購入品入力.TabIndex = 0
        '
        'btn更新
        '
        Me.btn更新.BackColor = System.Drawing.Color.ForestGreen
        Me.btn更新.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.btn更新.Location = New System.Drawing.Point(402, 626)
        Me.btn更新.Name = "btn更新"
        Me.btn更新.Size = New System.Drawing.Size(142, 23)
        Me.btn更新.TabIndex = 1
        Me.btn更新.Text = "更新"
        Me.btn更新.UseVisualStyleBackColor = False
        '
        'btn製品
        '
        Me.btn製品.Location = New System.Drawing.Point(12, 626)
        Me.btn製品.Name = "btn製品"
        Me.btn製品.Size = New System.Drawing.Size(68, 23)
        Me.btn製品.TabIndex = 2
        Me.btn製品.Text = "製品を開く"
        Me.btn製品.UseVisualStyleBackColor = True
        '
        'btn見積
        '
        Me.btn見積.Location = New System.Drawing.Point(86, 626)
        Me.btn見積.Name = "btn見積"
        Me.btn見積.Size = New System.Drawing.Size(68, 23)
        Me.btn見積.TabIndex = 7
        Me.btn見積.Text = "見積を開く"
        Me.btn見積.UseVisualStyleBackColor = True
        '
        'chk未処理
        '
        Me.chk未処理.AutoSize = True
        Me.chk未処理.Location = New System.Drawing.Point(250, 18)
        Me.chk未処理.Name = "chk未処理"
        Me.chk未処理.Size = New System.Drawing.Size(60, 16)
        Me.chk未処理.TabIndex = 9
        Me.chk未処理.Text = "未処理"
        Me.chk未処理.UseVisualStyleBackColor = True
        '
        'txt検索条件1
        '
        Me.txt検索条件1.Location = New System.Drawing.Point(504, 46)
        Me.txt検索条件1.Name = "txt検索条件1"
        Me.txt検索条件1.Size = New System.Drawing.Size(79, 19)
        Me.txt検索条件1.TabIndex = 11
        '
        'cmb記号1
        '
        Me.cmb記号1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb記号1.FormattingEnabled = True
        Me.cmb記号1.Location = New System.Drawing.Point(445, 45)
        Me.cmb記号1.Name = "cmb記号1"
        Me.cmb記号1.Size = New System.Drawing.Size(53, 20)
        Me.cmb記号1.TabIndex = 12
        '
        'cmb検索項目1
        '
        Me.cmb検索項目1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb検索項目1.FormattingEnabled = True
        Me.cmb検索項目1.Location = New System.Drawing.Point(334, 45)
        Me.cmb検索項目1.Name = "cmb検索項目1"
        Me.cmb検索項目1.Size = New System.Drawing.Size(105, 20)
        Me.cmb検索項目1.TabIndex = 13
        '
        'btn検索
        '
        Me.btn検索.BackColor = System.Drawing.SystemColors.Highlight
        Me.btn検索.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.btn検索.Location = New System.Drawing.Point(1091, 15)
        Me.btn検索.Name = "btn検索"
        Me.btn検索.Size = New System.Drawing.Size(77, 24)
        Me.btn検索.TabIndex = 14
        Me.btn検索.Text = "検索"
        Me.btn検索.UseVisualStyleBackColor = False
        '
        'dtp開始日
        '
        Me.dtp開始日.Checked = False
        Me.dtp開始日.Location = New System.Drawing.Point(778, 18)
        Me.dtp開始日.Name = "dtp開始日"
        Me.dtp開始日.ShowCheckBox = True
        Me.dtp開始日.Size = New System.Drawing.Size(139, 19)
        Me.dtp開始日.TabIndex = 15
        '
        'cmb検索項目日付
        '
        Me.cmb検索項目日付.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb検索項目日付.FormattingEnabled = True
        Me.cmb検索項目日付.Location = New System.Drawing.Point(709, 17)
        Me.cmb検索項目日付.Name = "cmb検索項目日付"
        Me.cmb検索項目日付.Size = New System.Drawing.Size(63, 20)
        Me.cmb検索項目日付.TabIndex = 16
        '
        'dtp終了日
        '
        Me.dtp終了日.Checked = False
        Me.dtp終了日.Location = New System.Drawing.Point(946, 18)
        Me.dtp終了日.Name = "dtp終了日"
        Me.dtp終了日.ShowCheckBox = True
        Me.dtp終了日.Size = New System.Drawing.Size(139, 19)
        Me.dtp終了日.TabIndex = 17
        '
        'cmb検索項目2
        '
        Me.cmb検索項目2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb検索項目2.FormattingEnabled = True
        Me.cmb検索項目2.Location = New System.Drawing.Point(651, 45)
        Me.cmb検索項目2.Name = "cmb検索項目2"
        Me.cmb検索項目2.Size = New System.Drawing.Size(105, 20)
        Me.cmb検索項目2.TabIndex = 20
        '
        'cmb記号2
        '
        Me.cmb記号2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb記号2.FormattingEnabled = True
        Me.cmb記号2.Location = New System.Drawing.Point(761, 45)
        Me.cmb記号2.Name = "cmb記号2"
        Me.cmb記号2.Size = New System.Drawing.Size(53, 20)
        Me.cmb記号2.TabIndex = 19
        '
        'txt検索条件2
        '
        Me.txt検索条件2.Location = New System.Drawing.Point(819, 46)
        Me.txt検索条件2.Name = "txt検索条件2"
        Me.txt検索条件2.Size = New System.Drawing.Size(86, 19)
        Me.txt検索条件2.TabIndex = 18
        '
        'cmb検索項目3
        '
        Me.cmb検索項目3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb検索項目3.FormattingEnabled = True
        Me.cmb検索項目3.Location = New System.Drawing.Point(967, 46)
        Me.cmb検索項目3.Name = "cmb検索項目3"
        Me.cmb検索項目3.Size = New System.Drawing.Size(105, 20)
        Me.cmb検索項目3.TabIndex = 23
        '
        'cmb記号3
        '
        Me.cmb記号3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb記号3.FormattingEnabled = True
        Me.cmb記号3.Location = New System.Drawing.Point(1078, 46)
        Me.cmb記号3.Name = "cmb記号3"
        Me.cmb記号3.Size = New System.Drawing.Size(52, 20)
        Me.cmb記号3.TabIndex = 22
        '
        'txt検索条件3
        '
        Me.txt検索条件3.Location = New System.Drawing.Point(1136, 46)
        Me.txt検索条件3.Name = "txt検索条件3"
        Me.txt検索条件3.Size = New System.Drawing.Size(103, 19)
        Me.txt検索条件3.TabIndex = 21
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(1171, 14)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(79, 24)
        Me.btnClear.TabIndex = 24
        Me.btnClear.Text = "クリア"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'btn送料
        '
        Me.btn送料.Location = New System.Drawing.Point(907, 626)
        Me.btn送料.Name = "btn送料"
        Me.btn送料.Size = New System.Drawing.Size(53, 23)
        Me.btn送料.TabIndex = 25
        Me.btn送料.Text = "送料"
        Me.btn送料.UseVisualStyleBackColor = True
        '
        'btn手数料
        '
        Me.btn手数料.Location = New System.Drawing.Point(838, 626)
        Me.btn手数料.Name = "btn手数料"
        Me.btn手数料.Size = New System.Drawing.Size(63, 23)
        Me.btn手数料.TabIndex = 26
        Me.btn手数料.Text = "手数料"
        Me.btn手数料.UseVisualStyleBackColor = True
        '
        'btn値引き
        '
        Me.btn値引き.Location = New System.Drawing.Point(778, 626)
        Me.btn値引き.Name = "btn値引き"
        Me.btn値引き.Size = New System.Drawing.Size(54, 23)
        Me.btn値引き.TabIndex = 27
        Me.btn値引き.Text = "値引き"
        Me.btn値引き.UseVisualStyleBackColor = True
        '
        'btnCOPY
        '
        Me.btnCOPY.Location = New System.Drawing.Point(710, 626)
        Me.btnCOPY.Name = "btnCOPY"
        Me.btnCOPY.Size = New System.Drawing.Size(62, 23)
        Me.btnCOPY.TabIndex = 28
        Me.btnCOPY.Text = "行コピー"
        Me.btnCOPY.UseVisualStyleBackColor = True
        '
        'btn一括承認
        '
        Me.btn一括承認.BackColor = System.Drawing.Color.ForestGreen
        Me.btn一括承認.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.btn一括承認.Location = New System.Drawing.Point(20, 47)
        Me.btn一括承認.Name = "btn一括承認"
        Me.btn一括承認.Size = New System.Drawing.Size(72, 23)
        Me.btn一括承認.TabIndex = 29
        Me.btn一括承認.Text = "一括承認"
        Me.btn一括承認.UseVisualStyleBackColor = False
        '
        'btn削除
        '
        Me.btn削除.BackColor = System.Drawing.Color.LightCoral
        Me.btn削除.Location = New System.Drawing.Point(334, 626)
        Me.btn削除.Name = "btn削除"
        Me.btn削除.Size = New System.Drawing.Size(65, 23)
        Me.btn削除.TabIndex = 30
        Me.btn削除.Text = "削除"
        Me.btn削除.UseVisualStyleBackColor = False
        '
        'btn注文書印刷
        '
        Me.btn注文書印刷.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btn注文書印刷.Location = New System.Drawing.Point(321, 15)
        Me.btn注文書印刷.Name = "btn注文書印刷"
        Me.btn注文書印刷.Size = New System.Drawing.Size(75, 23)
        Me.btn注文書印刷.TabIndex = 31
        Me.btn注文書印刷.Text = "承認済印刷"
        Me.btn注文書印刷.UseVisualStyleBackColor = False
        '
        'rbn未承認
        '
        Me.rbn未承認.AutoSize = True
        Me.rbn未承認.Location = New System.Drawing.Point(24, 18)
        Me.rbn未承認.Name = "rbn未承認"
        Me.rbn未承認.Size = New System.Drawing.Size(59, 16)
        Me.rbn未承認.TabIndex = 32
        Me.rbn未承認.TabStop = True
        Me.rbn未承認.Text = "未承認"
        Me.rbn未承認.UseVisualStyleBackColor = True
        '
        'rbn承認済
        '
        Me.rbn承認済.AutoSize = True
        Me.rbn承認済.Location = New System.Drawing.Point(89, 18)
        Me.rbn承認済.Name = "rbn承認済"
        Me.rbn承認済.Size = New System.Drawing.Size(59, 16)
        Me.rbn承認済.TabIndex = 33
        Me.rbn承認済.TabStop = True
        Me.rbn承認済.Text = "承認済"
        Me.rbn承認済.UseVisualStyleBackColor = True
        '
        'chk未印刷
        '
        Me.chk未印刷.AutoSize = True
        Me.chk未印刷.Location = New System.Drawing.Point(184, 18)
        Me.chk未印刷.Name = "chk未印刷"
        Me.chk未印刷.Size = New System.Drawing.Size(60, 16)
        Me.chk未印刷.TabIndex = 34
        Me.chk未印刷.Text = "未印刷"
        Me.chk未印刷.UseVisualStyleBackColor = True
        '
        'btn未承認印刷
        '
        Me.btn未承認印刷.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btn未承認印刷.Location = New System.Drawing.Point(402, 15)
        Me.btn未承認印刷.Name = "btn未承認印刷"
        Me.btn未承認印刷.Size = New System.Drawing.Size(75, 23)
        Me.btn未承認印刷.TabIndex = 35
        Me.btn未承認印刷.Text = "未承認印刷"
        Me.btn未承認印刷.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(293, 50)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 12)
        Me.Label1.TabIndex = 36
        Me.Label1.Text = "条件1"
        '
        'btnマスタ
        '
        Me.btnマスタ.Location = New System.Drawing.Point(160, 626)
        Me.btnマスタ.Name = "btnマスタ"
        Me.btnマスタ.Size = New System.Drawing.Size(68, 23)
        Me.btnマスタ.TabIndex = 37
        Me.btnマスタ.Text = "マスタを開く"
        Me.btnマスタ.UseVisualStyleBackColor = True
        '
        'btn技術DB
        '
        Me.btn技術DB.Location = New System.Drawing.Point(1079, 626)
        Me.btn技術DB.Name = "btn技術DB"
        Me.btn技術DB.Size = New System.Drawing.Size(89, 23)
        Me.btn技術DB.TabIndex = 38
        Me.btn技術DB.Text = "技術DB取込"
        Me.btn技術DB.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(610, 50)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(35, 12)
        Me.Label3.TabIndex = 40
        Me.Label3.Text = "条件2"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(926, 50)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(35, 12)
        Me.Label4.TabIndex = 41
        Me.Label4.Text = "条件3"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(650, 21)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(53, 12)
        Me.Label5.TabIndex = 43
        Me.Label5.Text = "日付条件"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(923, 21)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(17, 12)
        Me.Label6.TabIndex = 44
        Me.Label6.Text = "～"
        '
        'lblcatalog
        '
        Me.lblcatalog.AutoSize = True
        Me.lblcatalog.ForeColor = System.Drawing.Color.Blue
        Me.lblcatalog.Location = New System.Drawing.Point(12, 655)
        Me.lblcatalog.Name = "lblcatalog"
        Me.lblcatalog.Size = New System.Drawing.Size(54, 12)
        Me.lblcatalog.TabIndex = 45
        Me.lblcatalog.Text = "lblcatalog"
        '
        'btnExcel
        '
        Me.btnExcel.Location = New System.Drawing.Point(1171, 626)
        Me.btnExcel.Name = "btnExcel"
        Me.btnExcel.Size = New System.Drawing.Size(79, 23)
        Me.btnExcel.TabIndex = 46
        Me.btnExcel.Text = "Excel出力"
        Me.btnExcel.UseVisualStyleBackColor = True
        '
        'btn貼り付け
        '
        Me.btn貼り付け.Location = New System.Drawing.Point(253, 626)
        Me.btn貼り付け.Name = "btn貼り付け"
        Me.btn貼り付け.Size = New System.Drawing.Size(75, 23)
        Me.btn貼り付け.TabIndex = 47
        Me.btn貼り付け.Text = "貼り付け"
        Me.btn貼り付け.UseVisualStyleBackColor = True
        '
        'btn一括稟議承認
        '
        Me.btn一括稟議承認.BackColor = System.Drawing.Color.ForestGreen
        Me.btn一括稟議承認.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.btn一括稟議承認.Location = New System.Drawing.Point(98, 47)
        Me.btn一括稟議承認.Name = "btn一括稟議承認"
        Me.btn一括稟議承認.Size = New System.Drawing.Size(101, 23)
        Me.btn一括稟議承認.TabIndex = 48
        Me.btn一括稟議承認.Text = "一括稟議承認"
        Me.btn一括稟議承認.UseVisualStyleBackColor = False
        '
        'btn出荷データ作成
        '
        Me.btn出荷データ作成.Location = New System.Drawing.Point(983, 626)
        Me.btn出荷データ作成.Name = "btn出荷データ作成"
        Me.btn出荷データ作成.Size = New System.Drawing.Size(89, 23)
        Me.btn出荷データ作成.TabIndex = 49
        Me.btn出荷データ作成.Text = "出荷データ作成"
        Me.btn出荷データ作成.UseVisualStyleBackColor = True
        '
        'btnManual
        '
        Me.btnManual.Location = New System.Drawing.Point(500, 16)
        Me.btnManual.Name = "btnManual"
        Me.btnManual.Size = New System.Drawing.Size(63, 23)
        Me.btnManual.TabIndex = 50
        Me.btnManual.Text = "マニュアル"
        Me.btnManual.UseVisualStyleBackColor = True
        '
        'lblver
        '
        Me.lblver.AutoSize = True
        Me.lblver.ForeColor = System.Drawing.Color.Blue
        Me.lblver.Location = New System.Drawing.Point(145, 656)
        Me.lblver.Name = "lblver"
        Me.lblver.Size = New System.Drawing.Size(21, 12)
        Me.lblver.TabIndex = 51
        Me.lblver.Text = "ver"
        '
        'lblセル合計
        '
        Me.lblセル合計.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblセル合計.Location = New System.Drawing.Point(1137, 655)
        Me.lblセル合計.Name = "lblセル合計"
        Me.lblセル合計.Size = New System.Drawing.Size(113, 17)
        Me.lblセル合計.TabIndex = 52
        '
        'btn検索プログラム
        '
        Me.btn検索プログラム.Location = New System.Drawing.Point(569, 16)
        Me.btn検索プログラム.Name = "btn検索プログラム"
        Me.btn検索プログラム.Size = New System.Drawing.Size(75, 23)
        Me.btn検索プログラム.TabIndex = 53
        Me.btn検索プログラム.Text = "検索システム"
        Me.btn検索プログラム.UseVisualStyleBackColor = True
        '
        'btnスクロール
        '
        Me.btnスクロール.Location = New System.Drawing.Point(569, 626)
        Me.btnスクロール.Name = "btnスクロール"
        Me.btnスクロール.Size = New System.Drawing.Size(122, 23)
        Me.btnスクロール.TabIndex = 54
        Me.btnスクロール.Text = "横スクロール"
        Me.btnスクロール.UseVisualStyleBackColor = True
        '
        'cmb検索条件選択1
        '
        Me.cmb検索条件選択1.DropDownHeight = 600
        Me.cmb検索条件選択1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb検索条件選択1.DropDownWidth = 250
        Me.cmb検索条件選択1.FormattingEnabled = True
        Me.cmb検索条件選択1.IntegralHeight = False
        Me.cmb検索条件選択1.Location = New System.Drawing.Point(586, 46)
        Me.cmb検索条件選択1.Name = "cmb検索条件選択1"
        Me.cmb検索条件選択1.Size = New System.Drawing.Size(13, 20)
        Me.cmb検索条件選択1.TabIndex = 55
        '
        'cmb検索条件選択2
        '
        Me.cmb検索条件選択2.DropDownHeight = 600
        Me.cmb検索条件選択2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb検索条件選択2.DropDownWidth = 250
        Me.cmb検索条件選択2.FormattingEnabled = True
        Me.cmb検索条件選択2.IntegralHeight = False
        Me.cmb検索条件選択2.Location = New System.Drawing.Point(907, 46)
        Me.cmb検索条件選択2.Name = "cmb検索条件選択2"
        Me.cmb検索条件選択2.Size = New System.Drawing.Size(13, 20)
        Me.cmb検索条件選択2.TabIndex = 56
        '
        'cmb検索条件選択3
        '
        Me.cmb検索条件選択3.DropDownHeight = 600
        Me.cmb検索条件選択3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb検索条件選択3.DropDownWidth = 250
        Me.cmb検索条件選択3.FormattingEnabled = True
        Me.cmb検索条件選択3.IntegralHeight = False
        Me.cmb検索条件選択3.Location = New System.Drawing.Point(1241, 45)
        Me.cmb検索条件選択3.Name = "cmb検索条件選択3"
        Me.cmb検索条件選択3.Size = New System.Drawing.Size(13, 20)
        Me.cmb検索条件選択3.TabIndex = 57
        '
        'Main_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1263, 676)
        Me.Controls.Add(Me.cmb検索条件選択3)
        Me.Controls.Add(Me.cmb検索条件選択2)
        Me.Controls.Add(Me.cmb検索条件選択1)
        Me.Controls.Add(Me.btnスクロール)
        Me.Controls.Add(Me.btn検索プログラム)
        Me.Controls.Add(Me.lblセル合計)
        Me.Controls.Add(Me.lblver)
        Me.Controls.Add(Me.btnManual)
        Me.Controls.Add(Me.btn出荷データ作成)
        Me.Controls.Add(Me.btn一括稟議承認)
        Me.Controls.Add(Me.btn貼り付け)
        Me.Controls.Add(Me.btnExcel)
        Me.Controls.Add(Me.lblcatalog)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btn技術DB)
        Me.Controls.Add(Me.btnマスタ)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btn未承認印刷)
        Me.Controls.Add(Me.chk未印刷)
        Me.Controls.Add(Me.rbn承認済)
        Me.Controls.Add(Me.rbn未承認)
        Me.Controls.Add(Me.btn注文書印刷)
        Me.Controls.Add(Me.btn削除)
        Me.Controls.Add(Me.btn一括承認)
        Me.Controls.Add(Me.btnCOPY)
        Me.Controls.Add(Me.btn値引き)
        Me.Controls.Add(Me.btn手数料)
        Me.Controls.Add(Me.btn送料)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.cmb検索項目3)
        Me.Controls.Add(Me.cmb記号3)
        Me.Controls.Add(Me.txt検索条件3)
        Me.Controls.Add(Me.cmb検索項目2)
        Me.Controls.Add(Me.cmb記号2)
        Me.Controls.Add(Me.txt検索条件2)
        Me.Controls.Add(Me.dtp終了日)
        Me.Controls.Add(Me.cmb検索項目日付)
        Me.Controls.Add(Me.dtp開始日)
        Me.Controls.Add(Me.btn検索)
        Me.Controls.Add(Me.cmb検索項目1)
        Me.Controls.Add(Me.cmb記号1)
        Me.Controls.Add(Me.txt検索条件1)
        Me.Controls.Add(Me.chk未処理)
        Me.Controls.Add(Me.btn見積)
        Me.Controls.Add(Me.btn製品)
        Me.Controls.Add(Me.btn更新)
        Me.Controls.Add(Me.DGV購入品入力)
        Me.Name = "Main_Form"
        Me.Text = "Form1"
        CType(Me.DGV購入品入力, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DGV購入品入力 As System.Windows.Forms.DataGridView
    Friend WithEvents btn更新 As System.Windows.Forms.Button
    Friend WithEvents btn製品 As System.Windows.Forms.Button
    Friend WithEvents btn見積 As System.Windows.Forms.Button
    Friend WithEvents chk未処理 As System.Windows.Forms.CheckBox
    Friend WithEvents txt検索条件1 As System.Windows.Forms.TextBox
    Friend WithEvents cmb記号1 As System.Windows.Forms.ComboBox
    Friend WithEvents cmb検索項目1 As System.Windows.Forms.ComboBox
    Friend WithEvents btn検索 As System.Windows.Forms.Button
    Friend WithEvents dtp開始日 As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmb検索項目日付 As System.Windows.Forms.ComboBox
    Friend WithEvents dtp終了日 As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmb検索項目2 As System.Windows.Forms.ComboBox
    Friend WithEvents cmb記号2 As System.Windows.Forms.ComboBox
    Friend WithEvents txt検索条件2 As System.Windows.Forms.TextBox
    Friend WithEvents cmb検索項目3 As System.Windows.Forms.ComboBox
    Friend WithEvents cmb記号3 As System.Windows.Forms.ComboBox
    Friend WithEvents txt検索条件3 As System.Windows.Forms.TextBox
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents btn送料 As System.Windows.Forms.Button
    Friend WithEvents btn手数料 As System.Windows.Forms.Button
    Friend WithEvents btn値引き As System.Windows.Forms.Button
    Friend WithEvents btnCOPY As System.Windows.Forms.Button
    Friend WithEvents btn一括承認 As System.Windows.Forms.Button
    Friend WithEvents btn削除 As System.Windows.Forms.Button
    Friend WithEvents btn注文書印刷 As System.Windows.Forms.Button
    Friend WithEvents rbn未承認 As System.Windows.Forms.RadioButton
    Friend WithEvents rbn承認済 As System.Windows.Forms.RadioButton
    Friend WithEvents chk未印刷 As System.Windows.Forms.CheckBox
    Friend WithEvents btn未承認印刷 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnマスタ As System.Windows.Forms.Button
    Friend WithEvents btn技術DB As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblcatalog As System.Windows.Forms.Label
    Friend WithEvents btnExcel As System.Windows.Forms.Button
    Friend WithEvents btn貼り付け As System.Windows.Forms.Button
    Friend WithEvents btn一括稟議承認 As System.Windows.Forms.Button
    Friend WithEvents btn出荷データ作成 As System.Windows.Forms.Button
    Friend WithEvents btnManual As System.Windows.Forms.Button
    Friend WithEvents lblver As System.Windows.Forms.Label
    Friend WithEvents lblセル合計 As System.Windows.Forms.Label
    Friend WithEvents btn検索プログラム As System.Windows.Forms.Button
    Friend WithEvents btnスクロール As System.Windows.Forms.Button
    Friend WithEvents cmb検索条件選択1 As System.Windows.Forms.ComboBox
    Friend WithEvents cmb検索条件選択2 As System.Windows.Forms.ComboBox
    Friend WithEvents cmb検索条件選択3 As System.Windows.Forms.ComboBox

End Class
