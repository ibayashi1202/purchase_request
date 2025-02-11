﻿Imports Microsoft.Office.Interop
Public Class 見積_form

    Inherits System.Windows.Forms.Form
    Private dCon As New merrweth_init_DbConnection
    Public f1 As Main_Form

    Private dataGridViewComboBox As DataGridViewComboBoxEditingControl = Nothing
    Private searchFld As nameValue = New nameValue
    'Private CellEditStart As Boolean = False　使ってなさそうだからコメント図と

    Private Editcell As DataGridViewCell
    Private DT_po As DataTable
    Private DT見積リスト As DataTable
    Public i明細行数 As Integer = 10

    Public i明細START列 As Integer = 1
    Public 変更前index As Integer
    Public 変更前Value As String
    Private SwEditSave As Boolean = False
    Private 見積CloseFlg As Boolean
    Public kc As New 共通処理_Class
    Private DT見積部署別仕入先 As DataTable
    Public Flg印刷対象有 As Boolean = False '印刷対象ないのに印刷完了と出さないためにフラグを作った

    'Private arr見積() As Object
    'コールバック用としてデリゲートを宣言 
    'Public Delegate Sub ShowDelegate(ByVal arrOriginal() As Object, Form As String)


    '引数として受け取ったメソッド保存用 
    'Private CallBack As ShowDelegate

    '━─━─━─━─━─━─━─━─━─━─
    '参照サイト http://www.code-magagine.com/?p=6328
    '■デリゲードとは？
    'メソッドを呼び出す手法の一つで、メソッドを直接呼び出すのではなく、誰かに依頼をしてメソッドを呼び出します。
    'この誰かに当たるのが「デリゲート」になります。
    'デリゲート簡単に言えば「型」です。ただ、中に入るものはメソッド（への参照情報）が入ります。変数同様、メソッドを入れて保持、実行することが出来ます。
    'シグネチャさえ合っていれば呼び出すことができ、クラスをインスタンス化しなくても呼び出すことができるので、クラスを意識することなくメソッドを実行したい場合に使います。

    '■構文
    '定義：　Delegate Sub デリゲード名
    'デリゲードのメソッドの登録：　[アクセス修飾子] 変数名 As New 定義したデリゲート(AddressOf 登録するメソッド名)
    ' …デリゲートはデータ型の一種なので、Asキーワードを使います。

    '■条件
    '引数はデリゲートするメソッドの引数（数、型両方）と同一でなければならない
    '戻り値はデリゲートするメソッドの戻り値と同一でなければならない

    '他のフォームの操作を簡単にするためにはデリゲートが必要だったので使うことにした。
    '━─━─━─━─━─━─━─━─━─━─

    'Main_Formの「見積を開く」ボタンと「製品リスト」を開くボタンを押した時にも通る
    'newキーワードを使ってメソッドの参照先のインスタンスを作ってそれを渡す
    'Public Sub New(ByVal CallBack As ShowDelegate)

    '    ' この呼び出しはデザイナーで必要です。
    '    InitializeComponent()

    '    ' InitializeComponent() 呼び出しの後で初期化を追加します。
    '    Me.CallBack = CallBack

    'End Sub
    '****************************************************
    'Main_Formは画面書き換えたり印刷する前にも保存チェックしているが、製品と見積は保存漏れがあっても、たいした問題にならないから全ボタンに保存確認機能付けなくてよいとのことだったので、閉じる時だけ保存確認している
    '保存機能があるのがいいのかどうか現時点では判断できないので使っていく中で問題があれば追加すればよいらしい
    '→行消えちゃうから、発注申請ボタン押した時は保存確認行うことにした。
    '****************************************************
    Private Sub DGV見積_Read()
        Dim ds見積リスト As DataSet
        DGV見積.DataSource = Nothing

        Dim strSql As String = ""

        strSql = strSql & vbCrLf & "SELECT *"
        'strSql = strSql & vbCrLf & "UID"
        'strSql = strSql & vbCrLf & "部門ID"
        'strSql = strSql & vbCrLf & ",購入者"
        'strSql = strSql & vbCrLf & ",品名"
        'strSql = strSql & vbCrLf & ",型式"
        'strSql = strSql & vbCrLf & ",メーカー"
        'strSql = strSql & vbCrLf & ",数量"
        'strSql = strSql & vbCrLf & ",見積単価"
        'strSql = strSql & vbCrLf & ",仕入先"
        'strSql = strSql & vbCrLf & ",担当者"
        'strSql = strSql & vbCrLf & ",支払区分"
        'strSql = strSql & vbCrLf & ",科目ID"
        'strSql = strSql & vbCrLf & ",見積状態"
        'strSql = strSql & vbCrLf & ",購入理由"

        strSql = strSql & vbCrLf & "FROM TD_見積"
        strSql = strSql & vbCrLf & "WHERE 部署 = '" & Select_Form.s部署 & "'"
        ds見積リスト = dCon.DataSet(strSql, "見積一覧")
        DT見積リスト = ds見積リスト.Tables("見積一覧")
        DGV見積.DataSource = DT見積リスト

    End Sub

    Private Sub DGV見積_詳細設定()
        Dim strSql As String
        Dim Col部門 As New DataGridViewComboBoxColumn
        Dim Col購入者 As New DataGridViewComboBoxColumn
        Dim Col仕入先 As New DataGridViewComboBoxColumn
        Dim Col支払区分 As New DataGridViewComboBoxColumn
        Dim Col科目 As New DataGridViewComboBoxColumn
        Dim col発注状態 As New DataGridViewComboBoxColumn
        Dim Col科目候補 As New DataGridViewTextBoxColumn
        '------------------------------------------------------------------------------
        '部門 
        '------------------------------------------------------------------------------
        'DGVに現在存在しているdivision_id列と今作成したDataGridViewComboBoxColumnを入れ替える
        Col部門.DataPropertyName = DGV見積.Columns("部門ID").DataPropertyName
        Col部門.DataSource = Main_Form.DT部門
        Col部門.ValueMember = "id"
        Col部門.DisplayMember = "division"
        DGV見積.Columns.Insert(DGV見積.Columns("部門ID").Index, Col部門)
        'DGV購入.Columns("division_id").Visible = False
        Col部門.Name = "部門"
        'Col部門.DataPropertyName = "部門"

        '------------------------------------------------------------------------------
        '購入者…自由入力可
        '------------------------------------------------------------------------------
        For Each DR As DataRow In Main_Form.DT部署別購入者.Rows
            Col購入者.Items.Add(DR("employee"))
        Next
        DGV見積.Columns.Insert(DGV見積.Columns("購入者").Index + 1, Col購入者)
        'DGV購入.Columns("employee").Visible = False
        Col購入者.Name = "購入者候補"
        Col購入者.HeaderText = ""
        Col購入者.DropDownWidth = 200
        '------------------------------------------------------------------------------
        '仕入先…自由入力可
        '------------------------------------------------------------------------------
        '2021/04/12 ekawai del S↓------------------
        'For Each DR As DataRow In Main_Form.DT部署別仕入先.Rows
        '    Col仕入先.Items.Add(DR("kiban"))
        'Next
        '2021/04/12 ekawai del E↑------------------
        '2021/04/12 ekawai add S
        strSql = "SELECT"
        strSql = strSql & vbCrLf & "TD_見積.仕入先ID"
        strSql = strSql & vbCrLf & "  ,'×' + TM_facility.kiban AS kiban"
        strSql = strSql & vbCrLf & "  ,TM_paycode.pay_code"
        strSql = strSql & vbCrLf & "  , 2 AS 並び順 "
        strSql = strSql & vbCrLf & "FROM TD_見積 "
        strSql = strSql & vbCrLf & "  INNER JOIN TM_facility "
        strSql = strSql & vbCrLf & "    ON TD_見積.仕入先ID = TM_facility.facility_id "
        strSql = strSql & vbCrLf & "  INNER JOIN TM_paycode ON TM_facility.pay_method = TM_paycode.id"
        strSql = strSql & vbCrLf & "  LEFT OUTER JOIN ( "
        strSql = strSql & vbCrLf & "    SELECT * "
        strSql = strSql & vbCrLf & "    FROM TM_dep_facility "
        strSql = strSql & vbCrLf & "    WHERE"
        strSql = strSql & vbCrLf & "    TM_dep_facility.部署ID = " & Select_Form.i部署ID
        strSql = strSql & vbCrLf & "  ) AS Sub1 "
        strSql = strSql & vbCrLf & "    ON TD_見積.仕入先ID = Sub1.仕入先ID "
        strSql = strSql & vbCrLf & "WHERE"
        strSql = strSql & vbCrLf & "  TD_見積.部署 = " & kc.SQ(Select_Form.s部署)
        strSql = strSql & vbCrLf & "  AND Sub1.仕入先ID IS NULL "
        strSql = strSql & vbCrLf & "UNION "
        strSql = strSql & vbCrLf & "SELECT"
        strSql = strSql & vbCrLf & "  TM_dep_facility.仕入先ID"
        strSql = strSql & vbCrLf & "  , TM_facility.kiban"
        strSql = strSql & vbCrLf & "  , TM_paycode.pay_code"
        strSql = strSql & vbCrLf & "  , 1 AS 並び順 "
        strSql = strSql & vbCrLf & "FROM TM_dep_facility "
        strSql = strSql & vbCrLf & "  INNER JOIN TM_facility"
        strSql = strSql & vbCrLf & "    ON TM_dep_facility.仕入先ID = TM_facility.facility_id"
        strSql = strSql & vbCrLf & "  INNER JOIN TM_paycode ON TM_facility.pay_method = TM_paycode.id"
        strSql = strSql & vbCrLf & "WHERE"
        strSql = strSql & vbCrLf & "  TM_dep_facility.部署ID = " & Select_Form.i部署ID
        strSql = strSql & vbCrLf & "ORDER BY"
        strSql = strSql & vbCrLf & "  並び順"
        strSql = strSql & vbCrLf & "  ,kiban"
        DT見積部署別仕入先 = dCon.DataSet(strSql, "DT").Tables(0)
        '2021/04/12 ekawai add E
        For Each DR As DataRow In DT見積部署別仕入先.Rows
            Col仕入先.Items.Add(DR("kiban"))
        Next

        DGV見積.Columns.Insert(DGV見積.Columns("仕入先").Index + 1, Col仕入先)
        Col仕入先.Name = "仕入先候補"
        Col仕入先.HeaderText = ""
        Col仕入先.DropDownWidth = 200
        '------------------------------------------------------------------------------
        '支払区分
        '------------------------------------------------------------------------------
        '既存の購入依頼は、Excelに書いてある文字(現金、掛け、口座)を表示していただけだったので
        'TM_paycode(1:現金,2:掛け,3:その他)は使わず直接選択肢をAddすることにした。
        'でも仕入先から支払区分選べるようにしたら結局TM_paycode使わないといけないから
        'TM_paycodeのその他⇒口座に変えた
        strSql = "select * from TM_paycode"
        Dim DT支払区分 As DataTable = dCon.DataSet(strSql, "DT").Tables(0)
        'For Each DR As DataRow In DT支払区分.Rows()
        '    Col支払区分.Items.Add(DR("pay_code"))
        'Next

        Col支払区分.DataPropertyName = DGV見積.Columns("支払区分").DataPropertyName
        Col支払区分.DataSource = DT支払区分
        Col支払区分.ValueMember = "pay_code"
        Col支払区分.DisplayMember = "pay_code"
        '同じ列名が2個ある場合Visible = Falseにしたら前のほうだけ非表示になるっぽい
        DGV見積.Columns.Insert(DGV見積.Columns("支払区分").Index + 1, Col支払区分)
        Col支払区分.Name = "支払区分"

        '------------------------------------------------------------------------------
        '科目
        '------------------------------------------------------------------------------
        ''2021/04/12 ekawai add S　main_Formと仕様を統一
        strSql = "SELECT"
        strSql = strSql & vbCrLf & "TD_見積.科目ID"
        strSql = strSql & vbCrLf & ", '×' + CONVERT(nvarchar, TD_見積.科目ID) + ' : ' + TM_kamoku.kamoku AS kamoku"
        strSql = strSql & vbCrLf & ", 2 AS 並び順"
        strSql = strSql & vbCrLf & "FROM TD_見積"
        strSql = strSql & vbCrLf & "INNER JOIN TM_kamoku" '科目マスタ側で変更削除の対策するからLEFT OUTER JOINにしなくていいとのこと
        strSql = strSql & vbCrLf & "ON TD_見積.科目ID = TM_kamoku.id"
        strSql = strSql & vbCrLf & "LEFT OUTER JOIN ( "
        strSql = strSql & vbCrLf & "    SELECT * "
        strSql = strSql & vbCrLf & "    FROM TM_dep_kamoku "
        strSql = strSql & vbCrLf & "    WHERE"
        strSql = strSql & vbCrLf & "      TM_dep_kamoku.部署ID = " & Select_Form.i部署ID
        strSql = strSql & vbCrLf & "  ) AS Sub1 "
        strSql = strSql & vbCrLf & "ON TD_見積.科目ID = Sub1.科目ID "
        strSql = strSql & vbCrLf & "WHERE"
        strSql = strSql & vbCrLf & " TD_見積.部署 = '" & Select_Form.s部署 & "'"
        strSql = strSql & vbCrLf & "  AND Sub1.科目ID IS NULL "
        strSql = strSql & vbCrLf & "UNION "
        strSql = strSql & vbCrLf & "SELECT"
        strSql = strSql & vbCrLf & "  TM_dep_kamoku.科目ID"
        strSql = strSql & vbCrLf & "   , CONVERT(nvarchar,TM_dep_kamoku.科目ID) + ' : ' +TM_kamoku.kamoku as kamoku"
        strSql = strSql & vbCrLf & "  , 1 AS 並び順 "
        strSql = strSql & vbCrLf & "FROM TM_dep_kamoku "
        strSql = strSql & vbCrLf & "  INNER JOIN TM_kamoku "
        strSql = strSql & vbCrLf & "    ON TM_dep_kamoku.科目ID = TM_kamoku.id "
        strSql = strSql & vbCrLf & "WHERE"
        strSql = strSql & vbCrLf & "  TM_dep_kamoku.部署ID = " & Select_Form.i部署ID
        strSql = strSql & vbCrLf & "ORDER BY"
        strSql = strSql & vbCrLf & "  並び順"
        strSql = strSql & vbCrLf & "  , 科目ID"
        Dim DT部署別見積科目 As DataTable

        DT部署別見積科目 = dCon.DataSet(strSql, "DT").Tables(0)
        '2021/04/12 ekawai add E

        'DGVに現在存在しているdivision_id列と今作成したDataGridViewComboBoxColumnを入れ替える
        Col科目.DataPropertyName = DGV見積.Columns("科目ID").DataPropertyName
        '2021/04/12 ekawai del S↓------------------
        'Col科目.ValueMember = "kamoku_id" 
        'Col科目.DataSource = Main_Form.DT部署別科目 
        '2021/04/12 ekawai del E↑------------------
        '2021/04/12 ekawai add S
        Col科目.ValueMember = "科目ID"　
        Col科目.DisplayMember = "kamoku"    
        Col科目.DataSource = DT部署別見積科目 
        '2021/04/12 ekawai add E
		
        DGV見積.Columns.Insert(DGV見積.Columns("科目ID").Index, Col科目)
        DGV見積.Columns("科目ID").Visible = False
        Col科目.Name = "科目"
        Col科目.DropDownWidth = 150

        '製造日報のようにテキストボックスに入れた値で科目ID検索できるようにという指示があったため列を追加した
        DGV見積.Columns.Insert(DGV見積.Columns("科目").Index, Col科目候補)
        Col科目候補.Name = "科目番号"
        '------------------------------------------------------------------------------
        '発注状態
        '------------------------------------------------------------------------------
        col発注状態.Items.Add("保留")
        col発注状態.Items.Add("見積待ち")
        col発注状態.Items.Add("見積希望")
        col発注状態.Items.Add("発注希望")


        DGV見積.Columns.Insert(DGV見積.Columns("見積状態").Index, col発注状態)
        col発注状態.DataPropertyName = DGV見積.Columns("見積状態").DataPropertyName
        DGV見積.Columns("見積状態").Visible = False
        col発注状態.Name = "見積状態"


        '非表示
        DGV見積.Columns("UID").Visible = False
        DGV見積.Columns("仕入先ID").Visible = False
        DGV見積.Columns("支払区分").Visible = False
        DGV見積.Columns("部署").Visible = False
        DGV見積.Columns("部門ID").Visible = False

        '幅
        DGV見積.Columns("購入者候補").Width = 20
        DGV見積.Columns("購入者").Width = 100
        DGV見積.Columns("品名").Width = 150
        DGV見積.Columns("型式").Width = 150
        DGV見積.Columns("メーカー").Width = 100
        DGV見積.Columns("数量").Width = 60
        DGV見積.Columns("見積単価").Width = 60
        DGV見積.Columns("仕入先候補").Width = 20
        DGV見積.Columns("仕入先").Width = 100
        'DGV見積.Columns("支払区分").Width = 40
        DGV見積.Columns("科目").Width = 100
        DGV見積.Columns("科目番号").Width = 60
        DGV見積.Columns("見積状態").Width = 100
        DGV見積.Columns("購入理由").Width = 150

        DirectCast(DGV見積.Columns("購入者"), DataGridViewTextBoxColumn).MaxInputLength = 40
        DirectCast(DGV見積.Columns("品名"), DataGridViewTextBoxColumn).MaxInputLength = 125
        DirectCast(DGV見積.Columns("型式"), DataGridViewTextBoxColumn).MaxInputLength = 125
        DirectCast(DGV見積.Columns("メーカー"), DataGridViewTextBoxColumn).MaxInputLength = 125
        DirectCast(DGV見積.Columns("仕入先"), DataGridViewTextBoxColumn).MaxInputLength = 50
        DirectCast(DGV見積.Columns("購入理由"), DataGridViewTextBoxColumn).MaxInputLength = 120

    End Sub

    Private Sub 見積_form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '右上の閉じるボタンから閉じられるとき
        If e.CloseReason = CloseReason.UserClosing Then

            If SwEditSave Then
                'If MsgBox("保存せずに終了しますか？", vbYesNo + vbDefaultButton1) = vbNo Then
                'Main_Formは画面書き換えたり印刷する前にも保存チェックしているが、製品と見積はで保存漏れがあっても、たいした問題にならないから全ボタンに保存確認機能付けなくてよいとのことだったので、閉じる時だけ保存確認している
                '保存機能があるのがいいのかどうか現時点では判断できないので使っていく中で問題があれば追加すればよいらしい
                If MsgBox("保存してよろしいですか？" & vbCrLf & "「いいえ」を押すと更新していないデータは消えます", vbYesNo + vbDefaultButton1) = vbYes Then
                    'If 見積更新() = False Then
                    '    '更新失敗したら閉じない
                    '    e.Cancel = True
                    'End If
                    btn見積_更新.PerformClick()
                    If 見積CloseFlg = False Then
                        e.Cancel = True
                    End If
                Else
                    SwEditSave = False

                End If
            End If
        End If
    End Sub

    Private Sub 見積_form_Load(sender As Object, e As EventArgs) Handles Me.Load

        Me.Text = Select_Form.s部署
        DGV見積.MultiSelect = False '複数選択禁止
        DGV見積.RowHeadersWidth = 20

        DGV見積_Read()
        DGV見積_詳細設定()
        DGV見積.AutoGenerateColumns = False '列の自動生成禁止 これをしないと列の並び順が変わってしまう
        DGV見積.AllowUserToDeleteRows = False 'Deleteボタンによる行削除禁止(DGVにdatasourceセットする前にFalseにしたらエラーになる)
        DGV見積.MultiSelect = False '複数選択禁止
        If Select_Form.更新Flg = False And Select_Form.承認Flg = False And Select_Form.経理Flg = False Then
            btn見積_更新.Visible = False
            btn見積to製品.Visible = False
            btn見積印刷.Visible = False
            btn発注申請.Visible = False
        End If
        'グリッドの並び替えを禁止する。(ただしプログラムからの並び替えは可能とする)
        For Each c As DataGridViewColumn In DGV見積.Columns
            c.SortMode = DataGridViewColumnSortMode.Programmatic
        Next c
        'Cell_Enterでコンボが開く設定にしているので、Form_Load時はコンボ以外のセルを選択するようにする(デフォルトだとコンボが選択されてしまう)
        DGV見積.CurrentCell = DGV見積.Rows(0).Cells("購入者")
    End Sub

    Private Sub btn見積to製品_Click(sender As Object, e As EventArgs) Handles btn見積to製品.Click
        Dim strSql As String
        Dim i選択行No As Integer = DGV見積.SelectedCells(0).RowIndex
        Dim dt As DataTable = DGV見積.DataSource
        Dim dv As DataRowVersion
        Dim iエラー行 As Integer
        dv = DataRowVersion.Current

        If DGV見積.Rows.Count = i選択行No + 1 Then
            MsgBox("この行はコピーできません")
        Else
            Dim row As DataRow = dt.Rows(i選択行No)

            Dim tran As System.Data.SqlClient.SqlTransaction = Nothing

            iエラー行 = Main_Form.必須項目Check(dt.Rows.IndexOf(row) + 1, row, dv, "見積", "製品")
            If iエラー行 > 0 Then
                'エラー行にスクロール
                DGV見積.FirstDisplayedScrollingRowIndex = iエラー行 - 1
                DGV見積.Rows(iエラー行 - 1).Selected = True

                Exit Sub
            End If


            strSql = "INSERT INTO"
            strSql &= vbCrLf & "TM_製品("
            strSql &= vbCrLf & "部署"
            strSql &= vbCrLf & ",部門ID"
            strSql &= vbCrLf & ",購入者"
            strSql &= vbCrLf & ",品名"
            strSql &= vbCrLf & ",型式"
            strSql &= vbCrLf & ",数量"
            strSql &= vbCrLf & ",予算単価"
            strSql &= vbCrLf & ",仕入先ID"
            strSql &= vbCrLf & ",仕入先"
            strSql &= vbCrLf & ",支払区分"
            strSql &= vbCrLf & ",科目ID"
            strSql &= vbCrLf & ",購入理由"
            strSql &= vbCrLf & ") VALUES ("
            strSql &= vbCrLf & kc.SQ(Select_Form.s部署)
            strSql &= vbCrLf & "," & kc.nn(row("部門ID", dv))
            strSql &= vbCrLf & "," & kc.SQ(row("購入者", dv)) '製品では任意だが見積の必須項目
            strSql &= vbCrLf & "," & kc.SQ(row("品名", dv)) '必須
            strSql &= vbCrLf & "," & kc.nn(row("型式", dv))
            strSql &= vbCrLf & "," & kc.nz(row("数量", dv))
            strSql &= vbCrLf & "," & kc.nz(row("見積単価", dv))
            strSql &= vbCrLf & "," & kc.nn(row("仕入先ID", dv))
            strSql &= vbCrLf & "," & kc.SQ(row("仕入先", dv)) '必須
            strSql &= vbCrLf & "," & kc.SQ(row("支払区分", dv)) '必須
            strSql &= vbCrLf & "," & row("科目ID", dv) '必須
            strSql &= vbCrLf & "," & kc.nn(row("購入理由", dv))
            strSql &= vbCrLf & ")"

            tran = dCon.Connection.BeginTransaction
            If dCon.ExecuteSqlMW(tran, strSql) = False Then
                tran.Rollback()
                tran.Dispose()
                MsgBox("TM_製品への書き込みに失敗しました")
                Exit Sub
            End If
            tran.Commit()

            MsgBox("製品にコピーしました")
        End If
        DGV見積_Read() '見積の保存機能はボタン閉じる時だけでいいそうなので、このボタン押す前に更新ボタン押さずに変更した内容あったら消えるが、消えても別にたいした問題にならないから気にしなくていいらしい。

    End Sub
    '見積から購入品入力
    Private Sub btn発注申請_Click(sender As Object, e As EventArgs) Handles btn発注申請.Click
        Dim strSql As String
        'ReDim arr見積(11)
        'Dim Result As Boolean
        Dim dt As DataTable = DGV見積.DataSource
        Dim dv As DataRowVersion
        Dim iエラー行 As Integer
        Dim tran As System.Data.SqlClient.SqlTransaction = Nothing
        Dim row As DataRow
        Dim cnt As Integer

        'メインフォームが開いているかチェック
        If My.Application.OpenForms("Main_Form") IsNot Nothing Then

            For Each row In dt.Rows
                dv = DataRowVersion.Current
                If row("見積状態", dv) = "発注希望" Then
                    iエラー行 = Main_Form.必須項目Check(dt.Rows.IndexOf(row) + 1, row, dv, "見積", "購入品入力")
                    If iエラー行 > 0 Then
                        'エラー行にスクロール
                        DGV見積.FirstDisplayedScrollingRowIndex = iエラー行 - 1
                        DGV見積.Rows(iエラー行 - 1).Selected = True
                        'エラーが1つでもあれば値転送実施しない
                        Exit Sub
                    End If

                End If

            Next


            If SwEditSave Then
                If MsgBox("保存してよろしいですか?" & vbCrLf & "「いいえ」を押すと更新していないデータは消えます", vbYesNo + vbDefaultButton1) = vbYes Then
                    If 見積更新() = False Then
                        '更新失敗したら画面を書き換えずに終了
                        '発注希望の行以外でも更新時にエラーがあったら発注申請できない(特にこの仕様にしている根拠はないので問題があれば変更しても良いとのこと)
                        MsgBox("更新失敗。発注申請を中止します。")
                        Exit Sub
                    End If
                Else
                    '本来は更新成功しないとSqEditSave = Trueにならないが、この場合は強制的にFalseにする(そうしないと更新するまで画面切り替える度、永遠に更新しますか？というメッセージが出てしまうため)
                    SwEditSave = False
                End If
            End If

            cnt = 0

            For Each row In dt.Rows
                dv = DataRowVersion.Current
                If row("見積状態", dv) = "発注希望" Then

                    strSql = "INSERT INTO"
                    strSql &= vbCrLf & "TD_po("
                    strSql &= vbCrLf & "group_name"
                    strSql &= vbCrLf & ",employee"
                    strSql &= vbCrLf & ",part_name"
                    strSql &= vbCrLf & ",part_number"
                    strSql &= vbCrLf & ",number"
                    strSql &= vbCrLf & ",予算単価"
                    strSql &= vbCrLf & ",vendor_id"
                    strSql &= vbCrLf & ",vendor"
                    strSql &= vbCrLf & ",pay_code"
                    strSql &= vbCrLf & ",kamoku_id"
                    strSql &= vbCrLf & ",remark"
                    strSql &= vbCrLf & ",division_id"
                    strSql &= vbCrLf & ",購入ID"
                    strSql &= vbCrLf & ",insertStamp"
                    strSql &= vbCrLf & ") VALUES ("
                    strSql &= vbCrLf & kc.SQ(Select_Form.s部署)
                    strSql &= vbCrLf & "," & kc.SQ(row("購入者", dv)) '必須
                    strSql &= vbCrLf & "," & kc.SQ(row("品名", dv)) '必須
                    strSql &= vbCrLf & "," & kc.nn(row("型式", dv)).Replace(Chr(13), "").Replace(Chr(10), "")
                    strSql &= vbCrLf & "," & row("数量", dv) '必須　…事前チェックしているから0はあり得ない
                    strSql &= vbCrLf & "," & row("見積単価", dv) '必須　…事前チェックしているから0はあり得ない
                    strSql &= vbCrLf & "," & kc.nn(row("仕入先ID", dv))
                    strSql &= vbCrLf & "," & kc.SQ(row("仕入先", dv)) '必須
                    strSql &= vbCrLf & "," & kc.SQ(row("支払区分", dv)) '必須
                    strSql &= vbCrLf & "," & row("科目ID", dv) '必須(int型)
                    strSql &= vbCrLf & "," & kc.nn(row("購入理由", dv))
                    strSql &= vbCrLf & "," & kc.SQ(row("部門ID", dv)) '必須
                    strSql &= vbCrLf & "," & kc.SQ(Main_Form.getNext購入ID())
                    strSql &= vbCrLf & "," & kc.SQ(Now)
                    strSql &= vbCrLf & ")"
                    'データベース接続、トランザクション開始（行ロックする）
                    tran = dCon.Connection.BeginTransaction
                    If dCon.ExecuteSqlMW(tran, strSql) = False Then
                        tran.Rollback()
                        MsgBox(dt.Rows.IndexOf(row) + 1 & "行目 TD_poへの書き込みに失敗しました")
                        Exit Sub
                    Else

                    End If

                    strSql = "delete from TD_見積"
                    strSql &= vbCrLf & "where UID = " & row("UID", dv)
                    If dCon.ExecuteSqlMW(tran, strSql) = False Then
                        tran.Rollback()
                        MsgBox(dt.Rows.IndexOf(row) + 1 & "行目 TD_見積の削除に失敗しました")
                        'Exit Sub　1行ずつのループなのでそのまま次の行に進む
                    Else
                        tran.Commit()
                        tran.Dispose()
                        cnt = cnt + 1
                    End If

                End If

            Next


            'My.Application.OpenForms("Main_Form").WindowState = FormWindowState.Normal
            'My.Application.OpenForms("Main_Form").Activate()

        End If
        DGV見積_Read()

        If cnt > 0 Then
            MsgBox("申請した内容がコピーされました。" & vbCrLf & "メイン画面を一度閉じてから開き直すと確認できます")
        End If

        'Result = 見積更新()
        '更新成功したら現在のDGVの状態でDBに書きこむ
    End Sub

    Private Sub btn見積_更新_Click(sender As Object, e As EventArgs) Handles btn見積_更新.Click
        If 見積更新() Then
            見積CloseFlg = True
            DGV見積_Read()
        Else
            見積CloseFlg = False
            MsgBox("更新失敗")
        End If
    End Sub

    Private Function 見積更新() As Boolean
        Dim 更新対象Flg As String
        Dim strSql As String = ""
        Dim dv As DataRowVersion
        Dim iエラー行 As Integer
        Dim dt As DataTable = DGV見積.DataSource
        Dim tran As System.Data.SqlClient.SqlTransaction = Nothing


        更新対象Flg = False
        If dt.Rows.Count = DGV見積.Rows.Count Then
            '最終行の仕入先をクリアしたあとにコンボ選び直した時とかに、空行がdtの行としてカウントされてしまうことがある。空白行だから絶対エラーチェックにひっかかる。
            'だから、更新チェックの対象にならないように削除する
            dt.Rows(dt.Rows.Count - 1).Delete()
        End If

        If dt.Rows.Count > 0 Then
            '更新前のエラーチェック
            For Each row As DataRow In dt.Rows
                dv = DataRowVersion.Current
                Select Case row.RowState
                    Case DataRowState.Unchanged
                        '変更なし⇒次へ進む 
                        Continue For
                    Case DataRowState.Modified, DataRowState.Added
                        '更新 or 追加　⇒エラーチェック
                        iエラー行 = Main_Form.必須項目Check(dt.Rows.IndexOf(row) + 1, row, dv, "見積", "")
                        If iエラー行 > 0 Then
                            'エラー行にスクロール
                            DGV見積.FirstDisplayedScrollingRowIndex = iエラー行 - 1
                            DGV見積.Rows(iエラー行 - 1).Selected = True
                            'TD_poへの書き込みを中止する
                            Return False
                        End If
                End Select

            Next

            tran = dCon.Connection.BeginTransaction
            For Each row As DataRow In dt.Rows

                dv = DataRowVersion.Current
                Try
                    Select Case row.RowState
                        Case DataRowState.Unchanged
                            '次へ進む 
                            Continue For
                        Case DataRowState.Modified
                            ''更新の場合
                            ''更新前のチェック

                            ''Check_DGV見積(row, dv)
                            'strSql = "UPDATE TD_見積"
                            'strSql &= vbCrLf & "SET"
                            'strSql &= vbCrLf & "部署 = " & Main_Form.SQ(Select_Form.s部署)
                            ''品名：★必須
                            'strSql &= vbCrLf & ",品名 = " & Main_Form.SQ(row("品名", dv))
                            'If row("部門ID", dv) IsNot DBNull.Value Then
                            '    strSql &= vbCrLf & ", 部門ID = " & Main_Form.SQ(row("部門ID", dv))
                            'End If
                            ''購入者：★必須
                            'strSql &= vbCrLf & ", 購入者 = " & Main_Form.SQ(row("購入者", dv))
                            ''型式：空白可
                            ''改行コードを取る 旧まとめブックの仕様を継承
                            'If row("型式", dv) IsNot DBNull.Value Then
                            '    strSql &= vbCrLf & ", 型式 = " & Main_Form.SQ(row("型式", dv)).Replace(Chr(13), "").Replace(Chr(10), "")
                            'End If

                            ''見積単価：空白可
                            'If row("見積単価", dv) IsNot DBNull.Value Then
                            '    strSql &= vbCrLf & ", 見積単価 = " & Main_Form.nz(row("見積単価", dv))
                            'End If

                            ''数量：空白許可　NULLなら0
                            'strSql &= vbCrLf & ", 数量 = " & Main_Form.nz(row("数量", dv))

                            ''科目ID：★必須
                            'strSql &= vbCrLf & ", 科目ID = " & row("科目ID", dv)

                            ''仕入先ID：空白許可
                            'If row("仕入先ID", dv) IsNot DBNull.Value Then
                            '    strSql &= vbCrLf & ", 仕入先ID = " & Main_Form.SQ(row("仕入先ID", dv))
                            'End If

                            ''仕入先：★必須
                            'strSql &= vbCrLf & ", 仕入先 = " & Main_Form.SQ(row("仕入先", dv))

                            ''支払区分：★必須
                            'strSql &= vbCrLf & ", 支払区分 = " & Main_Form.SQ(row("支払区分", dv))

                            ''購入理由
                            'If row("購入理由", dv) IsNot DBNull.Value Then
                            '    strSql &= vbCrLf & ", 購入理由 = " & Main_Form.SQ(row("購入理由", dv))
                            'End If

                            ''メーカー
                            'If row("メーカー", dv) IsNot DBNull.Value Then
                            '    strSql &= vbCrLf & ", メーカー = " & Main_Form.SQ(row("メーカー", dv))
                            'End If


                            ''見積状態
                            'If row("見積状態", dv) IsNot DBNull.Value Then
                            '    strSql &= vbCrLf & ", 見積状態 = " & Main_Form.SQ(row("見積状態", dv))
                            'End If

                            '既存行の変更の場合
                            '文字列型の列はNULLの場合は空白で更新する

                            strSql = "UPDATE TD_見積"
                            strSql &= vbCrLf & "SET"
                            '部署：必ず入る
                            strSql &= vbCrLf & "部署 = " & kc.SQ(Select_Form.s部署)

                            '購入者：★必須
                            strSql &= vbCrLf & ", 購入者 = " & kc.SQ(row("購入者", dv))

                            '品名：★必須
                            strSql &= vbCrLf & ",品名 = " & kc.SQ(row("品名", dv))

                            '型式：空白可
                            '改行コードを取る →　改行取るのやめた
                            'If row("型式", dv) IsNot DBNull.Value Then
                                'strSql &= vbCrLf & ", 型式 = " & kc.SQ(row("型式", dv)).Replace(Chr(13), "").Replace(Chr(10), "")
                            'Else
                                'strSql &= vbCrLf & ", 型式 = ''"
                            'End If

                            '見積単価、数量：空白可　NULLなら0
                            strSql &= vbCrLf & ", 見積単価 = " & kc.nz(row("見積単価", dv))
                            strSql &= vbCrLf & ", 数量 = " & kc.nz(row("数量", dv))

                            '科目ID NULLの時はNULLにする(int型)
                            strSql &= vbCrLf & ", 科目ID = " & kc.nn(row("科目ID", dv))

                            '仕入先：★必須
                            strSql &= vbCrLf & ", 仕入先 = " & kc.SQ(row("仕入先", dv))

                            '仕入先、支払区分、購入理由、メーカー、見積状態、部門ID、型式…NULLだったらNULLにしてUPDATE
                            strSql &= vbCrLf & ", 仕入先ID = " & kc.nn(row("仕入先ID", dv))
                            strSql &= vbCrLf & ", 支払区分 = " & kc.nn(row("支払区分", dv))
                            strSql &= vbCrLf & ", 購入理由 = " & kc.nn(row("購入理由", dv))
                            strSql &= vbCrLf & ", メーカー = " & kc.nn(row("メーカー", dv))
                            strSql &= vbCrLf & ", 見積状態 = " & kc.nn(row("見積状態", dv))
                            strSql &= vbCrLf & ", 部門ID = " & kc.nn(row("部門ID", dv))
                            strSql &= vbCrLf & ", 型式 = " & kc.nn(row("型式", dv))
                            strSql &= vbCrLf & "WHERE UID = " & row("UID", dv)

                            'tran = dCon.Connection.BeginTransaction
                            If dCon.ExecuteSqlMW(tran, strSql) = False Then
                                tran.Rollback()
                                'エラー行にスクロール
                                DGV見積.FirstDisplayedScrollingRowIndex = dt.Rows.IndexOf(row)
                                DGV見積.Rows(dt.Rows.IndexOf(row)).Selected = True

                                MsgBox(dt.Rows.IndexOf(row) + 1 & "行目 UPDATE失敗")
                                Return False
                            Else
                                更新対象Flg = True
                            End If

                        Case DataRowState.Added

                            '新規行の場合


                            '追加前のチェック()

                            'Check_DGV見積(row, dv)


                            strSql = "INSERT INTO"
                            strSql &= vbCrLf & "TD_見積("
                            strSql &= vbCrLf & "部署"
                            strSql &= vbCrLf & ",品名"
                            strSql &= vbCrLf & ",部門ID"
                            strSql &= vbCrLf & ",購入者"
                            strSql &= vbCrLf & ",型式"
                            strSql &= vbCrLf & ",見積単価"
                            strSql &= vbCrLf & ",数量"
                            strSql &= vbCrLf & ",科目ID"
                            strSql &= vbCrLf & ",仕入先ID"
                            strSql &= vbCrLf & ",仕入先"
                            strSql &= vbCrLf & ",支払区分"
                            strSql &= vbCrLf & ",購入理由"
                            strSql &= vbCrLf & ",メーカー"
                            strSql &= vbCrLf & ",見積状態"
                            strSql &= vbCrLf & ") VALUES ("
                            strSql &= vbCrLf & kc.SQ(Select_Form.s部署)
                            strSql &= vbCrLf & "," & kc.SQ(row("品名", dv)) '必須
                            strSql &= vbCrLf & "," & kc.nn(row("部門ID", dv))
                            strSql &= vbCrLf & "," & kc.SQ(row("購入者", dv)) '必須
                            strSql &= vbCrLf & "," & kc.nn(row("型式", dv)).Replace(Chr(13), "").Replace(Chr(10), "")
                            strSql &= vbCrLf & "," & kc.nz(row("見積単価", dv))
                            strSql &= vbCrLf & "," & kc.nz(row("数量", dv))
                            strSql &= vbCrLf & "," & kc.nn(row("科目ID", dv))
                            strSql &= vbCrLf & "," & kc.nn(row("仕入先ID", dv))
                            strSql &= vbCrLf & "," & kc.SQ(row("仕入先", dv)) '必須
                            strSql &= vbCrLf & "," & kc.nn(row("支払区分", dv))
                            strSql &= vbCrLf & "," & kc.nn(row("購入理由", dv))
                            strSql &= vbCrLf & "," & kc.nn(row("メーカー", dv))
                            'strSql &= vbCrLf & "," & nr(row("担当者", dv))
                            strSql &= vbCrLf & "," & kc.nn(row("見積状態", dv))
                            strSql &= vbCrLf & ")"

                            'tran = dCon.Connection.BeginTransaction
                            If dCon.ExecuteSqlMW(tran, strSql) = False Then
                                tran.Rollback()
                                DGV見積.FirstDisplayedScrollingRowIndex = dt.Rows.IndexOf(row)
                                DGV見積.Rows(dt.Rows.IndexOf(row)).Selected = True
                                MsgBox(dt.Rows.IndexOf(row) + 1 & "行目 INSERT失敗")
                                Return False
                            End If
                            更新対象Flg = True

                            '削除行
                            '2022/02/08 btn削除_Clickに移動
                            'Case DataRowState.Deleted
                            '    更新対象Flg = True
                            '    'TD_見積にidがあれば削除する
                            '    'Deleted 行には Current 行バージョンがないため、列値にアクセスするときに DataRowVersion.Original を渡す必要があります。
                            '    If Main_Form.nz(row("UID", DataRowVersion.Original)) <> 0 Then
                            '        strSql = "DELETE FROM TD_見積"
                            '        strSql &= vbCrLf & "WHERE UID = " & row("UID", DataRowVersion.Original)

                            '        If dCon.ExecuteSqlMW(tran, strSql) = False Then
                            '            tran.Rollback()
                            '            MsgBox("UID =" & row("UID", DataRowVersion.Original) & " 削除失敗")
                            '            Return False
                            '        End If


                            '    Else
                            '        'なければ画面更新したら勝手に消えるから何もしない
                            '    End If

                    End Select
                Catch ex As Exception
                    MsgBox(ex.Message & "更新失敗")
                    If tran IsNot Nothing Then
                        tran.Rollback()
                    End If
                    Return False
                End Try
            Next

            If 更新対象Flg = True Then
                tran.Commit()
                dt.AcceptChanges() 'Rowstateをunchangedに更新する
                MsgBox("更新成功")

            Else
                tran.Rollback()
                MsgBox("更新対象がありません")
                tran.Dispose()
                '更新対象ない場合も失敗したわけではないからTRUEを返す
            End If

            SwEditSave = False
            tran.Dispose()
            Return True
        End If
    End Function



    'Private Sub Check_DGV見積(r As DataRow, dv As DataRowVersion)


    '    If Main_Form.ns(r("品名", dv)) = "" Then
    '        MsgBox("品名を入力してください")
    '        Throw New Exception
    '    End If

    '    If Main_Form.ns(r("部門ID", dv)) = "" Then
    '        MsgBox("部門を入力してください")
    '        Throw New Exception
    '    End If

    '    If Main_Form.ns(r("購入者", dv)) = "" Then
    '        MsgBox("購入者を入力してください")
    '        Throw New Exception
    '    End If
    '    '見積来るまでは入力できないので空白可
    '    'If nz(r("見積単価", dv)) = 0 Then
    '    '    MsgBox("見積単価を入力してください")
    '    '    Throw New Exception
    '    'End If

    '    If Main_Form.nz(r("数量", dv)) = 0 Then
    '        MsgBox("数量を入力してください")
    '        Throw New Exception
    '    End If

    '    If Main_Form.nz(r("科目ID", dv)) = 0 Then
    '        MsgBox("科目IDを入力してください")
    '        Throw New Exception
    '    End If

    '    If Main_Form.ns(r("仕入先", dv)) = "" Then
    '        MsgBox("仕入先を入力してください")
    '        Throw New Exception
    '    End If

    '    If Main_Form.ns(r("支払区分", dv)) = "" Then
    '        MsgBox("支払区分を入力してください")
    '        Throw New Exception
    '    End If

    'End Sub
    Private Sub dataGridViewComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim s選択値 As String
        Dim i As Integer

        'オブジェクト型からDataGridViewComboBoxEditingControlに変換
        Dim cb As DataGridViewComboBoxEditingControl = _
            CType(sender, DataGridViewComboBoxEditingControl)
        If cb.SelectedIndex = -1 Then
            Exit Sub
        End If
        s選択値 = cb.SelectedItem.ToString
        Debug.Print(s選択値)
        With DGV見積
            i = .CurrentCell.RowIndex
            Select Case .CurrentCell.OwningColumn.Name
                Case "仕入先候補"
                    '2021/04/12 ekawai add S
                    コンボイベントハンドラ削除()
                    If s選択値.Substring(0, 1) = "×" Then
                        MsgBox("現在使われていない項目です")
                        cb.SelectedIndex = 変更前index
                    Else
                        .Rows(i).Cells("仕入先").Value = s選択値
                        GET仕入先ID(s選択値, i)
                    End If
                    コンボイベントハンドラ追加()
                    '2021/04/12 ekawai add E
                Case "購入者候補"
                    コンボイベントハンドラ削除() '2021/04/12 ekawai add 
                    .Rows(i).Cells("購入者").Value = s選択値
                    コンボイベントハンドラ追加() '2021/04/12 ekawai add 

                Case "科目", "部門"
                    'カレントセルを取得する時こうやって指定しないとコンボの背景が黒くなるため
                    Dim sTarget As String = cb.EditingControlFormattedValue
                    変更前Value = sTarget
                    If sTarget.Substring(0, 1) = "×" Then
                        MsgBox("現在使われていない項目です")
                        コンボイベントハンドラ削除() '2021/04/12 ekawai add 
                        cb.SelectedIndex = 変更前index
                        コンボイベントハンドラ追加() '2021/04/12 ekawai add 
                    End If


            End Select
            '2021/04/12 ekawai del S↓------------------
            '    RemoveHandler Me.dataGridViewComboBox.SelectedIndexChanged, _
            'AddressOf dataGridViewComboBox_SelectedIndexChanged
            '2021/04/12 ekawai del E↑------------------
            コンボイベントハンドラ削除()
            Me.dataGridViewComboBox = Nothing
        End With

    End Sub

    Private Sub DGV見積_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DGV見積.CellEndEdit
        Dim DR As DataRow()

        Select Case DGV見積.Columns(e.ColumnIndex).Name
            'カーソル入れて別のセルに移動しただけでも発生するため、エラーの原因になりそうだったので、CellValueChangedに移動させた
            'Case "見積単価"
            '    If Main_Form.nz(DGV見積.Rows(e.RowIndex).Cells("見積単価").Value) > 0 Then
            '        DGV見積.Rows(e.RowIndex).Cells("見積状態").Value = "発注希望"
            '    Else
            '        DGV見積.Rows(e.RowIndex).Cells("見積状態").Value = "見積希望"
            '    End If
            Case "科目番号"
                If kc.nz(DGV見積(e.ColumnIndex, e.RowIndex).Value) <> 0 Then
                    Dim i科目ID As Integer
                    '科目をリセットする
                    DGV見積("科目", e.RowIndex).Value = DBNull.Value
                    i科目ID = kc.nz(DGV見積.Rows(e.RowIndex).Cells("科目番号").Value)
                    '科目番号に応じた科目コンボを選択する
                    DR = Main_Form.DT現科目.Select("id = " & i科目ID)
                    If DR.Length = 1 Then
                        DGV見積("科目", e.RowIndex).Value = DR(0)("id")
                    Else
                        DGV見積(e.ColumnIndex, e.RowIndex).Value = DBNull.Value
                    End If
                End If
                '仕入先マスタに存在する仕入先名を入力したら、仕入先IDが自動で入るようにする
            Case "仕入先"
                Dim str仕入先名 As String
                str仕入先名 = kc.ns(DGV見積("仕入先", e.RowIndex).Value)
                GET仕入先ID(str仕入先名, e.RowIndex)
        End Select


    End Sub
    	'2022/03/15 ekawai add S
    Private Sub DGV見積_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DGV見積.CellEnter
        If TypeOf DGV見積.Columns(e.ColumnIndex) Is DataGridViewComboBoxColumn Then
            'コンボボックスのドロップダウンリストが一回のクリックで表示されるようにする
            SendKeys.Send("{F4}")
        End If

    End Sub
   '2022/03/15 ekawai add E
    Private Sub DGV見積_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DGV見積.CellValueChanged
        Dim str編集列名 As String = DGV見積.Columns(e.ColumnIndex).Name

        If e.RowIndex >= 0 Then
            If str編集列名 = "見積単価" Then
                If kc.nz(DGV見積.Rows(e.RowIndex).Cells("見積単価").Value) = 0 Then
                    DGV見積.Rows(e.RowIndex).Cells("見積状態").Value = "見積希望"
                Else
                    DGV見積.Rows(e.RowIndex).Cells("見積状態").Value = "発注希望"
                End If

            End If
        End If
    End Sub

    Private Sub GET仕入先ID(sName As String, iRow As Integer)
        Dim DR As DataRow()
        DGV見積("仕入先ID", iRow).Value = DBNull.Value
        If sName <> "" Then
            DR = DT見積部署別仕入先.Select("kiban = " & kc.SQ(sName))
            Select Case DR.Length
                Case 0
                    MsgBox(Select_Form.s部署 & "の仕入先マスタに存在しない仕入先が入力されました。")

                Case 1
                    'DGV見積("仕入先ID", iRow).Value = DR(0)("facility_id")　'2021/04/12 ekawai del 
                    DGV見積("仕入先ID", iRow).Value = DR(0)("仕入先ID") '2021/04/12 ekawai add 
                    '支払区分自動入力
                    DGV見積("支払区分", iRow).Value = DR(0)("pay_code")

            End Select
        End If
    End Sub

    Private Sub DGV見積_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DGV見積.CurrentCellDirtyStateChanged
        'セルに文字入れたらここを通るので編集済みフラグを立てる
        SwEditSave = True
    End Sub


    Private Sub DGV見積_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DGV見積.DataError
        Dim retu As String = DGV見積.Columns(0).HeaderText
        Dim naiyou As String = DGV見積.Rows(e.RowIndex).Cells(retu).Value
        Dim str As String = e.Exception.Message
        'e.Cancel = True
    End Sub

    Private Sub DGV見積_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DGV見積.EditingControlShowing
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim str編集列名 As String = dgv.CurrentCell.OwningColumn.Name
        If str編集列名 = "仕入先候補" Or str編集列名 = "購入者候補" Or str編集列名 = "科目" Or str編集列名 = "部門" Then
            Me.dataGridViewComboBox = CType(e.Control, DataGridViewComboBoxEditingControl)
            'If str編集列名 = "科目" Or str編集列名 = "部門" Then　'2021/04/12 ekawai del
            If str編集列名 = "科目" Or str編集列名 = "部門" Or str編集列名 = "仕入先候補" Then '2021/04/12 ekawai add 
                If 変更前Value = "" Then
                    変更前index = -1
                Else
                    変更前index = Me.dataGridViewComboBox.SelectedIndex
                End If
            End If
            '2021/04/12 ekawai del S↓------------------
            'AddHandler Me.dataGridViewComboBox.SelectedIndexChanged, _
            'AddressOf dataGridViewComboBox_SelectedIndexChanged
            '2021/04/12 ekawai del E↑------------------
            コンボイベントハンドラ追加() '2021/04/12 ekawai add 
        Else
            If TypeOf e.Control Is DataGridViewTextBoxEditingControl Then
                'DGrVTBEC = CType(e.Control, DataGridViewTextBoxEditingControl)
                '編集のために表示されているコントロールを取得
                Dim tb As DataGridViewTextBoxEditingControl = _
                    CType(e.Control, DataGridViewTextBoxEditingControl)
                '数字(マイナス許可)
                RemoveHandler tb.KeyPress, AddressOf dataGridViewTextBox_KeyPressMinus
                If str編集列名 = "数量" Or str編集列名 = "見積単価" Then
                    'KeyPressイベントハンドラを追加
                    AddHandler tb.KeyPress, AddressOf dataGridViewTextBox_KeyPressMinus
                End If
                '数字(プラスのみ)
                RemoveHandler tb.KeyPress, AddressOf dataGridViewTextBox_KeyPress
                If str編集列名 = "科目番号" Then
                    'KeyPressイベントハンドラを追加
                    AddHandler tb.KeyPress, AddressOf dataGridViewTextBox_KeyPress
                End If


            End If
        End If

    End Sub
    Private Sub dataGridViewTextBox_KeyPressMinus(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        '数字しか入力できないようにする
        If Asc(e.KeyChar) = 8 Then Exit Sub
        If e.KeyChar = "." Then
            Exit Sub
        End If
        If e.KeyChar = "-" Then
            Exit Sub
        End If
        If e.KeyChar < "0"c Or e.KeyChar > "9"c Then
            If e.KeyChar <> ""c Then
                e.Handled = True
                Exit Sub
            End If
        End If
    End Sub
    '----------------------------------------------------------------------------------------------------------
    'DataGridViewに表示されているテキストボックスのKeyPressイベントハンドラ 数値ＯＮＬＹ
    '----------------------------------------------------------------------------------------------------------

    Private Sub dataGridViewTextBox_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        '数字しか入力できないようにする
        If Asc(e.KeyChar) = 8 Then Exit Sub 'バックスペース
        If e.KeyChar = "." Then Exit Sub

        If e.KeyChar = "-" Then
            e.Handled = True
            sender.text = kc.nz(sender.text) * -1
            Exit Sub
        End If

        If e.KeyChar < "0"c Or e.KeyChar > "9"c Then
            If e.KeyChar <> ""c Then
                e.Handled = True
                Exit Sub
            End If
        End If
    End Sub

    Private Sub btn見積印刷_Click(sender As Object, e As EventArgs) Handles btn見積印刷.Click
        Dim strSql As String
        Dim xlsTemplate As String = ""
        xlsTemplate = "\\FS1\060System\Excelテンプレート\購入依頼_見積書.xlsx"
        Dim oApp As Excel.Application
        Dim oBooks As Excel.Workbooks
        Dim oBook As Excel.Workbook
        Dim oSheets As Excel.Sheets
        Dim oSheet As Excel.Worksheet
        Dim oDialogs As Excel.Dialogs
        Dim oDialog As Excel.Dialog
        Dim Arr(,) As Object = Nothing

        If SwEditSave = True Then
            If MsgBox("保存してよろしいですか?" & vbCrLf & "「いいえ」を押すと更新していないデータは消えます", vbYesNo + vbDefaultButton1) = vbYes Then
                If 見積更新() = False Then
                    Exit Sub
                End If
            Else
                SwEditSave = False
            End If
        End If

        '更新成功したら印刷処理に進む
        'Applicationインスタンスの生成
        oApp = New Excel.Application

        'oAppが持つWorkbooksプロパティを取得
        oBooks = oApp.Workbooks
        'oBooksに、ExcelFileNameで指定したファイルをオープンし、そのインスタンスをoBookに格納します。
        oBook = oBooks.Open(xlsTemplate)
        'ワークシートを選択
        oSheets = oBook.Worksheets
        oSheet = DirectCast(oSheets("見積書"), Excel.Worksheet)
        oDialogs = oApp.Dialogs
        oDialog = oDialogs(Excel.XlBuiltInDialog.xlDialogPrint)

        'セルの領域を選択
        Dim oCells1 As Excel.Range = Nothing  ' 中継用セル
        Dim rngFrom1 As Excel.Range = Nothing  ' 始点セル指定用
        Dim rngTo1 As Excel.Range = Nothing    ' 終点セル指定用
        Dim rngTarget1 As Excel.Range = Nothing ' 貼付け範囲指定用

        Dim rng仕入先 As Excel.Range = Nothing

        Dim rng購入者 As Excel.Range = Nothing

        Dim fileName As String = ""

        Dim 見積書明細S As Integer = 8
        Dim i明細END列 As Integer = 5
        Dim i明細数 As Integer = 10

        oApp.Visible = False
        oApp.DisplayAlerts = False

        Dim dtOriginal As New DataTable
        Dim dtCopy As New DataTable
        dtOriginal = DGV見積.DataSource
        'ヘッダー
        dtCopy.Columns.Add("仕入先")

        '明細
        dtCopy.Columns.Add("UID")
        dtCopy.Columns.Add("購入者")
        dtCopy.Columns.Add("品名")
        dtCopy.Columns.Add("型式")
        dtCopy.Columns.Add("メーカー")
        dtCopy.Columns.Add("数量")
        dtCopy.Columns.Add("購入理由")
        dtCopy.Columns.Add("見積状態")

        Dim CopyRow As DataRow

        For Each OriRow In dtOriginal.Rows
            '現金は除く
            'If ns(OriRow("pay_code")) <> "現金" Then
            CopyRow = dtCopy.NewRow
            CopyRow("UID") = kc.ns(OriRow("UID"))
            CopyRow("購入者") = kc.ns(OriRow("購入者"))
            CopyRow("仕入先") = kc.ns(OriRow("仕入先"))
            CopyRow("品名") = kc.ns(OriRow("品名"))
            CopyRow("型式") = kc.ns(OriRow("型式"))
            CopyRow("メーカー") = kc.ns(OriRow("メーカー"))
            CopyRow("数量") = kc.nz(OriRow("数量"))
            CopyRow("購入理由") = kc.ns(OriRow("購入理由"))
            CopyRow("見積状態") = kc.ns(OriRow("見積状態"))
            dtCopy.Rows.Add(CopyRow)
            'End If
        Next

        Dim viw As New DataView(dtCopy)

        Dim dt仕入先リスト As DataTable = viw.ToTable(True, {"仕入先", "購入者"}) '重複を削除する
        dt仕入先リスト.Columns.Add("Count", GetType(Integer))
        'For Each row As DataRow In dt仕入先リスト.Rows
        '    Dim expr As String = String.Format("仕入先 = '{0}' AND 購入者 = '{1}'", row("仕入先"), row("購入者"))
        '    row("Count") = dtOriginal.Compute("COUNT(数量)", expr)
        'Next
        Dim dt仕入先別 As New DataTable
        Dim dt見積書 As New DataTable

        Dim str見積書ID As String = ""

        For i = 0 To dt仕入先リスト.Rows.Count - 1
            oApp.Visible = True
            '直接dtに入れると0件の時エラーになるのでカウントが1以上になるのを確認してから入れる
            If dt仕入先別 IsNot Nothing Then
                dt仕入先別.Clear()
            End If
            dt仕入先別 = dtCopy.Select("仕入先 = '" & dt仕入先リスト.Rows(i)("仕入先") & "'").CopyToDataTable
            dt仕入先別 = dtCopy.Select("購入者 = '" & dt仕入先リスト.Rows(i)("購入者") & "'").CopyToDataTable
            '見積希望のものだけ印刷する
            If dt仕入先別.Select("見積状態 = '見積希望'").Count > 0 Then
                If dt見積書 IsNot Nothing Then
                    dt見積書.Clear()
                End If
                dt見積書 = dt仕入先別.Select("見積状態 = '見積希望'").CopyToDataTable
                If dt見積書 Is Nothing Then
                    '次の仕入先+購入者の組み合わせへ
                    Continue For
                Else
                    Dim view As New DataView(dt見積書)
                    Dim dt印刷対象 As New DataTable

                    dt印刷対象 = view.ToTable(False, {"品名", "型式", "メーカー", "数量", "購入理由"}) '重複を削除しない

                    rng仕入先 = oSheet.Range("A5")
                    rng仕入先.Value = dt仕入先リスト.Rows(i)("仕入先")
                    MRComObject(rng仕入先)


                    rng購入者 = oSheet.Range("C22")
                    rng購入者.Value = Select_Form.s部署 & " " & dt仕入先リスト.Rows(i)("購入者")
                    MRComObject(rng購入者)

                    '対象シートのセル
                    oCells1 = oSheet.Cells

                    '始点セル
                    rngFrom1 = DirectCast(oCells1(見積書明細S, i明細START列), Excel.Range)
                    'rngFrom = oCells(見積書明細S, i明細START列)


                    '終点セル
                    rngTo1 = DirectCast(oCells1(見積書明細S + i明細行数 - 1, dt印刷対象.Columns.Count), Excel.Range)
                    'rngTo = oCells(見積書明細S + dt見積書.Rows.Count - 1, i明細END列)

                    '貼り付け範囲作成
                    rngTarget1 = oSheet.Range(rngFrom1, rngTo1)


                    MRComObject(rngTo1)
                    MRComObject(rngFrom1)
                    MRComObject(oCells1)


                    ' 配列を明細の範囲に貼り付け
                    '1ページ10行なので入らなかったら次のページへ
                    Dim p = 0 'ページ数

                    Dim ret As Integer = Math.Ceiling(dt見積書.Rows.Count / i明細行数) - 1


                    '印刷頁のループ
                    For p = 0 To ret
                        Dim sep As String = ""
                        str見積書ID = ""
                        sep = "" 'ページごとに初期化

                        '二次元配列データ作成　
                        ReDim Arr(i明細行数 - 1, dt印刷対象.Columns.Count - 1)
                        Dim counter As Integer = 1

                        Try
                            '見積書の明細が10行なので0~9までループ
                            For iRow As Integer = 0 To i明細行数 - 1
                                'counterのほうが大きくなったら配列に空白を入れる
                                If counter <= dt印刷対象.Rows.Count - (i明細数 * p) Then
                                    str見積書ID = str見積書ID & sep & kc.SQ(dt見積書.Rows(iRow + (i明細数 * p))("UID"))
                                    sep = ","
                                End If


                                '見積書の明細列数は5

                                For iCol As Integer = 0 To dt印刷対象.Columns.Count - 1
                                    If counter <= dt印刷対象.Rows.Count - (i明細数 * p) Then
                                        Arr(iRow, iCol) = dt印刷対象.Rows(iRow + (i明細数 * p))(iCol)
                                    Else
                                        Arr(iRow, iCol) = ""
                                    End If

                                Next
                                counter = counter + 1
                            Next

                        Catch ex As Exception
                            Throw
                        End Try
                        rngTarget1.Value = Arr

                        oSheet.Select()
                        'FAXで送りたい仕入先もあるから印刷ダイアログ出す
                        Try
                            oDialog.Show()
                        Catch ex As System.Runtime.InteropServices.COMException
                            'PCFAXのダイヤログ開いた後で「キャンセル」押すとハンドルされていない例外が発生する対策
                        End Try
                        rngTarget1.ClearContents()
						Flg印刷対象有 = True
                        '見積希望から見積待ちにする
                        '発注済みのデータを消す
                        strSql = ""
                        strSql &= vbCrLf & "UPDATE TD_見積 SET 見積状態 = '見積待ち'"
                        strSql &= vbCrLf & "WHERE UID IN (" & str見積書ID & ")"
                        Dim tran As System.Data.SqlClient.SqlTransaction = Nothing

                        tran = dCon.Connection.BeginTransaction


                        If dCon.ExecuteSqlMW(tran, strSql) = False Then
                            tran.Rollback()
                            MsgBox("見積状態の更新に失敗しました!")
                        Else
                            tran.Commit()

                        End If
                        tran.Dispose()

                    Next
                    MRComObject(rngTarget1)
                End If
            End If '見積希望のループ

        Next '仕入先リストのループ
        MRComObject(oSheet)
        MRComObject(oSheets)
        oBook.Close(False)
        MRComObject(oBook)
        MRComObject(oBooks)
        MRComObject(oDialog)
        MRComObject(oDialogs)

        oApp.Quit()
        oApp.DisplayAlerts = True
        MRComObject(oApp)
        DGV見積_Read()
        If Flg印刷対象有 Then
            MsgBox("印刷完了")
        End If
        Flg印刷対象有 = False






    End Sub
    'Private Sub 配列格納(ByVal dt As DataTable, ByVal p As Integer, ByRef arr(,) As Object)
    '    Dim view As New DataView(dt)
    '    Dim counter As Integer = 1 'カウント用変数
    '    dt = view.ToTable(False, {"品名", "型式", "メーカー", "数量", "購入理由"}) '重複を削除しない


    '    '二次元配列データ作成　

    '    Try

    '        ReDim arr(i明細行数 - 1, i明細END列 - 1)
    '        '見積書の明細が10行なので0~9までループ
    '        For iRow As Integer = 0 To i明細行数 - 1
    '            'counterのほうが大きくなったら配列に空白を入れる
    '            If (i明細行数 * p) + counter > dt.Rows.Count Then
    '                '見積書の明細列数は5
    '                For iCol As Integer = 0 To i明細END列 - 1
    '                    arr(iRow, iCol) = ""
    '                Next
    '            Else
    '                For iCol As Integer = 0 To i明細END列 - 1
    '                    '明細の行数が10なのでページ数に10をかけている
    '                    arr(iRow, iCol) = dt.Rows(iRow + (i明細行数 * p))(iCol)
    '                Next

    '            End If
    '            counter = counter + 1
    '        Next

    '    Catch ex As Exception
    '        Throw
    '    End Try


    'End Sub
    Public Sub MRComObject(Of T As Class)(ByRef objCom As T, Optional ByVal force As Boolean = False)
        Dim IDEEnvironment As Boolean = False  'メッセージボックスを表示させたい場合は、True に設定
        If objCom Is Nothing Then
            If IDEEnvironment = True Then
                'テスト環境の場合は下記を実施し、後は、コメントにしておいて下さい。
                MessageBox.Show(Me, "Nothing です。")
            End If
            Return
        End If
        Try
            If System.Runtime.InteropServices.Marshal.IsComObject(objCom) Then
                Dim count As Integer
                If force Then
                    count = System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objCom)
                Else
                    count = System.Runtime.InteropServices.Marshal.ReleaseComObject(objCom)
                End If
                If IDEEnvironment = True AndAlso count <> 0 Then
                    Try
                        'テスト環境の場合は下記を実施し、後は、コメントにしておいて下さい。
                        MessageBox.Show(Me, TypeName(objCom) & " 要調査！ デクリメントされていません。")
                    Catch ex As Exception
                        MessageBox.Show(Me, " 要調査！ デクリメントされていません。")
                    End Try
                End If
            Else
                If IDEEnvironment = True Then
                    'テスト環境の場合は下記を実施し、後は、コメントにしておいて下さい。
                    MessageBox.Show(Me, "ComObject ではありませんので、解放処理の必要はありません。")
                End If
            End If
        Finally
            objCom = Nothing
        End Try
    End Sub

    Private Sub btn削除_Click(sender As Object, e As EventArgs) Handles btn削除.Click
        Dim i削除行No As Integer = DGV見積.SelectedCells(0).RowIndex
        Dim strSql As String = ""
        Dim tran As System.Data.SqlClient.SqlTransaction = Nothing
        Dim Response As String
        Dim msg As String = ""

        If i削除行No + 1 = DGV見積.Rows.Count Then
            'DataGridViewの一番下の新規行(自動で作られる行)は削除できない
            MsgBox("新規行は削除できません")
            Exit Sub
        End If



        DGV見積.Rows(i削除行No).Selected = True
        If kc.nz(DGV見積.Rows(i削除行No).Cells("UID").Value) = 0 Then
            'まだデータベースに登録されていないから見た目上行を消したら終わり
            Response = MsgBox("現在選択している行を削除してよろしいですか？", MsgBoxStyle.YesNo, "確認")
            If Response = vbYes Then
                DGV見積.Rows.RemoveAt(i削除行No)
            End If
        Else

            Response = MsgBox("現在選択している行を削除してよろしいですか？", MsgBoxStyle.YesNo, "確認")
            'If Response = vbYes Then
            'DGV見積.Rows.RemoveAt(i削除行No)
            'SwEditSave = True
            'End If


            If Response = vbYes Then
                strSql = "DELETE FROM TD_見積"
                strSql &= vbCrLf & "WHERE UID = " & DGV見積.Rows(i削除行No).Cells("UID").Value
                tran = dCon.Connection.BeginTransaction
                If dCon.ExecuteSqlMW(tran, strSql) = False Then
                    tran.Rollback()
                    tran.Dispose()

                    MsgBox("削除失敗")
                    Exit Sub

                End If

                tran.Commit()
                tran.Dispose()

                DGV見積_Read()

                MsgBox("削除成功")

            End If

        End If


        SwEditSave = False
    End Sub
    Sub コンボイベントハンドラ削除() 'SelectedIndexChangedが何回も発生してStackOverflowでプログラムが落ちる対策
        RemoveHandler Me.dataGridViewComboBox.SelectedIndexChanged, AddressOf dataGridViewComboBox_SelectedIndexChanged
    End Sub
    Sub コンボイベントハンドラ追加()
        AddHandler Me.dataGridViewComboBox.SelectedIndexChanged, AddressOf dataGridViewComboBox_SelectedIndexChanged
    End Sub

End Class