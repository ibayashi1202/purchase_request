﻿Imports Microsoft.Office.Interop
Public Class 製品_Form
    Inherits System.Windows.Forms.Form
    Private dCon As New merrweth_init_DbConnection
    Public f1 As Main_Form

    Private dataGridViewComboBox As DataGridViewComboBoxEditingControl = Nothing
    Private searchFld As nameValue = New nameValue
    'Private CellEditStart As Boolean = False 使ってなさそうだからコメントアウト

    Private Editcell As DataGridViewCell
    Private DT_po As DataTable
    Private SwEditSave As Boolean = False
    Private arr() As Object
    '引数として受け取ったメソッド保存用 
    Private CallBack As ValueChange

    Private 検索条件 As String
    Private 検索Flg As Boolean = False
    Private 日付検索Flg As Boolean = False

    'コールバック用としてデリゲートを宣言 
    Public Delegate Sub ValueChange(ByVal arrOriginal() As Object, Form As String)

    Private arr見積情報(9) As Object
    Public 変更前index As Integer
    Public 変更前Value As String

    'Private s型 As String 
    Private 製品CloseFlg As Boolean '右上の閉じるボタンを押した時にTRUEにする
    Public kc As New 共通処理_Class
    Private DT製品部署別仕入先 As DataTable '2022/03/24 ekawai

    '2022/07/13 ekawai メインフォームの縦貼り付け機能を移植するために追加
    Public Active行 As Integer
    Public Active列 As Integer



    Private Sub 製品_Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '右上の閉じるボタンから閉じられるとき
        If e.CloseReason = CloseReason.UserClosing Then
            If SwEditSave = True Then
                'Main_Formは画面書き換えたり印刷する前にも保存チェックしているが、製品と見積はで保存漏れがあっても、たいした問題にならないから全ボタンに保存確認機能付けなくてよいとのことだったので、閉じる時だけ保存確認している
                '保存機能があるのがいいのかどうか現時点では判断できないので使っていく中で問題があれば追加すればよいらしい
                If MsgBox("保存してよろしいですか？" & vbCrLf & "「いいえ」を押すと更新していないデータは消えます", vbYesNo + vbDefaultButton1) = vbYes Then
                    'If 製品更新() = False Then
                    '    '更新失敗したら閉じない
                    '    e.Cancel = True
                    'End If

                    btn製品_更新.PerformClick()
                    If 製品CloseFlg = False Then
                        '万一更新が失敗した場合はフォームを閉じない…内容を修正して再度更新するか更新しないで閉じるかユーザーに選ばせる
                        e.Cancel = True
                    End If

                Else
                    SwEditSave = False

                End If
            End If
        End If
    End Sub

    Private Sub 製品_Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim strSql As String
        '///検索項目コンボボックスの共通設定////
        Dim DT変換1 As New DataTable
        Dim DT変換2 As New DataTable
        Dim DT変換3 As New DataTable


        DGV製品.MultiSelect = False '複数選択禁止
        DGV製品.RowHeadersWidth = 20
        DGV製品_Read()
        DGV製品_詳細設定()
        DGV製品.AllowUserToDeleteRows = False 'Deleteボタンによる行削除禁止(DGVにdatasourceセットする前にFalseにしたらエラーになる)
        DGV製品.AutoGenerateColumns = False

        cmb検索項目1.DataSource = Nothing
        cmb検索項目2.DataSource = Nothing
        cmb検索項目3.DataSource = Nothing

        strSql = ""
        strSql = "SELECT * FROM TM_列名変換 WHERE 型 <> '日付' AND 製品 = 1"


        DT変換1 = dCon.DataSet(strSql, "TB1").Tables(0)
        DT変換2 = dCon.DataSet(strSql, "TB2").Tables(0)
        DT変換3 = dCon.DataSet(strSql, "TB3").Tables(0)


        For i = 1 To 3
            RemoveHandler CType(Me.Controls("cmb検索項目" & i), ComboBox).SelectedIndexChanged, AddressOf cmb検索項目_SelectedIndexChanged
            CType(Me.Controls("cmb検索項目" & i), ComboBox).DisplayMember = "表示列"
            'CType(Me.Controls("cmb検索項目" & i), ComboBox).ValueMember = "表示列" '製品の時はValueMemberも表示列になる(日本語の列名だから)
            CType(Me.Controls("cmb検索項目" & i), ComboBox).ValueMember = "製品用データ列" '科目名で検索できるようにするためTM_列名変換にフィールド追加した
            Select Case i
                Case 1
                    CType(Me.Controls("cmb検索項目" & i), ComboBox).DataSource = DT変換1
                Case 2
                    CType(Me.Controls("cmb検索項目" & i), ComboBox).DataSource = DT変換2
                Case 3
                    CType(Me.Controls("cmb検索項目" & i), ComboBox).DataSource = DT変換3
            End Select

            CType(Me.Controls("cmb検索項目" & i), ComboBox).SelectedIndex = -1
            AddHandler CType(Me.Controls("cmb検索項目" & i), ComboBox).SelectedIndexChanged, AddressOf cmb検索項目_SelectedIndexChanged
            CType(Me.Controls("txt検索条件" & i), TextBox).Enabled = False

        Next


        If Select_Form.更新Flg = False And Select_Form.承認Flg = False And Select_Form.経理Flg = False Then
            btn製品_更新.Visible = False
            btn製品_見積コピー.Visible = False
            btn製品to購入.Visible = False

        End If
        'グリッドの並び替えを禁止する。(ただしプログラムからの並び替えは可能とする)
        For Each c As DataGridViewColumn In DGV製品.Columns
            c.SortMode = DataGridViewColumnSortMode.Programmatic
        Next c
        'ヘッダーを除く表示行の幅に自動調整
        DGV製品.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders

    End Sub

    Public Sub DGV製品_Read()
        Dim ds製品リスト As DataSet
        DGV製品.DataSource = Nothing

        Dim strSql As String = ""
        '貼り付けボタン利用時、仕入先と仕入先IDの順番が重要！
        strSql = strSql & vbCrLf & "SELECT TM_製品.UID,TM_製品.登録番号,TM_製品.部署,TM_製品.部門ID,TM_製品.購入者,TM_製品.品名,TM_製品.型式,TM_製品.数量,TM_製品.予算単価"
        strSql = strSql & vbCrLf & ",TM_製品.仕入先,TM_製品.仕入先ID,TM_製品.支払区分,TM_製品.科目ID,TM_製品.購入理由,TM_製品.備考"
        strSql = strSql & vbCrLf & ",TM_kamoku.kamoku as 科目名,TM_division.division as 部門名"
        strSql = strSql & vbCrLf & "FROM TM_製品 "
        strSql = strSql & vbCrLf & "LEFT OUTER JOIN TM_division"
        strSql = strSql & vbCrLf & "ON 部門ID = TM_division.id"
        strSql = strSql & vbCrLf & "LEFT OUTER JOIN TM_kamoku"
        strSql = strSql & vbCrLf & "ON 科目ID = TM_kamoku.id"

        strSql = strSql & vbCrLf & "WHERE 部署 = '" & Select_Form.s部署 & "'"
        If 検索Flg Then
            strSql = strSql & vbCrLf & 検索条件
        End If
        strSql = strSql & vbCrLf & "ORDER BY TM_製品.UID"

        ds製品リスト = dCon.DataSet(strSql, "製品一覧")
        DGV製品.DataSource = ds製品リスト.Tables("製品一覧")
        DGV製品.FirstDisplayedScrollingRowIndex = DGV製品.RowCount - 1

    End Sub

    Private Sub DGV製品_詳細設定()
        Dim strSql As String
        Dim Col部門 As New DataGridViewComboBoxColumn
        Dim Col購入者 As New DataGridViewComboBoxColumn
        Dim Col仕入先 As New DataGridViewComboBoxColumn
        Dim Col支払区分 As New DataGridViewComboBoxColumn
        Dim Col科目 As New DataGridViewComboBoxColumn
        Dim Col科目候補 As New DataGridViewTextBoxColumn
        '------------------------------------------------------------------------------
        '部門 
        '------------------------------------------------------------------------------
        'DGVに現在存在しているdivision_id列と今作成したDataGridViewComboBoxColumnを入れ替える
        Col部門.DataPropertyName = DGV製品.Columns("部門ID").DataPropertyName
        Col部門.DataSource = Main_Form.DT部門
        Col部門.ValueMember = "id"
        Col部門.DisplayMember = "division"
        DGV製品.Columns.Insert(DGV製品.Columns("部門ID").Index, Col部門)
        'DGV購入.Columns("division_id").Visible = False
        Col部門.Name = "部門"
        '------------------------------------------------------------------------------
        '購入者…自由入力可
        '------------------------------------------------------------------------------
        For Each DR As DataRow In Main_Form.DT部署別購入者.Rows
            Col購入者.Items.Add(DR("employee"))
        Next
        DGV製品.Columns.Insert(DGV製品.Columns("購入者").Index + 1, Col購入者)
        'DGV購入.Columns("employee").Visible = False
        Col購入者.Name = "購入者候補"
        Col購入者.HeaderText = ""
        Col購入者.DropDownWidth = 200
        '------------------------------------------------------------------------------
        '仕入先…自由入力可
        '------------------------------------------------------------------------------
        '2021/04/12 ekawai add S
        strSql = "SELECT"
        strSql = strSql & vbCrLf & "TM_製品.仕入先ID"
        strSql = strSql & vbCrLf & "  ,'×' + TM_facility.kiban AS kiban"
        strSql = strSql & vbCrLf & "  ,TM_paycode.pay_code"
        strSql = strSql & vbCrLf & "  , 2 AS 並び順 "
        strSql = strSql & vbCrLf & "FROM TM_製品 "
        strSql = strSql & vbCrLf & "  INNER JOIN TM_facility "
        strSql = strSql & vbCrLf & "    ON TM_製品.仕入先ID = TM_facility.facility_id "
        strSql = strSql & vbCrLf & "  INNER JOIN TM_paycode ON TM_facility.pay_method = TM_paycode.id"
        strSql = strSql & vbCrLf & "  LEFT OUTER JOIN ( "
        strSql = strSql & vbCrLf & "    SELECT * "
        strSql = strSql & vbCrLf & "    FROM TM_dep_facility "
        strSql = strSql & vbCrLf & "    WHERE"
        strSql = strSql & vbCrLf & "    TM_dep_facility.部署ID = " & Select_Form.i部署ID
        strSql = strSql & vbCrLf & "  ) AS Sub1 "
        strSql = strSql & vbCrLf & "    ON TM_製品.仕入先ID = Sub1.仕入先ID "
        strSql = strSql & vbCrLf & "WHERE"
        strSql = strSql & vbCrLf & "  TM_製品.部署 = " & kc.SQ(Select_Form.s部署)
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
        DT製品部署別仕入先 = dCon.DataSet(strSql, "DT").Tables(0)
        '2021/04/12 ekawai add E
        'For Each DR As DataRow In Main_Form.DT部署別仕入先.Rows　'2021/04/12 ekawai del
        For Each DR As DataRow In DT製品部署別仕入先.Rows

            Col仕入先.Items.Add(DR("kiban"))
        Next
        DGV製品.Columns.Insert(DGV製品.Columns("仕入先").Index + 1, Col仕入先)
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

        Col支払区分.DataPropertyName = DGV製品.Columns("支払区分").DataPropertyName
        Col支払区分.DataSource = DT支払区分
        Col支払区分.ValueMember = "pay_code"
        Col支払区分.DisplayMember = "pay_code"
        '同じ列名が2個ある場合Visible = Falseにしたら前のほうだけ非表示になるっぽい
        DGV製品.Columns.Insert(DGV製品.Columns("支払区分").Index + 1, Col支払区分)
        Col支払区分.Name = "支払区分"
        Col支払区分.Width = 60
        '------------------------------------------------------------------------------
        '科目
        '-----------------------------------------------------------------------------
        strSql = "SELECT"
        strSql = strSql & vbCrLf & "TM_製品.科目ID"
        strSql = strSql & vbCrLf & ", '×' + CONVERT(nvarchar, TM_製品.科目ID) + ' : ' + TM_kamoku.kamoku AS kamoku"
        strSql = strSql & vbCrLf & ", 2 AS 並び順"
        strSql = strSql & vbCrLf & "FROM TM_製品"
        strSql = strSql & vbCrLf & "INNER JOIN TM_kamoku"
        strSql = strSql & vbCrLf & "ON TM_製品.科目ID = TM_kamoku.id"
        strSql = strSql & vbCrLf & "LEFT OUTER JOIN ( "
        strSql = strSql & vbCrLf & "    SELECT * "
        strSql = strSql & vbCrLf & "    FROM TM_dep_kamoku "
        strSql = strSql & vbCrLf & "    WHERE"
        strSql = strSql & vbCrLf & "      TM_dep_kamoku.部署ID = " & Select_Form.i部署ID
        strSql = strSql & vbCrLf & "  ) AS Sub1 "
        strSql = strSql & vbCrLf & "ON TM_製品.科目ID = Sub1.科目ID "
        strSql = strSql & vbCrLf & "WHERE"
        strSql = strSql & vbCrLf & " TM_製品.部署 = '" & Select_Form.s部署 & "'"
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
        Dim DT部署別製品科目 As DataTable

        DT部署別製品科目 = dCon.DataSet(strSql, "DT").Tables(0)

        'DGVに現在存在しているdivision_id列と今作成したDataGridViewComboBoxColumnを入れ替える
        Col科目.DataPropertyName = DGV製品.Columns("科目ID").DataPropertyName
        Col科目.DataSource = DT部署別製品科目
        'Col科目.ValueMember = "id
        Col科目.ValueMember = "科目ID"
        Col科目.DisplayMember = "kamoku"
        'Col科目.DataSource = Main_Form.DT部署別科目


        DGV製品.Columns.Insert(DGV製品.Columns("科目ID").Index, Col科目)
        DGV製品.Columns("科目ID").Visible = False
        Col科目.Name = "科目"
        Col科目.DropDownWidth = 150

        '製造日報のようにテキストボックスに入れた値で科目ID検索できるようにという指示があったため列を追加した
        DGV製品.Columns.Insert(DGV製品.Columns("科目").Index, Col科目候補)
        Col科目候補.Name = "科目番号"

        '非表示
        DGV製品.Columns("UID").Visible = False
        DGV製品.Columns("部門ID").Visible = False
        'DGV製品.Columns("仕入先ID").Visible = False
        DGV製品.Columns("支払区分").Visible = False
        DGV製品.Columns("科目ID").Visible = False
        DGV製品.Columns("部署").Visible = False
        DGV製品.Columns("科目名").Visible = False
        DGV製品.Columns("部門名").Visible = False

        '幅
        DGV製品.Columns("登録番号").Width = 40
        DGV製品.Columns("部門").Width = 100
        DGV製品.Columns("購入者候補").Width = 20
        DGV製品.Columns("購入者").Width = 60
        DGV製品.Columns("品名").Width = 120
        DGV製品.Columns("型式").Width = 120
        DGV製品.Columns("数量").Width = 40
        DGV製品.Columns("予算単価").Width = 60
        DGV製品.Columns("仕入先候補").Width = 20
        DGV製品.Columns("仕入先").Width = 80
        DGV製品.Columns("仕入先ID").Width = 40
        'DGV製品.Columns("支払区分").Width = 40
        DGV製品.Columns("科目番号").Width = 40
        DGV製品.Columns("科目").Width = 140
        DGV製品.Columns("購入理由").Width = 140
        DGV製品.Columns("備考").Width = 100

        '読取専用
        DGV製品.Columns("仕入先ID").ReadOnly = True

        DirectCast(DGV製品.Columns("購入者"), DataGridViewTextBoxColumn).MaxInputLength = 40
        DirectCast(DGV製品.Columns("品名"), DataGridViewTextBoxColumn).MaxInputLength = 125
        DirectCast(DGV製品.Columns("型式"), DataGridViewTextBoxColumn).MaxInputLength = 125
        DirectCast(DGV製品.Columns("仕入先"), DataGridViewTextBoxColumn).MaxInputLength = 50
        DirectCast(DGV製品.Columns("購入理由"), DataGridViewTextBoxColumn).MaxInputLength = 120

        '折返し設定
        DGV製品.Columns("購入者").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        DGV製品.Columns("品名").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        DGV製品.Columns("型式").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        DGV製品.Columns("仕入先").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        DGV製品.Columns("購入理由").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        DGV製品.Columns("備考").DefaultCellStyle.WrapMode = DataGridViewTriState.True
    End Sub

    Private Sub btn製品to購入_Click(sender As Object, e As EventArgs) Handles btn製品to購入.Click
        Dim i選択行No As Integer = DGV製品.SelectedCells(0).RowIndex
        Dim dt As DataTable = DGV製品.DataSource
        Dim dv As DataRowVersion
        Dim iエラー行 As Integer

        dv = DataRowVersion.Current

        If DGV製品.Rows.Count = i選択行No + 1 Then
            'DGVの一番下の空白行が選択されている場合は何もしない
            MsgBox("コピーできない行が選択されています")
        Else
            Dim row As DataRow = dt.Rows(i選択行No)
            iエラー行 = Main_Form.必須項目Check(dt.Rows.IndexOf(row) + 1, row, dv, "製品", "購入品入力")
            If iエラー行 > 0 Then
                'エラー行にスクロール
                DGV製品.FirstDisplayedScrollingRowIndex = iエラー行 - 1
                DGV製品.Rows(iエラー行 - 1).Selected = True
                Exit Sub
            End If

            ReDim arr(11)
            'メインフォームが開いているかチェック
            If My.Application.OpenForms("Main_Form") Is Nothing Then
                MsgBox("購入品入力画面が開いていないためコピーできません")
            Else
                arr(0) = DGV製品.CurrentRow.Cells("登録番号").Value
                arr(1) = DGV製品.CurrentRow.Cells("部門ID").Value
                arr(2) = DGV製品.CurrentRow.Cells("購入者").Value
                arr(3) = DGV製品.CurrentRow.Cells("品名").Value
                arr(4) = DGV製品.CurrentRow.Cells("型式").Value
                arr(5) = DGV製品.CurrentRow.Cells("数量").Value
                arr(6) = DGV製品.CurrentRow.Cells("予算単価").Value
                arr(7) = DGV製品.CurrentRow.Cells("仕入先ID").Value
                arr(8) = DGV製品.CurrentRow.Cells("仕入先").Value
                arr(9) = DGV製品.CurrentRow.Cells("支払区分").Value
                arr(10) = DGV製品.CurrentRow.Cells("科目ID").Value
                arr(11) = DGV製品.CurrentRow.Cells("購入理由").Value

                If Not Me.CallBack Is Nothing Then Me.CallBack(arr, "製品")
                My.Application.OpenForms("Main_Form").WindowState = FormWindowState.Normal
                My.Application.OpenForms("Main_Form").Activate()

            End If

        End If
    End Sub


    Public Sub New(ByVal CallBack As ValueChange)
        MyBase.New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        Me.CallBack = CallBack
    End Sub



    Private Sub btn製品_更新_Click(sender As Object, e As EventArgs) Handles btn製品_更新.Click

        If 製品更新() Then
            製品CloseFlg = True
            DGV製品_Read()
        Else
            製品CloseFlg = False
            MsgBox("更新失敗")
        End If

    End Sub

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
        With DGV製品
            i = .CurrentCell.RowIndex
            Select Case .CurrentCell.OwningColumn.Name
                '仕入先に応じて支払区分を自動で入れる
                Case "仕入先候補"
                    '.Rows(i).Cells("仕入先").Value = s選択値　'2021/04/12 ekawai del
                    'GET仕入先ID(s選択値, i)
                    '×から始まる仕入先名を選んだら
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
                    コンボイベントハンドラ削除()
                    .Rows(i).Cells("購入者").Value = s選択値
                    コンボイベントハンドラ追加()
                Case "科目", "部門"
                    'カレントセルを取得する時こうやって指定しないとコンボの背景が黒くなるため
                    Dim sTarget As String = cb.EditingControlFormattedValue
                    If sTarget.Substring(0, 1) = "×" Then
                        MsgBox("現在使われていない項目です")
                        コンボイベントハンドラ削除()
                        cb.SelectedIndex = 変更前index
                        コンボイベントハンドラ追加()
                    End If

            End Select

            'If .CurrentCell.OwningColumn.Name = "仕入先候補" Then
            '    .Rows(i).Cells("仕入先").Value = s選択値
            '    strSql = "SELECT *"
            '    strSql &= vbCrLf & "FROM TM_facility"
            '    strSql &= vbCrLf & "WHERE kiban = '" & s選択値 & "'"
            '    Dim DT仕入先ID As DataTable = dCon.DataSet(strSql, "DT").Tables(0)
            '    'マスタに該当するIDがなければ空白を入れる
            '    If DT仕入先ID.Rows.Count = 0 Then
            '        .Rows(i).Cells("仕入先ID").Value = ""
            '    Else
            '        .Rows(i).Cells("仕入先ID").Value = DT仕入先ID.Rows(0)("facility_id")
            '    End If
            'ElseIf .CurrentCell.OwningColumn.Name = "購入者候補" Then
            '    .Rows(i).Cells("購入者").Value = s選択値

            'End If
        End With

        '2021/04/12 ekawai del S↓------------------
        'コンボの有無を判定してからRemoveしないとエラーになるためIF文追加
        'If Not (Me.dataGridViewComboBox Is Nothing) Then
        '    'SelectedIndexChangedイベントハンドラを削除
        '    RemoveHandler Me.dataGridViewComboBox.SelectedIndexChanged, _
        '       AddressOf dataGridViewComboBox_SelectedIndexChanged
        '    Me.dataGridViewComboBox = Nothing
        'End If
        '2021/04/12 ekawai del E↑------------------


    End Sub

    Private Sub DGV製品_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV製品.CellClick
        Active行 = e.RowIndex
        Active列 = e.ColumnIndex
    End Sub

    Private Sub DGV製品_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DGV製品.CellEndEdit
        Dim DR As DataRow()

        Select Case DGV製品.Columns(e.ColumnIndex).Name

            Case "科目番号"
                If kc.nz(DGV製品(e.ColumnIndex, e.RowIndex).Value) <> 0 Then
                    Dim i科目ID As Integer
                    '科目をリセットする
                    DGV製品("科目", e.RowIndex).Value = DBNull.Value
                    i科目ID = kc.nz(DGV製品.Rows(e.RowIndex).Cells("科目番号").Value)
                    '科目番号に応じた科目コンボを選択する
                    DR = Main_Form.DT現科目.Select("id = " & i科目ID)
                    If DR.Length = 1 Then
                        DGV製品("科目", e.RowIndex).Value = DR(0)("id")
                    Else
                        DGV製品(e.ColumnIndex, e.RowIndex).Value = DBNull.Value
                    End If
                End If

                '仕入先マスタに存在する仕入先名を入力したら、仕入先IDが自動で入るようにする
            Case "仕入先"
                Dim str仕入先名 As String
                str仕入先名 = kc.ns(DGV製品("仕入先", e.RowIndex).Value)
                GET仕入先ID(str仕入先名, e.RowIndex)
        End Select

    End Sub
    Private Sub GET仕入先ID(sName As String, iRow As Integer)
        Dim DR As DataRow()
        DGV製品("仕入先ID", iRow).Value = DBNull.Value
        If sName <> "" Then
            'DR = Main_Form.DT部署別仕入先.Select("kiban = " & kc.SQ(sName))　'2021/04/12 ekawai del
            DR = DT製品部署別仕入先.Select("kiban = " & kc.SQ(sName)) '2021/04/12 ekawai add S
            Select Case DR.Length
                Case 0
                    MsgBox(Select_Form.s部署 & "の仕入先マスタに存在しない仕入先が入力されました。")

                Case 1
                    DGV製品("仕入先ID", iRow).Value = DR(0)("仕入先ID") '2021/04/12 ekawai add S
                    '支払区分自動入力
                    DGV製品("支払区分", iRow).Value = DR(0)("pay_code")

            End Select
        End If
    End Sub

    Private Sub DGV製品_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DGV製品.CellEnter
        Dim dgv As DataGridView = CType(sender, DataGridView)
        If TypeOf dgv.Columns(e.ColumnIndex) Is DataGridViewComboBoxColumn Then
            'コンボボックスのドロップダウンリストが一回のクリックで表示されるようにする
            SendKeys.Send("{F4}")
            If dgv.Columns(e.ColumnIndex).Name = "科目" Or dgv.Columns(e.ColumnIndex).Name = "部門" Then
                変更前Value = dgv.Rows(e.RowIndex).Cells(e.ColumnIndex).EditedFormattedValue

            End If
        End If

    End Sub

    Private Sub DGV製品_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DGV製品.CurrentCellDirtyStateChanged
        SwEditSave = True
    End Sub
    Private Sub DGV製品_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DGV製品.DataError
        MsgBox(e.Exception.Message)
    End Sub
    Private Sub DGV製品_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DGV製品.EditingControlShowing
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim str編集列名 As String = dgv.CurrentCell.OwningColumn.Name
        'If str編集列名 = "仕入先候補" Or str編集列名 = "購入者候補" Or str編集列名 = "科目" Or str編集列名 = "部門" Then　'2021/04/12 ekawai del 
        If TypeOf e.Control Is DataGridViewComboBoxEditingControl And str編集列名 <> "支払区分" Then '2021/04/12 ekawai add 
            '編集のために表示されているコントロールを取得
            Me.dataGridViewComboBox = CType(e.Control, DataGridViewComboBoxEditingControl)
            'If str編集列名 = "科目" Or str編集列名 = "部門" Then　'2021/04/12 ekawai del 
            If str編集列名 = "科目" Or str編集列名 = "部門" Or str編集列名 = "仕入先" Then '2021/04/12 ekawai add 
                If 変更前Value = "" Then
                    変更前index = -1
                Else
                    変更前index = Me.dataGridViewComboBox.SelectedIndex
                End If
            End If
            'SelectedIndexChangedイベントハンドラを追加
            'AddHandler Me.dataGridViewComboBox.SelectedIndexChanged, _
            'AddressOf dataGridViewComboBox_SelectedIndexChanged
            コンボイベントハンドラ追加()
        Else
            If TypeOf e.Control Is DataGridViewTextBoxEditingControl Then
                'DGrVTBEC = CType(e.Control, DataGridViewTextBoxEditingControl)
                '編集のために表示されているコントロールを取得
                Dim tb As DataGridViewTextBoxEditingControl = _
                    CType(e.Control, DataGridViewTextBoxEditingControl)
                '数字(プラスのみ)
                RemoveHandler tb.KeyPress, AddressOf dataGridViewTextBox_KeyPress
                If str編集列名 = "科目番号" Then
                    'KeyPressイベントハンドラを追加
                    AddHandler tb.KeyPress, AddressOf dataGridViewTextBox_KeyPress
                End If
                '数字(マイナス許可)
                RemoveHandler tb.KeyPress, AddressOf dataGridViewTextBox_KeyPressMinus
                If str編集列名 = "数量" Or str編集列名 = "予算単価" Then
                    'KeyPressイベントハンドラを追加
                    AddHandler tb.KeyPress, AddressOf dataGridViewTextBox_KeyPressMinus
                End If

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



    '----------------------------------------------------------------------------------------------------------
    'DataGridViewに表示されているテキストボックスのKeyPressイベントハンドラ 数値ＯＮＬＹ(マイナス許可)
    '----------------------------------------------------------------------------------------------------------
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

    Private Sub btn承認_Click(sender As Object, e As EventArgs)

    End Sub


    Private Sub btn製品_見積コピー_Click(sender As Object, e As EventArgs) Handles btn製品_見積コピー.Click
        Dim strSql As String
        Dim i選択行No As Integer = DGV製品.SelectedCells(0).RowIndex
        Dim dt As DataTable = DGV製品.DataSource
        Dim dv As DataRowVersion
        Dim iエラー行 As Integer
        dv = DataRowVersion.Current



        If DGV製品.Rows.Count = i選択行No + 1 Then
            'DGVの一番下の空白行が選択されている場合は何もしない
            MsgBox("コピーできない行が選択されています")
        Else
            Dim tran As System.Data.SqlClient.SqlTransaction = Nothing
            Dim row As DataRow = dt.Rows(i選択行No)

            iエラー行 = Main_Form.必須項目Check(dt.Rows.IndexOf(row) + 1, row, dv, "製品", "見積")
            If iエラー行 > 0 Then
                'エラー行にスクロール
                DGV製品.FirstDisplayedScrollingRowIndex = iエラー行 - 1
                DGV製品.Rows(iエラー行 - 1).Selected = True

                Exit Sub
            End If
            strSql = "INSERT INTO"
            strSql &= vbCrLf & "TD_見積("
            strSql &= vbCrLf & "部署"
            strSql &= vbCrLf & ",部門ID"
            strSql &= vbCrLf & ",購入者"
            strSql &= vbCrLf & ",品名"
            strSql &= vbCrLf & ",型式"
            strSql &= vbCrLf & ",数量"
            strSql &= vbCrLf & ",見積単価"
            strSql &= vbCrLf & ",仕入先ID"
            strSql &= vbCrLf & ",仕入先"
            strSql &= vbCrLf & ",支払区分"
            strSql &= vbCrLf & ",科目ID"
            strSql &= vbCrLf & ",購入理由"
            strSql &= vbCrLf & ",見積状態"
            strSql &= vbCrLf & ") VALUES ("
            strSql &= vbCrLf & kc.SQ(Select_Form.s部署)
            strSql &= vbCrLf & "," & kc.nn(row("部門ID", dv))
            strSql &= vbCrLf & "," & kc.SQ(row("購入者", dv)) '必須
            strSql &= vbCrLf & "," & kc.SQ(row("品名", dv)) '必須
            strSql &= vbCrLf & "," & kc.nn(row("型式", dv))
            strSql &= vbCrLf & "," & kc.nz(row("数量", dv)) '空白の場合は0 2021/10/7 見積の数量と見積単価は空白可になった
            strSql &= vbCrLf & "," & kc.nz(row("予算単価", dv)) '空白の場合は0
            strSql &= vbCrLf & "," & kc.nn(row("仕入先ID", dv))
            strSql &= vbCrLf & "," & kc.SQ((row("仕入先", dv))) '必須
            strSql &= vbCrLf & "," & kc.SQ((row("支払区分", dv))) '必須
            strSql &= vbCrLf & "," & row("科目ID", dv) '必須
            strSql &= vbCrLf & "," & kc.nn(row("購入理由", dv))
            If kc.nz(row("予算単価", dv)) > 0 Then
                strSql &= vbCrLf & ",'発注希望'"
            ElseIf kc.nz(row("予算単価", dv)) = 0 Then
                strSql &= vbCrLf & ",'見積希望'"
            End If
            strSql &= vbCrLf & ")"
            'データベース接続、トランザクション開始（行ロックする）
            tran = dCon.Connection.BeginTransaction
            If dCon.ExecuteSqlMW(tran, strSql) = False Then
                tran.Rollback()
                MsgBox("UPDATE失敗")
                Exit Sub
            Else
                tran.Commit()
                SwEditSave = False
                MsgBox("見積へのコピーに成功しました")
            End If
            tran.Dispose()
        End If
    End Sub


    Private Sub btn削除_Click(sender As Object, e As EventArgs) Handles btn削除.Click
        Dim i削除行No As Integer = DGV製品.SelectedCells(0).RowIndex
        Dim strSql As String = ""
        Dim tran As System.Data.SqlClient.SqlTransaction = Nothing
        Dim Response As String
        Dim msg As String = ""

        '最終行をRemoveAtで消そうとすると「コミットされていない新しい行を削除することはできません。」というエラーが出てしまった。(Deleteキーの場合だと最終行削除できないよう制御されている)
        'Datatableの行数とDataGridviewの行数で最終行かどうか判定することにした
        If i削除行No + 1 = DGV製品.Rows.Count Then
            'DataGridViewの一番下の新規行(自動で作られる行)は削除できない
            MsgBox("新規行は削除できません")
            Exit Sub

        End If

        DGV製品.Rows(i削除行No).Selected = True
        If kc.nz(DGV製品.Rows(i削除行No).Cells("UID").Value) = 0 Then
            'まだデータベースに登録されていないから見た目上行を消したら終わり
            Response = MsgBox("選択されている行を削除してよろしいですか？", MsgBoxStyle.YesNo, "確認")
            If Response = vbYes Then
                DGV製品.Rows.RemoveAt(i削除行No)
            End If
        Else
            '2022/02/08 ekawai del Main_Formの削除で選択行+一番下の行が消える不具合があった。製品と見積もりの削除ボタンの動きには問題なかったが、
            '「通常削除はその場で反映させるほうがよい」とのことなので、仮削除はやめて即削除するように変えた。削除履歴は残さないでいいことを確認済み
            'Response = MsgBox("このデータを仮削除してよろしいですか？" & _
            '                  vbCrLf & "※表示上消すだけで、実際には削除されていません。" & _
            '                  vbCrLf & "仮削除後、更新ボタンを押すか保存をすると完全に削除されます", MsgBoxStyle.YesNo, "確認")

            'If Response = vbYes Then

            '    DGV製品.Rows.RemoveAt(i削除行No)
            '    SwEditSave = True
            'End If
            '2022/02/08 ekawai del
            Response = MsgBox("選択されている行を削除してよろしいですか？", MsgBoxStyle.YesNo, "確認")
            If Response = vbYes Then

                strSql = "DELETE FROM TM_製品"
                strSql &= vbCrLf & "WHERE UID = " & DGV製品.Rows(i削除行No).Cells("UID").Value
                tran = dCon.Connection.BeginTransaction
                If dCon.ExecuteSqlMW(tran, strSql) = False Then
                    tran.Rollback()
                    tran.Dispose()

                    MsgBox("削除失敗")
                    Exit Sub

                End If

                tran.Commit()
                tran.Dispose()

                DGV製品_Read()

                MsgBox("削除成功")

            End If


        End If
        SwEditSave = False
    End Sub
    Public Function 製品更新() As Boolean
        'テーブルスタイルの取得

        'DataRowViewを使いDataGridViewの現在の行（または任意の行）からソース元のDataTableのDataRowを取得します。
        'Dim dgr As System.Windows.Forms.DataGridViewRow = Me.DGV購入.CurrentRow
        'Dim drv As System.Data.DataRowView = CType(dgr.DataBoundItem, System.Data.DataRowView)
        'Dim dr As System.Data.DataRow = CType(drv.Row, System.Data.DataRow)
        'Dim tran As System.Data.SqlClient.SqlTransaction = Nothing
        Dim dt As DataTable = DGV製品.DataSource
        Dim iエラー行 As Integer
        Dim 更新対象Flg As String
        Dim strSql As String = ""

        Dim dv As DataRowVersion

        Dim tran As System.Data.SqlClient.SqlTransaction = Nothing

        If dt.Rows.Count = DGV製品.Rows.Count Then
            '最終行の仕入先をクリアしたあとにコンボ選び直した時とかに、空行がdtの行としてカウントされてしまうことがある。空白行だから絶対エラーチェックにひっかかる。
            'だから、更新チェックの対象にならないように削除する
            dt.Rows(dt.Rows.Count - 1).Delete()
        End If

        更新対象Flg = False
        '更新前のエラーチェック
        If dt.Rows.Count > 0 Then
            'If dt.Rows.Count = DGV製品.Rows.Count Then
            '    '空行がdtの行としてカウントされてしまうことがある。仕入先のセル3回クリックした時とか。
            '    '更新チェックの対象にしたくないし後のループにも入れたくないので削除する
            '    dt.Rows(dt.Rows.Count - 1).Delete()
            'End If

            For Each row As DataRow In dt.Rows
                dv = DataRowVersion.Current
                Select Case row.RowState
                    Case DataRowState.Unchanged
                        '変更なし⇒次へ進む 
                        Continue For
                    Case DataRowState.Modified, DataRowState.Added
                        '更新 or 追加　⇒エラーチェック
                        iエラー行 = Main_Form.必須項目Check(dt.Rows.IndexOf(row) + 1, row, dv, "製品", "")
                        If iエラー行 > 0 Then
                            'エラー行にスクロール
                            DGV製品.FirstDisplayedScrollingRowIndex = iエラー行 - 1
                            DGV製品.Rows(iエラー行 - 1).Selected = True
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
                            '更新の場合
                            '更新前のチェック
                            'Check_DGV(row, dv, "製品")
                            strSql = "UPDATE TM_製品"
                            strSql &= vbCrLf & "SET"
                            strSql &= vbCrLf & "部署 = " & kc.SQ(Select_Form.s部署)
                            strSql &= vbCrLf & ",品名 = " & kc.SQ(row("品名", dv)) '必須
                            strSql &= vbCrLf & ", 部門ID = " & kc.nn(row("部門ID", dv))
                            strSql &= vbCrLf & ", 購入者 = " & kc.nn(row("購入者", dv))
                            strSql &= vbCrLf & ", 型式 = " & kc.nn(row("型式", dv))
                            strSql &= vbCrLf & ", 予算単価 = " & kc.nn(row("予算単価", dv))
                            strSql &= vbCrLf & ", 数量 = " & kc.nn(row("数量", dv))
                            '科目ID(int型) 必須
                            strSql &= vbCrLf & ", 科目ID = " & row("科目ID", dv)
                            strSql &= vbCrLf & ", 仕入先ID = " & kc.nn(row("仕入先ID", dv))
                            strSql &= vbCrLf & ", 仕入先 = " & kc.SQ(row("仕入先", dv)) '必須
                            strSql &= vbCrLf & ", 支払区分 = " & kc.SQ(row("支払区分", dv)) '必須
                            strSql &= vbCrLf & ", 購入理由 = " & kc.nn(row("購入理由", dv))
                            strSql &= vbCrLf & ", 登録番号 = " & kc.nn(row("登録番号", dv))
                            strSql &= vbCrLf & ", 備考 = " & kc.nn(row("備考", dv))

                            strSql &= vbCrLf & "WHERE UID = " & row("UID", dv)

                            If dCon.ExecuteSqlMW(tran, strSql) = False Then
                                tran.Rollback()
                                'エラー行にスクロール
                                DGV製品.FirstDisplayedScrollingRowIndex = dt.Rows.IndexOf(row)
                                DGV製品.Rows(dt.Rows.IndexOf(row)).Selected = True

                                MsgBox(dt.Rows.IndexOf(row) + 1 & "行目 UPDATE失敗")
                                Return False
                            Else
                                更新対象Flg = True
                            End If

                        Case DataRowState.Added

                            '新規行の場合


                            '追加前のチェック
                            'Check_DGV(row, dv, "製品")
                            strSql = "INSERT INTO"
                            strSql &= vbCrLf & "TM_製品("
                            strSql &= vbCrLf & "部署"
                            strSql &= vbCrLf & ",品名"
                            strSql &= vbCrLf & ",部門ID"
                            strSql &= vbCrLf & ",購入者"
                            strSql &= vbCrLf & ",型式"
                            strSql &= vbCrLf & ",予算単価"
                            strSql &= vbCrLf & ",数量"
                            strSql &= vbCrLf & ",科目ID"
                            strSql &= vbCrLf & ",仕入先ID"
                            strSql &= vbCrLf & ",仕入先"
                            strSql &= vbCrLf & ",支払区分"
                            strSql &= vbCrLf & ",購入理由"
                            strSql &= vbCrLf & ",登録番号"
                            strSql &= vbCrLf & ",備考"
                            strSql &= vbCrLf & ") VALUES ("
                            strSql &= vbCrLf & kc.SQ(Select_Form.s部署)
                            strSql &= vbCrLf & "," & kc.SQ(row("品名", dv)) '必須
                            strSql &= vbCrLf & "," & kc.nn(row("部門ID", dv))
                            strSql &= vbCrLf & "," & kc.nn(row("購入者", dv))
                            strSql &= vbCrLf & "," & kc.nn(row("型式", dv))
                            strSql &= vbCrLf & "," & kc.nn(row("予算単価", dv)) 'NULL許可
                            strSql &= vbCrLf & "," & kc.nn(row("数量", dv)) 'NULL許可
                            strSql &= vbCrLf & "," & row("科目ID", dv) '必須
                            strSql &= vbCrLf & "," & kc.nn(row("仕入先ID", dv))
                            strSql &= vbCrLf & "," & kc.SQ(row("仕入先", dv))
                            strSql &= vbCrLf & "," & kc.SQ(row("支払区分", dv))
                            strSql &= vbCrLf & "," & kc.nn(row("購入理由", dv))
                            strSql &= vbCrLf & "," & kc.nn(row("登録番号", dv))
                            strSql &= vbCrLf & "," & kc.nn(row("備考", dv))
                            strSql &= vbCrLf & ")"



                            If dCon.ExecuteSqlMW(tran, strSql) = False Then
                                tran.Rollback()
                                'エラー行にスクロール
                                DGV製品.FirstDisplayedScrollingRowIndex = dt.Rows.IndexOf(row)
                                DGV製品.Rows(dt.Rows.IndexOf(row)).Selected = True
                                MsgBox(dt.Rows.IndexOf(row) + 1 & "行目 INSERT失敗")
                                Return False
                            End If
                            更新対象Flg = True

                            '削除行
                            '2022/02/08 btn削除_Clickに移動
                            'Case DataRowState.Deleted
                            '更新対象Flg = True
                            ''TM_製品にidがあれば削除する
                            ''Deleted 行には Current 行バージョンがないため、列値にアクセスするときに DataRowVersion.Original を渡す必要があります。
                            'strSql = "DELETE FROM TM_製品"
                            'strSql &= vbCrLf & "WHERE UID = " & row("UID", DataRowVersion.Original)

                            'If dCon.ExecuteSqlMW(tran, strSql) = False Then
                            '    tran.Rollback()
                            '    MsgBox("id=" & row("id", DataRowVersion.Original) & " 削除失敗")
                            '    Return False
                            'End If


                    End Select
                Catch ex As Exception
                    MsgBox(ex.Message & "更新失敗")
                    If tran IsNot Nothing Then
                        tran.Rollback()
                    End If
                    Return False
                End Try

                'End If

            Next

            If 更新対象Flg = True Then
                tran.Commit()
                dt.AcceptChanges() 'Rowstateをunchangedに更新する
                MsgBox("更新成功")

            Else
                MsgBox("更新対象がありません")


            End If
            SwEditSave = False
            tran.Dispose()
            Return True
        End If
    End Function

    Private Sub btn製品検索_Click(sender As Object, e As EventArgs) Handles btn製品検索.Click
        Disp検索結果()
    End Sub

    Sub Disp検索結果()
        Dim i As Integer
        Dim sValue列 As String
        Dim strSql As String
        Dim str検索型 As String
        検索条件 = ""
        検索Flg = False
        日付検索Flg = False

        For i = 1 To 3
            sValue列 = CType(Me.Controls("cmb検索項目" & i), ComboBox).SelectedValue
            If CType(Me.Controls("cmb検索項目" & i), ComboBox).SelectedIndex >= 0 _
                And CType(Me.Controls("cmb記号" & i), ComboBox).SelectedIndex >= 0 And CType(Me.Controls("txt検索条件" & i), TextBox).Text <> "" Then
                検索Flg = True
                strSql = ""
                '科目IDではなくて科目名で検索できるようにするために新しく製品用データ列を作った
                'strSql = "SELECT 型 FROM TM_列名変換 WHERE 表示列 = '" & CType(Me.Controls("cmb検索項目" & i), ComboBox).SelectedValue & "'"
                strSql = "SELECT 型 FROM TM_列名変換 WHERE 製品用データ列 = '" & CType(Me.Controls("cmb検索項目" & i), ComboBox).SelectedValue & "'"
                Dim DT As DataTable = dCon.DataSet(strSql, "DT").Tables(0)
                str検索型 = DT.Rows(0)("型")
                Select Case str検索型

                    Case "文字"
                        If CType(Me.Controls("cmb記号" & i), ComboBox).SelectedItem.ToString = "を含む" Then
                            検索条件 = 検索条件 & vbCrLf & "AND " & sValue列 & " LIKE '%" & CType(Me.Controls("txt検索条件" & i), TextBox).Text & "%'"
                        Else
                            検索条件 = 検索条件 & vbCrLf & "AND " & sValue列 & " " & CType(Me.Controls("cmb記号" & i), ComboBox).SelectedItem.ToString & " " & kc.SQ(CType(Me.Controls("txt検索条件" & i), TextBox).Text)
                        End If
                    Case "数値"
                        検索条件 = 検索条件 & vbCrLf & "AND " & sValue列 & " " & CType(Me.Controls("cmb記号" & i), ComboBox).SelectedItem.ToString & " " & CType(Me.Controls("txt検索条件" & i), TextBox).Text

                End Select
            End If
        Next


        If SwEditSave = True Then

            If MsgBox("保存してよろしいですか?" & vbCrLf & "「いいえ」を押すと更新していないデータは消えます", vbYesNo + vbDefaultButton1) = vbYes Then
                If 製品更新() = False Then
                    '更新失敗したら画面を書き換えずに終了
                    Exit Sub
                End If
            Else
                '本来は更新成功しないとSqEditSave = Trueにならないが、この場合は強制的にFalseにする(そうしないと更新するまで画面切り替える度、永遠に更新しますか？というメッセージが出てしまうため)
                SwEditSave = False
            End If

        End If
        DGV製品_Read()

    End Sub
    Private Sub cmb検索項目_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb検索項目1.SelectedIndexChanged, cmb検索項目3.SelectedIndexChanged, cmb検索項目2.SelectedIndexChanged
        Dim strSql As String
        Dim s型 As String
        Dim cmb項目 As ComboBox = CType(sender, ComboBox)
        Dim i As Integer
        Dim dt選択肢 As DataTable
        i = Strings.Right(cmb項目.Name, 1)
        CType(Me.Controls("cmb検索条件選択" & i), ComboBox).Items.Clear()
        If cmb項目.SelectedIndex >= 0 Then
            strSql = ""
            '科目名で検索できるようにしたので、TM_製品の列名そのまま使えなくなったので、新たにTM_列名変換に製品用データ列を追加した
            'strSql = "SELECT * FROM TM_列名変換 WHERE 表示列 = '" & cmb項目.SelectedValue & "'"
            strSql = "SELECT * FROM TM_列名変換 WHERE 製品用データ列 = '" & cmb項目.SelectedValue & "'"

            Dim DT As DataTable = dCon.DataSet(strSql, "DT").Tables(0)
            If DT.Rows.Count = 1 Then
                s型 = DT.Rows(0)("型")
                CType(Me.Controls("cmb記号" & i), ComboBox).Items.Clear()
                CType(Me.Controls("cmb記号" & i), ComboBox).Items.Add("=")
                CType(Me.Controls("cmb記号" & i), ComboBox).Items.Add("<>")

                If s型 = "数値" Then
                    CType(Me.Controls("cmb記号" & i), ComboBox).Items.Add(">=")
                    CType(Me.Controls("cmb記号" & i), ComboBox).Items.Add("<=")
                    CType(Me.Controls("cmb記号" & i), ComboBox).Items.Add(">")
                    CType(Me.Controls("cmb記号" & i), ComboBox).Items.Add("<")

                    CType(Me.Controls("txt検索条件" & i), TextBox).ImeMode = Windows.Forms.ImeMode.Disable
                Else
                    '数値以外の時
                    CType(Me.Controls("cmb記号" & i), ComboBox).Items.Add("を含む")
                    CType(Me.Controls("txt検索条件" & i), TextBox).ImeMode = Windows.Forms.ImeMode.Hiragana
                    dt選択肢 = makeSelectList(DT.Rows(0)("表示列"))
                    For Each r In dt選択肢.Rows
                        CType(Me.Controls("cmb検索条件選択" & i), ComboBox).Items.Add(r(DT.Rows(0)("表示列")))
                    Next

                End If
                CType(Me.Controls("cmb記号" & i), ComboBox).SelectedIndex = -1
                CType(Me.Controls("txt検索条件" & i), TextBox).Text = ""
                CType(Me.Controls("txt検索条件" & i), TextBox).Enabled = True
            Else
                MsgBox("TM_列名変換テーブルに型が登録されていません")
                cmb項目.SelectedIndex = -1
            End If
        End If

    End Sub
    Private Function makeSelectList(ColName As String) As DataTable
        Dim dt検索用リスト As DataTable
        Dim dtView As DataView
        dtView = New DataView(DGV製品.DataSource)
        dtView.Sort = ColName & " ASC"
        dt検索用リスト = dtView.ToTable(True, ColName)
        Return dt検索用リスト
    End Function

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        For i = 1 To 3
            CType(Me.Controls("cmb検索項目" & i), ComboBox).SelectedIndex = -1
            CType(Me.Controls("cmb記号" & i), ComboBox).SelectedIndex = -1
            CType(Me.Controls("txt検索条件" & i), TextBox).Text = ""
            CType(Me.Controls("txt検索条件" & i), TextBox).Enabled = False
        Next
        検索Flg = False
        If SwEditSave = True Then

            If 製品更新() = False Then
                '更新失敗したら画面を書き換えずに終了
                Exit Sub
            End If

        End If

        DGV製品_Read()

    End Sub

    Private Sub txt検索条件_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt検索条件1.KeyPress, txt検索条件3.KeyPress, txt検索条件2.KeyPress
        Dim strSql As String
        Dim str検索型 As String
        Dim i As Integer
        i = Strings.Right(CType(sender, TextBox).Name, 1)
        'メインフォームの時はデータ列だけど、製品の時はForm_Load時にValueMemberに表示列入れているので注意！
        strSql = "SELECT 型 FROM TM_列名変換 WHERE 表示列 = '" & CType(Me.Controls("cmb検索項目" & i), ComboBox).SelectedValue & "'"
        Dim DT As DataTable = dCon.DataSet(strSql, "DT").Tables(0)
        str検索型 = DT.Rows(0)("型")

        If str検索型 = "数値" Then
            '0～9とバックスペース、Enter以外が押された場合はイベントをキャンセルする

            If (e.KeyChar < "0"c OrElse "9"c < e.KeyChar) AndAlso _
                e.KeyChar <> ControlChars.Back AndAlso e.KeyChar <> Microsoft.VisualBasic.ChrW(Keys.Enter) Then

                MsgBox("数字以外は入力できません")
                e.Handled = True
                Exit Sub
            End If
        End If
        '型関係なく、Enterだったら検索結果を表示する
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Enter) Then 'エンターキーが押されたら
            If CType(Me.Controls("txt検索条件" & i), TextBox).Text <> "" Then
                Disp検索結果()
            End If
        End If

    End Sub

    Sub コンボイベントハンドラ削除() 'SelectedIndexChangedが何回も発生してStackOverflowでプログラムが落ちる対策
        RemoveHandler Me.dataGridViewComboBox.SelectedIndexChanged, AddressOf dataGridViewComboBox_SelectedIndexChanged
    End Sub
    Sub コンボイベントハンドラ追加()
        AddHandler Me.dataGridViewComboBox.SelectedIndexChanged, AddressOf dataGridViewComboBox_SelectedIndexChanged
    End Sub

    Private Sub cmb検索条件選択_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb検索条件選択1.SelectedIndexChanged, cmb検索条件選択2.SelectedIndexChanged, cmb検索条件選択3.SelectedIndexChanged
        Dim cmb項目 As ComboBox = CType(sender, ComboBox)
        Dim i As Integer
        i = Strings.Right(cmb項目.Name, 1)

        CType(Me.Controls("txt検索条件" & i), TextBox).Text = CType(Me.Controls("cmb検索条件選択" & i), ComboBox).SelectedItem.ToString

    End Sub
    '2022/07/25 ekawai add S
    '技術部から似たような製品を登録するから、製品シートにも行コピーボタンが欲しいという要望があった。同じ製品を2行登録することはないので、書き換える前提の行コピーよりも縦貼り付けのほうがいいのではということで、
    'メインフォームから下記コードをコピーした

    Private Sub btn貼り付け_Click(sender As Object, e As EventArgs) Handles btn貼り付け.Click

        '開発当初はスピード重視でいいということだったので、複数行への貼り付けは許可しない前提で作っていたがユーザーテスト後に方針が変わって、結局複数行貼り付け機能を追加することになってしまった。2021/10/12
        'キーボード入力の文字に対しては、文字の種類チェックや文字数チェックなど行っているので、貼り付けに対しても同じように更新前にチェックが必要かと思ったが
        'この機能使う人は少数派だろうから、更新時にエラーが出て何行目失敗って分かればそれでいいそうなので、更新前のチェックは行っていない。
        '数字のセルに文字を貼り付けたり、コンボボックスに貼り付けしようとしてもできないようにはなっている。(特別対応したわけではないが既存のコードのおかげで？制限できている)
        '読取専用のセルが貼り付け範囲に含まれる場合は貼り付けできないようにしている
        '既にデータがある行に貼り付ける場合の確認はしていない。即上書きする

        'Dim iTargetRow As Integer = DGV購入品入力.CurrentCellAddress.Y
        'Dim iTargetCol As Integer = DGV購入品入力.CurrentCellAddress.X

        Dim iTargetRow As Integer = Active行
        Dim iTargetCol As Integer = Active列
        Dim DGV行Cnt As Integer = DGV製品.Rows.Count

        'クリップボードの内容を取得
        Dim sClipText As String = Clipboard.GetText

        '改行を変換　改行コードLFはセル内改行なので、半角スペースに置き換える。
        'またセル内改行をクリップボードに貼り付けると勝手にダブルクォーテーションで囲まれるので、ダブルクォーテーションChr(34)も置き換えている
        '※頑張ったらクリップボードに貼り付けたときのダブルクォーテーションか人が入れたものか判定出来そうだけど、今のところそんな需要あるか分からないので、単純にダブルクォーテーションは一律空白に置換している
        sClipText = sClipText.Replace(vbCrLf, vbCr)
        'sClipText = sClipText.Replace(vbCr, vbLf)
        sClipText = sClipText.Replace(vbLf, " ")
        sClipText = sClipText.Replace(Chr(34), "")

        '改行でコピー
        Dim lines() As String = sClipText.Split(vbCr)
        Dim cols() As String = lines(0).Split(vbTab)
        Dim colMaxIndex As Integer = cols.GetUpperBound(0)
        Dim rowMaxIndex As Integer = lines.GetUpperBound(0) - 1
        Dim strArr(,) As String


        'テキストボックス内でコピーした時と
        'セルを選択肢してコピーした時で配列の数が変わる
        Dim iLoop As Integer
        '2022/02/14 ekawai add S
        If DGV製品.GetCellCount(DataGridViewElementStates.Selected) > 1 Then
            MsgBox("複数セルを選択している場合は貼り付けできません")
            Exit Sub
        End If
        '2022/02/14 ekawai add E

        If "".Equals(lines(lines.GetLength(0) - 1)) Then
            'セル選択だと余計な空白の配列ができてしまうのでマイナス1している
            iLoop = lines.GetLength(0) - 1
        Else
            iLoop = 1
        End If

        ReDim strArr(rowMaxIndex, colMaxIndex)
        For rowindex As Integer = 0 To rowMaxIndex
            cols = lines(rowindex).Split(vbTab)
            For colindex As Integer = 0 To colMaxIndex
                strArr(rowindex, colindex) = cols(colindex)
            Next

        Next

        Dim i貼り付け行数 As Integer
        Dim i貼り付け列数 As Integer
        '貼り付け前のチェック
        For i貼り付け行数 = 1 To iLoop

            For i貼り付け列数 = 1 To colMaxIndex + 1
                '読取専用の行には貼り付けできないようにする
                If DGV製品(iTargetCol + i貼り付け列数 - 1, iTargetRow + i貼り付け行数 - 1).ReadOnly = True Then
                    MsgBox("貼り付け範囲内にロックされているセルがあるため貼り付けできません")
                    Exit Sub
                End If

                If TypeOf DGV製品(iTargetCol + i貼り付け列数 - 1, iTargetRow + i貼り付け行数 - 1) Is DataGridViewComboBoxCell Then
                    MsgBox("貼り付け範囲内にコンボボックスがあるため貼り付けできません")
                    Exit Sub
                End If
            Next
            '新規行ならRows.Addする
            If DGV行Cnt - 1 <= iTargetRow + i貼り付け行数 - 1 Then

                Dim dt As DataTable = DGV製品.DataSource
                Dim row As DataRow = dt.NewRow

                dt.Rows.Add(row)
            End If

        Next
        Dim val As String
        For i貼り付け行数 = 1 To iLoop
            '空白になったら抜ける
            'If r = lines.GetLength(0) - 1 And "".Equals(lines(r)) Then
            '    Exit For
            'End If
            For i貼り付け列数 = 1 To colMaxIndex + 1
                val = strArr(i貼り付け行数 - 1, i貼り付け列数 - 1).ToString
                DGV製品(iTargetCol + i貼り付け列数 - 1, iTargetRow + i貼り付け行数 - 1).Value = val
                '仕入先の場合、仕入先IDと科目が自動で選ばれる必要がある
                DGVテキストボックス変更時処理(iTargetRow + i貼り付け行数 - 1, iTargetCol + i貼り付け列数 - 1)
            Next
        Next

    End Sub

    Sub DGVテキストボックス変更時処理(iRow As Integer, iCol As Integer)

        Dim DR As DataRow()

        Select Case DGV製品.Columns(iCol).Name
            Case "科目番号"
                If kc.nz(DGV製品(iCol, iRow).Value) <> 0 Then
                    Dim i科目ID As Integer
                    '科目をリセットする
                    DGV製品("科目", iRow).Value = DBNull.Value
                    i科目ID = kc.nz(DGV製品.Rows(iRow).Cells("科目番号").Value)
                    '科目番号に応じた科目コンボを選択する
                    DR = Main_Form.DT現科目.Select("id = " & i科目ID)
                    If DR.Length = 1 Then
                        DGV製品("科目", iRow).Value = DR(0)("id")
                    Else
                        DGV製品(iCol, iRow).Value = DBNull.Value
                    End If
                End If

                '仕入先マスタに存在する仕入先名を入力したら、仕入先IDと支払区分が自動で入る
                '製品マスタに登録している行を購入品入力にコピーした場合、製品マスタに登録されている支払区分がコピーした時点では入るが
                'あとで仕入先のセルを触ると仕入先マスタに登録されている支払区分に書き変わってしまうという指摘があったのでITで方針を話し合った
                '→今のままでいい。1つの仕入先に支払区分が複数あるはずない。経理が全ての仕入先に支払区分登録すればよいとのこと　2022.02.17
            Case "仕入先"
                Dim str仕入先名 As String
                str仕入先名 = kc.ns(DGV製品("仕入先", iRow).Value)
                GET仕入先ID(str仕入先名, iRow)

        End Select


    End Sub
    '2022/07/13 ekawai add E

    '2022/10/4 ekawai add 鈴木GMから三吉さんから製品マスタのExcel出力機能がほしいというリクエストがあったということで機能追加したが、
    '翌日、本人から連絡が来て、1回全体のリストが欲しかっただけで、このボタンは要らなかったらしい。
    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click

        Dim dsOutput As DataSet     '出力リストのいれもの
        'DataSetに取得
        dsOutput = dCon.DataSet(DGV製品SQL, "t_エクセル出力")

        '以下よりExcelへ転送する
        Dim oExcel As Excel.Application
        Dim oBook As Excel.Workbook
        Dim oSheet As Excel.Worksheet
        Dim oRange As Excel.Range

        'Start a new workbook in Excel.
        oExcel = CreateObject("Excel.Application")
        oBook = oExcel.Workbooks.Add

        'Datasetの行、列数分だけの２次元配列を作成する
        '２次元配列宣言
        With dsOutput.Tables("t_エクセル出力")
            Dim DataArray(.Rows.Count, .Columns.Count) As Object
            Dim i As Integer    'ループ変数

            '配列にDataTableの中身をいれる
            For i = 0 To .Rows.Count - 1
                For k = 0 To .Columns.Count - 1
                    DataArray(i, k) = .Rows(i)(k)
                Next
            Next
            oSheet = oBook.Sheets(1)
            oSheet.Name = "データ"

            'シートの１行目に列名を表示する
            For j As Integer = 0 To .Columns.Count - 1
                oSheet.Cells(1, j + 1).Value = .Columns(j).ColumnName
            Next
            'セルA2に配列を転送する（貼り付ける）
            oRange = oSheet.Range("A2")
            oRange.Resize(.Rows.Count, .Columns.Count).Value = DataArray

            '以下、エクセルで加工処理
            oSheet.Cells.EntireColumn.AutoFit()                                 'Sheet1列幅最適化
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape  'Sheet1印刷横向き

        End With

        'エクセルを表示する
        oExcel.Visible = True

        '終了処理
        oRange = Nothing
        oSheet = Nothing
        oBook = Nothing
        oExcel = Nothing
        GC.Collect()




    End Sub
    Function DGV製品SQL() As String
        Dim strSql As String = ""
        '分岐させるの手間だから部署ごとに表示項目を変える必要はないそうなので、棚番号などの部署独自の項目もあえて全部署表示しています
        strSql = strSql & vbCrLf & "SELECT TM_製品.UID,TM_製品.登録番号,TM_製品.部署,"
        strSql = strSql & vbCrLf & "TM_製品.部門ID,TM_division.division as 部門,TM_製品.購入者,TM_製品.品名,"
        strSql = strSql & vbCrLf & "TM_製品.型式,TM_製品.数量,TM_製品.予算単価,TM_製品.仕入先ID,TM_製品.仕入先,"
        strSql = strSql & vbCrLf & "TM_製品.支払区分,TM_製品.科目ID,TM_kamoku.kamoku as 科目,TM_製品.購入理由,TM_製品.備考"
        strSql = strSql & vbCrLf & "FROM TM_製品"
        strSql = strSql & vbCrLf & "INNER JOIN TM_kamoku ON TM_製品.科目ID = TM_kamoku.id"
        strSql = strSql & vbCrLf & "INNER JOIN TM_division ON TM_製品.部門ID = TM_division.id"
        strSql = strSql & vbCrLf & "WHERE TM_製品.部署 = " & kc.SQ(Select_Form.s部署)

        If 検索Flg Then

            strSql = strSql & vbCrLf & 検索条件

        Else

        End If

        strSql = strSql & vbCrLf & "ORDER BY TM_製品.UID"

        Return strSql
    End Function
    '2022/10/4 ekawai end 品質から製品マスタのExcel出力機能がほしいというリクエストがあったため追加
End Class