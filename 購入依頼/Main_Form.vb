Option Explicit On
Imports Microsoft.Office.Interop
Public Class Main_Form
    Private dCon As New merrweth_init_DbConnection
    Public kc As New 共通処理_Class
    Private dataGridViewComboBox As DataGridViewComboBoxEditingControl = Nothing
    Private searchFld As nameValue = New nameValue
    'Private CellEditStart As Boolean = False　使ってなさそうだからコメントアウト
    Private arr製品情報(11) As Object
    Private arr見積情報(11) As Object
    Dim f1 As Form
    Private c日付セル As DataGridViewCell
    Private frm製品 As 製品_Form
    'Form1オブジェクトを保持するためのフィールド
    'Private Shared _form1Instance As Main_Form
    'Private DGrVTBEC As DataGridViewTextBoxEditingControl
    Private DT購入品リスト As DataTable
    Private 検索Flg As Boolean = False
    Private 日付検索Flg As Boolean = False
    Private sFileType As String
    Private 検索条件 As String
    Private 日付検索条件 As String
    'Private s型 As String

    Private 最終更新時刻 As Date
    Private CloseFlg As Boolean '右上の閉じるボタンを押した時にTRUEにする

    Public SwEditSave As Boolean = False
    Public 変更前index As Integer
    Public 変更前Value As String
    Public DT現科目 As DataTable
    Public DT現ワークコード As DataTable
    Public DT部署別科目 As DataTable
    Public DT部署別仕入先 As DataTable
    Public DTワークコード As DataTable
    Public DT部署別購入者 As DataTable
    Public DT部門 As DataTable
    Public Active行 As Integer
    Public Active列 As Integer
    Public EventFlg As Boolean 'コンボボックスのイベント立てたことを覚えておくフラグ
    Public Flg再表示 As Boolean 'DGV購入品入力_表示の時はCellValidatingとRowAddedイベント発生しないようにするためのフラグ
    Private strScroll As String = "Left" 'スクロール用フラグ
    Public Flg印刷対象有 As Boolean = False '1回でもダイヤログ開いたら印刷対象ありと判定する
    Public Const cVer As String = "221207-1"

    Private Sub btn見積_Click(sender As Object, e As EventArgs) Handles btn見積.Click
        'Dim f2 As Form
        'If My.Application.OpenForms("見積_form") IsNot Nothing Then
        '    My.Application.OpenForms("見積_form").WindowState = FormWindowState.Normal
        '    My.Application.OpenForms("見積_form").Activate()
        'Else
        '    'AddressOf 演算子は、プロシージャ名に適用されるときにデリゲート オブジェクトを返します。
        '    f2 = New 見積_form(AddressOf Me.値転送)
        '    f2.Show()
        'End If
        Dim f2 As New 見積_form
        f2.Show()
    End Sub

    Private Sub btn製品_Click(sender As Object, e As EventArgs) Handles btn製品.Click
        '製品_Formクラスのインスタンスを作成する
        'frm製品 = New 製品_Form()
        'frm製品.f1 = Me
        ''モードレスフォームとして表示する,所有者をMain_formとする
        'frm製品.Show()

        If My.Application.OpenForms("製品_form") IsNot Nothing Then
            My.Application.OpenForms("製品_form").WindowState = FormWindowState.Normal
            My.Application.OpenForms("製品_form").Activate()
        Else
            f1 = New 製品_Form(AddressOf Me.値転送)
            f1.Show()
        End If

    End Sub

    Private Sub 値転送(ByVal arrOriginal() As Object, Form As String)

        Dim tran As System.Data.SqlClient.SqlTransaction = Nothing
        Select Case Form
            Case "製品"

                For i As Integer = 0 To 11
                    arr製品情報(i) = arrOriginal(i)
                Next

                Set製品情報()

                'Case "見積"
                '    For i As Integer = 0 To 11
                '        arr見積情報(i) = arrOriginal(i)
                '    Next

                '    Set見積情報()


        End Select

        DGV購入品入力.FirstDisplayedScrollingRowIndex = DGV購入品入力.Rows.Count - 1

    End Sub


    Private Sub Main_Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '右上の閉じるボタンから閉じられるとき
        If e.CloseReason = CloseReason.UserClosing Then
            If SwEditSave Then
                'If MsgBox("保存せずに終了しますか？", vbYesNo + vbDefaultButton1) = vbNo Then
                If MsgBox("保存してよろしいですか？" & vbCrLf & "「いいえ」を押すと更新していないデータは消えます", vbYesNo + vbDefaultButton1) = vbYes Then
                    'If メイン更新() = False Then
                    '    '更新失敗したら閉じない
                    '    e.Cancel = True
                    'End If

                    '何かしらボタンを押さないと変更した行がModifiedにならないから、直接メイン更新プロシージャを呼び出さず、更新ボタンを押したように見せかけることにした。
                    btn更新.PerformClick()
                    If CloseFlg Then
                        Me.Close() 'Main_Formの場合、更新後は何故か勝手にフォームが閉じなかったからCloseメソッド書いている
                    Else
                        '万一更新が失敗した場合はフォームを閉じない…内容を修正して再度更新するか更新しないで閉じるかユーザーに選ばせる
                        e.Cancel = True
                    End If


                Else
                    SwEditSave = False

                End If
            End If
        End If
    End Sub


    Private Sub Main_Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        '行の数だけ発生してしまうのでフォームLoad終わるまでは停止しておく
        RemoveHandler DGV購入品入力.CellEnter, AddressOf DGV購入品入力_CellEnter
        RemoveHandler DGV購入品入力.CellValidating, AddressOf DGV購入品入力_CellValidating
        RemoveHandler DGV購入品入力.CellValueChanged, AddressOf DGV購入品入力_CellValueChanged
        RemoveHandler DGV購入品入力.RowsAdded, AddressOf DGV購入品入力_RowsAdded

        'Private Sub Main_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Main_Form.MainfrmInstance = Me

        Me.Text = Select_Form.s部署
        lblcatalog.Text = "接続先:" & merrweth_init_DbConnection.strCatalog
        '使用バージョンをフォームに表示
        lblver.Text = "Ver." & cVer

        '///日付のコンボボックスの設定////
        Dim DS日付 As DataSet
        Dim DT変換日付 As New DataTable
        Dim strSql As String
        cmb検索項目日付.DataSource = Nothing
        '
        strSql = ""
        strSql = "SELECT * FROM TM_列名変換 WHERE 型 = '日付' AND 入力 = 1"

        DS日付 = dCon.DataSet(strSql, "変換テーブル2")
        DT変換日付 = DS日付.Tables(0)
        cmb検索項目日付.DisplayMember = "表示列"
        cmb検索項目日付.ValueMember = "データ列"
        cmb検索項目日付.DataSource = DT変換日付

        cmb検索項目日付.SelectedIndex = -1

        'cmb検索項目日付.Text = "入荷日"
        'dtp開始日.Text = Today.AddYears(-2) 'データ数多い部署開くとき時間かかるのでとりあえず初期値は2年分だけ表示することになった
        日付検索Flg = True
        日付検索条件 = "AND insertStamp >= " & Today.AddYears(-2) '入荷日だと日付の検索条件とかぶってしまってややこしいからinsertStampにした



        'マスタフォーム用テーブル作成(コンボボックスの中身とは別)
        '【科目】
        strSql = ""
        strSql = "SELECT TM_kamoku.id, TM_kamoku.kamoku"
        strSql &= vbCrLf & "FROM TM_dep_kamoku INNER JOIN"
        strSql &= vbCrLf & "TM_kamoku ON TM_dep_kamoku.科目ID = TM_kamoku.id"
        strSql &= vbCrLf & "WHERE TM_dep_kamoku.部署ID =" & Select_Form.i部署ID
        DT現科目 = dCon.DataSet(strSql, "DT").Tables(0)

        '【ワークコード】終了(old)かどうかに関わらず全て表示する
        strSql = "select work_code"
        strSql &= vbCrLf & ",class1 as 部門"
        strSql &= vbCrLf & ",class2 as 区分"
        strSql &= vbCrLf & ",department as 部署"
        strSql &= vbCrLf & ",line as ライン"
        strSql &= vbCrLf & ",seiri_no as 整理番号"
        strSql &= vbCrLf & ",facility as 設備"
        strSql &= vbCrLf & ",distribution as 按分"
        strSql &= vbCrLf & ",old as 終了"
        strSql &= vbCrLf & "from TM_seigi_work_list"
        strSql &= vbCrLf & "order by work_code"
        DT現ワークコード = dCon.DataSet(strSql, "DT").Tables(0)

        DGV購入品入力_表示()
        DGV購入品入力_詳細設定()
        DGV購入品入力_表示条件()
        DGV購入品入力_セル設定()
        'ヘッダーを除く表示行の幅に自動調整
        DGV購入品入力.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders
        DGV購入品入力.RowHeadersWidth = 25 '▸が見える最小のサイズに設定
        DGV購入品入力.AutoGenerateColumns = False '列の自動生成禁止 これをしないと列の並び順が変わってしまう
        DGV購入品入力.AllowUserToDeleteRows = False 'Deleteボタンによる行削除禁止(DGVにdatasourceセットする前にFalseにしたらエラーになる)
        'DGV購入品入力.MultiSelect = False '複数選択禁止 →　選択したセルの合計表示するようにしたいから許可することにした



        '///検索項目コンボボックスの共通設定////
        Dim DT変換1 As New DataTable
        Dim DT変換2 As New DataTable
        Dim DT変換3 As New DataTable



        cmb検索項目1.DataSource = Nothing
        cmb検索項目2.DataSource = Nothing
        cmb検索項目3.DataSource = Nothing

        strSql = ""
        strSql = "SELECT * FROM TM_列名変換 WHERE 型 <> '日付' AND 入力 = 1"


        DT変換1 = dCon.DataSet(strSql, "TB1").Tables(0)
        DT変換2 = dCon.DataSet(strSql, "TB2").Tables(0)
        DT変換3 = dCon.DataSet(strSql, "TB3").Tables(0)


        For i = 1 To 3
            RemoveHandler CType(Me.Controls("cmb検索項目" & i), ComboBox).SelectedIndexChanged, AddressOf cmb検索項目_SelectedIndexChanged
            CType(Me.Controls("cmb検索項目" & i), ComboBox).DisplayMember = "表示列"
            CType(Me.Controls("cmb検索項目" & i), ComboBox).ValueMember = "データ列"
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

            CType(Me.Controls("txt検索条件" & i), TextBox).Enabled = False 'コンボ選択前にテキスト入力しようとすると位置０に行がありませんと出る対策
        Next



        'AddHandler DGV購入品入力.CellValueChanged, AddressOf DGV購入品入力_CellValueChanged

        '閲覧権限しかない人には更新ボタンを表示しない
        If Select_Form.更新Flg = False And Select_Form.承認Flg = False And Select_Form.経理Flg = False Then
            btn更新.Visible = False
            btn削除.Visible = False
            btnCOPY.Visible = False
            btn送料.Visible = False
            btn手数料.Visible = False
            btn値引き.Visible = False
        End If

        '個人画面サイズに合わせてフォームの幅変えようと思ったが、フォームサイズは聖域だから触ってはいけないとアドバイスがありやめた
        'Me.StartPosition = FormStartPosition.Manual
        'Me.DesktopLocation = New Point(1, 1)
        'Me.Width = Select_Form.i画面幅 - 20
        'DGV購入品入力.Width = Select_Form.i画面幅 - 60

        '経理の時だけ画面サイズ拡大する(1日数十件請求書チェックするため、画面小さいとスクロール何度もしないといけなくて作業効率が悪い)
        If Select_Form.経理Flg Then
            Me.StartPosition = FormStartPosition.Manual
            Me.DesktopLocation = New Point(1, 1)
            Me.Width = Select_Form.i画面幅 - 20
            DGV購入品入力.Width = Select_Form.i画面幅 - 60
            'TM_facilityに存在するのに仕入先IDが空白の物に対して仕入先名をキーとして仕入先IDを入れる
            '元々全データに対して行えばいいとのことだったが、やはりそれだと無駄なデータが多すぎるのでとりあえず3ヶ月で十分ではということで3ヶ月分にしているだけで特に根拠はない
            '元々誰が開いてもアップデート実行すればいいとのことだったが、ミスを減らすためにも、起動時間短縮のためにも、経理が開く時だけでよいということになった。どの部署を開いているか関係なく行えばいいとのこと
            '【この処理を追加した理由】
            '仕入先名はコンボボックスではないためTM_facilityやTM_dep_facilityに存在していなくても入力が可能
            '経理が仕入先マスタ更新する前に、各部で製品マスタや購入品入力画面で未登録の仕入先を入力すると仕入先IDは空白になってしまう
            '一度更新済みのデータに対しては、コンボボックス再選択しない限りあとから購入ID入れるタイミングがないので、フォーム開くたびにアップデートすることにした
            '2021/04/12 ekawai add S
            Dim DT仕入先ID未取得リスト As DataTable
            strSql = "SELECT TD_po.id,TM_facility.facility_id,TM_facility.kiban"
            strSql &= vbCrLf & "FROM TM_facility INNER JOIN"
            strSql &= vbCrLf & "TD_po ON TM_facility.kiban = TD_po.vendor"
            strSql &= vbCrLf & "WHERE ((TD_po.vendor_id IS NULL) or (TD_po.vendor_id = ''))"
            strSql &= vbCrLf & "AND insertStamp > " & kc.SQ(Today.AddMonths(-3))
            DT仕入先ID未取得リスト = dCon.DataSet(strSql, "DT未取得リスト").Tables(0)
            For Each DR In DT仕入先ID未取得リスト.Rows
                strSql = "UPDATE TD_po SET"
                strSql &= vbCrLf & "vendor_id = " & kc.SQ(DR("facility_id").ToString)
                strSql &= vbCrLf & "WHERE id = " & DR("id").ToString
                dCon.Command()
                dCon.ExecuteSQL(strSql)
            Next
            '2021/04/12 ekawai add E
        End If

        '一括承認ボタン最初は表示しない
        btn一括承認.Visible = False
        btn一括稟議承認.Visible = False
        '印刷ボタン最初は表示しない
        btn注文書印刷.Visible = False
        btn未承認印刷.Visible = False
        'フォーム閉じる時に押させるボタンは非表示
        'Button1.Visible = False

        'グリッドの並び替えを禁止する。(ただしプログラムからの並び替えは可能とする)
        For Each c As DataGridViewColumn In DGV購入品入力.Columns
            c.SortMode = DataGridViewColumnSortMode.Programmatic
        Next c

        If Select_Form.s部署 <> "金型技術部" Then '2024/05/14 ekawai 組織変更に伴い部署名変更

            btn技術DB.Visible = False
            btn出荷データ作成.Visible = False
        End If
        'セル設定でも行っているがLoad時は敢えてプロシージャの最後にも同じことをしている
        'スクロールバーの位置が一番下よりもだいぶ上のほうになってしまうため
        If DGV購入品入力.Rows.Count > 0 Then
            'グリッドに表示される最初（一番上）の行の取得／設定
            DGV購入品入力.FirstDisplayedScrollingRowIndex = DGV購入品入力.Rows.Count - 1
        End If


        AddHandler DGV購入品入力.CellEnter, AddressOf DGV購入品入力_CellEnter
        AddHandler DGV購入品入力.CellValidating, AddressOf DGV購入品入力_CellValidating
        AddHandler DGV購入品入力.CellValueChanged, AddressOf DGV購入品入力_CellValueChanged
        AddHandler DGV購入品入力.RowsAdded, AddressOf DGV購入品入力_RowsAdded

    End Sub
    Function DGV購入品入力SQL() As String
        Dim strSql As String = ""
        '分岐させるの手間だから部署ごとに表示項目を変える必要はないそうなので、棚番号などの部署独自の項目もあえて全部署表示しています
        strSql = strSql & vbCrLf & "SELECT TD_po.id"
        strSql = strSql & vbCrLf & ", TD_po.購入ID"
        strSql = strSql & vbCrLf & ", TD_po.承認済"
        strSql = strSql & vbCrLf & ", TD_po.決裁"
        strSql = strSql & vbCrLf & ", TD_po.group_name"
        strSql = strSql & vbCrLf & ", TD_po.印刷日" '時刻まで表示されて列幅とるからFormatしていたが、そうすると○/△と入れたときに日付に変換されないからそのままSELECTすることにした
        'strSql = strSql & vbCrLf & ", format(TD_po.印刷日,'yyyy/MM/dd') AS 印刷日"
        strSql = strSql & vbCrLf & ", TD_po.order_date AS 発注日"
        strSql = strSql & vbCrLf & ", TD_po.arrival_date AS 入荷日"
        'CellFormattingイベントでFormatしているから、ここでFormatしても意味なくなった
        'strSql = strSql & vbCrLf & ", format(order_date, 'yyyy/MM/dd') AS 発注日"
        'strSql = strSql & vbCrLf & ", format(arrival_date, 'yyyy/MM/dd') AS 入荷日"
        strSql = strSql & vbCrLf & ",TD_po.登録番号"
        strSql = strSql & vbCrLf & ", TD_po.work_code AS ワークコード"

        strSql = strSql & vbCrLf & ", TD_po.division_id AS 部門ID"
        strSql = strSql & vbCrLf & ", TD_po.employee AS 購入者"
        strSql = strSql & vbCrLf & ", TD_po.part_name AS 品名"
        strSql = strSql & vbCrLf & ", TD_po.part_number AS 型式"
        strSql = strSql & vbCrLf & ", TD_po.改訂番号"
        strSql = strSql & vbCrLf & ", TD_po.number AS 数量"
        strSql = strSql & vbCrLf & ", TD_po.予算単価"
        strSql = strSql & vbCrLf & ", TD_po.number * TD_po.予算単価 AS 予算額" 'SQL上で計算することになった
        strSql = strSql & vbCrLf & ", TD_po.unit_price AS 税抜単価"
        strSql = strSql & vbCrLf & ", TD_po.number * TD_po.unit_price AS 金額" 'SQL上で計算することになった
        strSql = strSql & vbCrLf & ", TD_po.vendor_id　AS 仕入先ID"
        strSql = strSql & vbCrLf & ", TD_po.vendor as 仕入先"
        'strSql = strSql & vbCrLf & ", 営業所"
        strSql = strSql & vbCrLf & ", TD_po.pay_code"
        strSql = strSql & vbCrLf & ", TD_po.kamoku_id　AS　科目ID"
        strSql = strSql & vbCrLf & ", TD_po.remark AS 購入理由"
        strSql = strSql & vbCrLf & ", TD_po.経理検査"

        strSql = strSql & vbCrLf & ", TM_seigi_work_list.facility as 設備名"
        strSql = strSql & vbCrLf & ", TD_po.整理番号"
        strSql = strSql & vbCrLf & ", TD_po.tana_bango AS 棚番号"
        strSql = strSql & vbCrLf & ", TD_po.入荷ID"
        strSql = strSql & vbCrLf & ", TD_po.備考1"

        '元々はSELECTしていなかったが、隠す意味ないとのことで表示することになった。ユーザーから問い合わせがあった時にすぐ見てもらえるし、スクロール機能あれば問題ないとのことでスクロール機能つけた。
        strSql = strSql & vbCrLf & ", TD_po.insertStamp"
        strSql = strSql & vbCrLf & ", TD_po.作成者"
        strSql = strSql & vbCrLf & ", TD_po.作成PC名"
        strSql = strSql & vbCrLf & ", TD_po.作成バージョン"
        strSql = strSql & vbCrLf & ", TD_po.更新者"
        strSql = strSql & vbCrLf & ", TD_po.更新PC名"
        strSql = strSql & vbCrLf & ", TD_po.更新バージョン"
        strSql = strSql & vbCrLf & ", TD_po.updateStamp"
        strSql = strSql & vbCrLf & ", TM_kamoku.kamoku AS 科目名"
        strSql = strSql & vbCrLf & ", TM_division.division AS 部門名"
        '技術部のデータが重すぎてプログラムが落ちる問題の対策
        strSql = strSql & vbCrLf & ", TD_po.コピーフラグ" '2022/10/31 ekawai add

        strSql = strSql & vbCrLf & "FROM TD_po"
        strSql = strSql & vbCrLf & "LEFT OUTER JOIN TM_division"
        strSql = strSql & vbCrLf & "ON TD_po.division_id = TM_division.id"
        strSql = strSql & vbCrLf & "LEFT OUTER JOIN TM_kamoku"
        strSql = strSql & vbCrLf & "ON TD_po.kamoku_id = TM_kamoku.id "
        strSql = strSql & vbCrLf & "LEFT OUTER JOIN TM_seigi_work_list "
        strSql = strSql & vbCrLf & "ON TD_po.work_code = TM_seigi_work_list.work_code"
        strSql = strSql & vbCrLf & "WHERE TD_po.group_name = " & kc.SQ(Select_Form.s部署)

        If 日付検索Flg Then
            strSql = strSql & vbCrLf & 日付検索条件
        Else
            '技術などデータが多い部署は全データ表示すると重くなってしまうため日付条件なければinsertStampが過去2年分のものしか出さない
            strSql = strSql & vbCrLf & "AND insertStamp >= " & Today.AddYears(-2)
        End If

        If 検索Flg Then

            strSql = strSql & vbCrLf & 検索条件

        Else


        End If

        If chk未処理.Checked = True Then

            If 日付検索Flg = True And cmb検索項目日付.SelectedValue = "arrival_date" Then
                strSql = strSql & vbCrLf & "AND (unit_price IS NULL OR unit_price = 0)"
            Else
                '①入荷日　または　②税抜単価が空白　…納品書届いてから入力する項目
                strSql = strSql & vbCrLf & "AND(arrival_date is null OR unit_price IS NULL OR unit_price = 0)"
                'これ以外の条件のチェックは入力時にCheck_DGV購入品入力でやってるから不要
            End If

        End If


        If rbn未承認.Checked = True Then
            strSql = strSql & vbCrLf & "AND 承認済 = 0"
        End If

        If rbn承認済.Checked = True Then
            strSql = strSql & vbCrLf & "AND 承認済 = 1"
        End If

        If chk未印刷.Checked = True Then
            strSql = strSql & vbCrLf & "AND 印刷日 IS NULL"
        End If

        strSql = strSql & vbCrLf & "ORDER BY 購入ID,TD_po.id"

        Return strSql
    End Function

    Private Sub DGV購入品入力_表示()
        '自動計算プロシージャが起きないよう、CellValidatingとRowsAddが発生しないようにしたかったのでフラグで制御することになった
        Flg再表示 = True
        'DataSource変更した時のイベント停止したかったが、ここに書いても停止しなかったのでコメントアウト
        'RemoveHandler DGV購入品入力.CellValidating, AddressOf DGV購入品入力_CellValidating
        'RemoveHandler DGV購入品入力.RowsAdded, AddressOf DGV購入品入力_RowsAdded
        Dim ds購入 As DataSet
        'DataSourceでバインドしている時では初期化

        DGV購入品入力.DataSource = Nothing

        ds購入 = dCon.DataSet(DGV購入品入力SQL, "購入品一覧")
        DT購入品リスト = ds購入.Tables("購入品一覧")

        DGV購入品入力.DataSource = DT購入品リスト

        'If DGV購入品入力.Rows.Count > 0 Then
        '    'グリッドに表示される最初（一番上）の行の取得／設定
        '    DGV購入品入力.FirstDisplayedScrollingRowIndex = DGV購入品入力.Rows.Count - 1
        'End If

        DGV購入品入力.RowsDefaultCellStyle.BackColor = Color.FromArgb(221, 235, 247)
        DGV購入品入力.AlternatingRowsDefaultCellStyle.BackColor = Color.White
        最終更新時刻 = Now

        'DataSource変更した時のイベント停止したかったが、ここに書いても停止しなかったのでコメントアウト
        'AddHandler DGV購入品入力.CellValidating, AddressOf DGV購入品入力_CellValidating
        'AddHandler DGV購入品入力.RowsAdded, AddressOf DGV購入品入力_RowsAdded
        Flg再表示 = False

    End Sub
    Sub DGV購入品入力_詳細設定()

        Dim Col部門 As New DataGridViewComboBoxColumn
        Dim Col仕入先 As New DataGridViewComboBoxColumn
        Dim Col支払区分 As New DataGridViewComboBoxColumn
        Dim Col購入者 As New DataGridViewComboBoxColumn
        Dim Col科目 As New DataGridViewComboBoxColumn
        Dim Col科目候補 As New DataGridViewTextBoxColumn
        Dim Colワークコード As New DataGridViewComboBoxColumn
        Dim Col決裁 As New DataGridViewComboBoxColumn
        Dim Col設備名 As New DataGridViewComboBoxColumn
        Dim Col予算額 As New DataGridViewTextBoxColumn
        Dim Col金額 As New DataGridViewTextBoxColumn

        'Dim col整理番号 As New DataGridViewComboBoxColumn
        'Dim col棚番号 As New DataGridViewComboBoxColumn

        'Dim Colコピー As New DataGridViewButtonColumn

        Dim strSql As String
        '------------------------------------------------------------------------------
        '部門 
        '------------------------------------------------------------------------------
        'not_use_flg = 0という条件使うと旧データ表示できなくなるから全部表示して使っていないものは下にまとめることにした
        'strSql = "select * from TM_division"
        'strSql &= vbCrLf & "where not_use_flg = 0"

        strSql = "SELECT id, '×' + division AS division, keiri_sort_no, not_use_flg, 2 AS 並び順"
        strSql &= vbCrLf & "FROM TM_division"
        strSql &= vbCrLf & "WHERE (not_use_flg = 1)"
        strSql &= vbCrLf & "UNION"
        strSql &= vbCrLf & "SELECT id, division, keiri_sort_no, not_use_flg, 1 AS 並び順"
        strSql &= vbCrLf & "FROM TM_division"
        strSql &= vbCrLf & "WHERE (not_use_flg = 0)"
        strSql &= vbCrLf & "ORDER BY 並び順,id"
        DT部門 = dCon.DataSet(strSql, "DT").Tables(0)
        'DGVに現在存在しているdivision_id列と今作成したDataGridViewComboBoxColumnを入れ替える
        Col部門.DataPropertyName = DGV購入品入力.Columns("部門ID").DataPropertyName
        Col部門.ValueMember = "id"
        Col部門.DisplayMember = "division"
        Col部門.DataSource = DT部門

        DGV購入品入力.Columns.Insert(DGV購入品入力.Columns("部門ID").Index, Col部門)
        'DGV購入.Columns("division_id").Visible = False
        Col部門.Name = "部門"
        'Col部門.DataPropertyName = "部門"

        '------------------------------------------------------------------------------
        '購入者…自由入力可
        '------------------------------------------------------------------------------

        '購入者 idではなく文字で記録されている。退職者や苗字変わった人も出さないとダメだから2列にするしかない
        'strSql = "select TM_employee.id,TM_employee.employee"
        'strSql &= vbCrLf & "from TM_group INNER JOIN"
        'strSql &= vbCrLf & "TM_employee ON TM_group.id = TM_employee.group_id INNER JOIN"
        'strSql &= vbCrLf & "TM_dep ON TM_group.department_id = TM_dep.DepID"
        'strSql &= vbCrLf & "where retire = 0 "
        'strSql &= vbCrLf & "and TM_dep.DepName = '経営管理部'"

        strSql = "SELECT TM_employee.employee "
        strSql &= vbCrLf & "FROM TM_購入依頼権限 INNER JOIN"
        strSql &= vbCrLf & "TM_employee ON TM_購入依頼権限.社員番号 = TM_employee.id"
        strSql &= vbCrLf & "WHERE 部署ID = " & Select_Form.i部署ID
        strSql &= vbCrLf & "AND (更新 = 1 OR 承認 = 1 OR 経理 = 1 )"
        DT部署別購入者 = dCon.DataSet(strSql, "DT").Tables(0)
        For Each DR As DataRow In DT部署別購入者.Rows()
            Col購入者.Items.Add(DR("employee"))
        Next

        'Col購入者.DataPropertyName = DGV購入.Columns("employee").DataPropertyName
        'DGV購入.Columns.Insert(DGV購入.Columns("employee").Index, Col購入者)
        DGV購入品入力.Columns.Insert(DGV購入品入力.Columns("購入者").Index + 1, Col購入者)
        'DGV購入.Columns("employee").Visible = False
        Col購入者.Name = "購入者候補"
        Col購入者.HeaderText = ""
        Col購入者.DropDownWidth = 200

        '------------------------------------------------------------------------------
        '仕入先…自由入力可
        '------------------------------------------------------------------------------
        '2021/04/12 ekawai del S↓------------------
        'strSql = "SELECT facility_id, kiban, tel, fax, person, pay_code"
        'strSql &= vbCrLf & "FROM TM_facility left join TM_paycode on TM_facility.pay_method = TM_paycode.id"
        'Select Case Select_Form.i部署ID
        '    Case 24 '人事総務 
        '        strSql &= vbCrLf & "where use_kansetsu = 1"
        '    Case 25 '経営管理
        '        strSql &= vbCrLf & "where use_kansetsu = 1 or use_densan = 1"
        '    Case 30, 31 '営業 グループ統括
        '        strSql &= vbCrLf & "where use_eigyo = 1"
        '    Case 40 '第一製造
        '        strSql &= vbCrLf & "where use_dai1 = 1"
        '    Case 50 '第二製造
        '        strSql &= vbCrLf & "where use_dai2 = 1"
        '    Case 60 '品質保証部
        '        strSql &= vbCrLf & "where use_hinshitsu = 1"
        '    Case 70 '工務部
        '        strSql &= vbCrLf & "where use_komu = 1"
        '    Case 80 '技術部
        '        strSql &= vbCrLf & "where use_gizyutsu = 1"
        '    Case 81 '生技部
        '        strSql &= vbCrLf & "where use_seigi = 1"
        'End Select
        'strSql &= vbCrLf & "order by kiban"
        '2021/04/12 ekawai del E↑------------------
        '2021/04/12 ekawai add S
        strSql = "SELECT"
        strSql = strSql & vbCrLf & "TD_po.vendor_id"
        strSql = strSql & vbCrLf & "  ,'×' + TM_facility.kiban AS kiban"
        strSql = strSql & vbCrLf & "  ,TM_paycode.pay_code"
        strSql = strSql & vbCrLf & "  , 2 AS 並び順 "
        strSql = strSql & vbCrLf & "FROM TD_po "
        strSql = strSql & vbCrLf & "  INNER JOIN TM_facility "
        strSql = strSql & vbCrLf & "    ON TD_po.vendor_id = TM_facility.facility_id "
        strSql = strSql & vbCrLf & "  INNER JOIN TM_paycode ON TM_facility.pay_method = TM_paycode.id"
        strSql = strSql & vbCrLf & "  LEFT OUTER JOIN ( "
        strSql = strSql & vbCrLf & "    SELECT * "
        strSql = strSql & vbCrLf & "    FROM TM_dep_facility "
        strSql = strSql & vbCrLf & "    WHERE"
        strSql = strSql & vbCrLf & "    TM_dep_facility.部署ID = " & Select_Form.i部署ID
        strSql = strSql & vbCrLf & "  ) AS Sub1 "
        strSql = strSql & vbCrLf & "    ON TD_po.vendor_id = Sub1.仕入先ID "
        strSql = strSql & vbCrLf & "WHERE"
        strSql = strSql & vbCrLf & "  TD_po.group_name = " & kc.SQ(Select_Form.s部署)
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
        '2021/04/12 ekawai add E
        DT部署別仕入先 = dCon.DataSet(strSql, "DT").Tables(0)

        For Each DR As DataRow In DT部署別仕入先.Rows
            Col仕入先.Items.Add(DR("kiban"))
        Next
        'Col仕入先.DataPropertyName = DGV購入.Columns("vendor").DataPropertyName
        'DGV購入.Columns.Insert(DGV購入.Columns("vendor").Index - 1, Col仕入先)
        DGV購入品入力.Columns.Insert(DGV購入品入力.Columns("仕入先").Index + 1, Col仕入先)
        'DGV購入.Columns("vendor").Visible = False
        Col仕入先.Name = "仕入先候補"
        Col仕入先.HeaderText = ""
        Col仕入先.DropDownWidth = 200
        '------------------------------------------------------------------------------
        '支払区分
        '------------------------------------------------------------------------------
        '既存の購入依頼は、Excelに書いてある文字(現金、掛け、口座)を表示していただけだったので
        'TM_paycode(0:現金,1:掛け,2:その他)は使わず直接選択肢をAddすることにした。
        'でも仕入先から支払区分選べるようにしたら結局TM_paycode使わないといけないから
        'TM_paycodeのその他⇒口座に変えた
        strSql = "select * from TM_paycode"
        Dim DT支払区分 As DataTable = dCon.DataSet(strSql, "DT").Tables(0)
        '見積や製品フォームのようにdatasourceで指定したかったけど上手く行かなかったので
        'Main_Formはあえてこの方法で支払区分入れている
        For Each DR As DataRow In DT支払区分.Rows()
            Col支払区分.Items.Add(DR("pay_code"))
        Next


        Col支払区分.DataPropertyName = DGV購入品入力.Columns("pay_code").DataPropertyName
        DGV購入品入力.Columns.Insert(DGV購入品入力.Columns("pay_code").Index, Col支払区分)
        DGV購入品入力.Columns("pay_code").Visible = False
        Col支払区分.Name = "支払区分"
        'Col支払区分.DataPropertyName = "支払区分"

        '------------------------------------------------------------------------------
        '科目
        '------------------------------------------------------------------------------

        'strSql = "select * from TM_kamoku"

        'Select Case Select_Form.i部署ID
        '    Case 40 '第一製造
        '        strSql &= vbCrLf & "where use_dai1 = 1"
        '    Case 50 '第二製造
        '        strSql &= vbCrLf & "where use_dai2 = 1"
        '    Case 60 '品質保証部
        '        strSql &= vbCrLf & "where use_hinshitsu = 1"
        '    Case 81 '生技部
        '        strSql &= vbCrLf & "where use_seigi = 1"
        '    Case 24 '人事総務部
        '        strSql &= vbCrLf & "where use_kansetsu = 1"
        '    Case 25 '経営管理部
        '        strSql &= vbCrLf & "where use_kansetsu = 1 or use_densan = 1"
        '    Case 30, 31 '営業部 グループ統括
        '        strSql &= vbCrLf & "where use_eigyo = 1"
        '    Case 70 '工務
        '        strSql &= vbCrLf & "where use_komu = 1"
        '    Case 80  '技術
        '        strSql &= vbCrLf & "where use_gizyutsu = 1"
        'End Select

        strSql = "SELECT"
        strSql = strSql & vbCrLf & " TD_po.kamoku_id"
        strSql = strSql & vbCrLf & "    , '×' + CONVERT(nvarchar,TD_po.kamoku_id) + ' : ' + TM_kamoku.kamoku AS kamoku"
        strSql = strSql & vbCrLf & "  , 2 AS 並び順 "
        strSql = strSql & vbCrLf & "FROM TD_po "
        strSql = strSql & vbCrLf & "  INNER JOIN TM_kamoku "
        strSql = strSql & vbCrLf & "    ON TD_po.kamoku_id = TM_kamoku.id "
        strSql = strSql & vbCrLf & "  LEFT OUTER JOIN ( "
        strSql = strSql & vbCrLf & "    SELECT * "
        strSql = strSql & vbCrLf & "    FROM TM_dep_kamoku "
        strSql = strSql & vbCrLf & "    WHERE"
        strSql = strSql & vbCrLf & "      TM_dep_kamoku.部署ID = " & Select_Form.i部署ID
        strSql = strSql & vbCrLf & "  ) AS Sub1 "
        strSql = strSql & vbCrLf & "    ON TD_po.kamoku_id = Sub1.科目ID "
        strSql = strSql & vbCrLf & "WHERE"
        strSql = strSql & vbCrLf & "  TD_po.group_name = " & kc.SQ(Select_Form.s部署)
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
        strSql = strSql & vbCrLf & "  , kamoku_id"


        DT部署別科目 = dCon.DataSet(strSql, "DT").Tables(0)
        'DGVに現在存在しているdivision_id列と今作成したDataGridViewComboBoxColumnを入れ替える
        Col科目.DataPropertyName = DGV購入品入力.Columns("科目ID").DataPropertyName
        'Col科目.ValueMember = "id
        Col科目.ValueMember = "kamoku_id"
        Col科目.DisplayMember = "kamoku"
        Col科目.DataSource = DT部署別科目

        DGV購入品入力.Columns.Insert(DGV購入品入力.Columns("科目ID").Index, Col科目)
        DGV購入品入力.Columns("科目ID").Visible = False
        Col科目.Name = "科目"
        Col科目.DropDownWidth = 150

        '製造日報のようにテキストボックスに入れた値で科目ID検索できるようにという指示があったため列を追加した
        DGV購入品入力.Columns.Insert(DGV購入品入力.Columns("科目").Index, Col科目候補)
        Col科目候補.Name = "科目番号"
        '------------------------------------------------------------------------------
        'ワークコード　
        '------------------------------------------------------------------------------

        'strSql = "select work_code"
        'strSql &= vbCrLf & ",class1 as 部門"
        'strSql &= vbCrLf & ",class2 as 区分"
        'strSql &= vbCrLf & ",department as 部署"
        'strSql &= vbCrLf & ",line as ライン"
        'strSql &= vbCrLf & ",seiri_no as 整理番号"
        'strSql &= vbCrLf & ",facility as 設備"
        'strSql &= vbCrLf & ",distribution as 按分"
        'strSql &= vbCrLf & ",old as 終了"
        'strSql &= vbCrLf & "from TM_seigi_work_list"
        'strSql &= vbCrLf & "where old = 0 order by work_code"

        '
        strSql = "SELECT TD_po.work_code"
        strSql = strSql & vbCrLf & ",'×' + TD_po.work_code AS WK"
        strSql = strSql & vbCrLf & ", 2 AS 並び順 "
        strSql = strSql & vbCrLf & "FROM"
        strSql = strSql & vbCrLf & "  TD_po "
        strSql = strSql & vbCrLf & "  LEFT OUTER JOIN ( "
        strSql = strSql & vbCrLf & "    SELECT *"
        strSql = strSql & vbCrLf & "    FROM"
        strSql = strSql & vbCrLf & "      TM_seigi_work_list "
        strSql = strSql & vbCrLf & "    WHERE"
        strSql = strSql & vbCrLf & "      TM_seigi_work_list.old = 0"
        strSql = strSql & vbCrLf & "  ) AS T1 "
        strSql = strSql & vbCrLf & "    ON TD_po.work_code = T1.work_code "
        strSql = strSql & vbCrLf & "WHERE"
        strSql = strSql & vbCrLf & "  TD_po.group_name = '生産技術部' " '2024/05/14 ekawai 組織変更に伴い部署名変更
        strSql = strSql & vbCrLf & "  AND TD_po.work_code <> N'' "
        strSql = strSql & vbCrLf & "  AND TD_po.work_code IS NOT NULL "
        strSql = strSql & vbCrLf & "  AND T1.work_code IS NULL "
        strSql = strSql & vbCrLf & "GROUP BY"
        strSql = strSql & vbCrLf & "TD_po.work_code,'×' + TD_po.work_code"
        strSql = strSql & vbCrLf & "UNION"
        strSql = strSql & vbCrLf & "SELECT TM_seigi_work_list.work_code,'' + TM_seigi_work_list.work_code as WK, 1 AS 並び順"
        strSql = strSql & vbCrLf & "FROM TM_seigi_work_list"
        strSql = strSql & vbCrLf & "WHERE old = 0"
        strSql = strSql & vbCrLf & "ORDER BY 並び順"

        DTワークコード = dCon.DataSet(strSql, "DT").Tables(0)


        'DGVに現在存在しているwork_code列と今作成したDataGridViewComboBoxColumnを入れ替える
        Colワークコード.DataPropertyName = DGV購入品入力.Columns("ワークコード").DataPropertyName
        Colワークコード.ValueMember = "work_code"
        Colワークコード.DisplayMember = "WK" '×がついているものが表示値
        Colワークコード.DataSource = DTワークコード
        'For Each DR_wk As DataRow In DTワークコード.Rows
        '    Colワークコード.Items.Add(DR_wk("work_code"))
        'Next

        DGV購入品入力.Columns.Insert(DGV購入品入力.Columns("ワークコード").Index, Colワークコード)
        DGV購入品入力.Columns("ワークコード").Visible = False
        Colワークコード.Name = "ワークコード"
        'Colワークコード.DataPropertyName = "ワークコード"



        '------------------------------------------------------------------------------
        '整理番号　…技術に確認したらコンボは使っていないとのこと。何千行もあるからコンボだと逆に選ぶの大変なのでは？また、コンボだけど整理番号以外を入力することもあるそう。例：5,6
        '------------------------------------------------------------------------------
        'strSql = "SELECT 整理番号 FROM M_製品"
        ''strSql = "select die_shelf from TM_die_shelf"
        'Dim DT_seiriNo As DataTable = dCon.DataSet(strSql, "DT").Tables(0)


        ''DGVに現在存在しているwork_code列と今作成したDataGridViewComboBoxColumnを入れ替える
        'col整理番号.DataPropertyName = DGV購入品入力.Columns("整理番号").DataPropertyName
        'col整理番号.DataSource = DT_seiriNo
        'col整理番号.ValueMember = "整理番号"
        'col整理番号.DisplayMember = "整理番号"
        'DGV購入品入力.Columns.Insert(DGV購入品入力.Columns("整理番号").Index, col整理番号)
        'DGV購入品入力.Columns("整理番号").Visible = False
        'col整理番号.Name = "整理番号"
        ''Col整理番号.DataPropertyName = "整理番号"

        '------------------------------------------------------------------------------
        '棚番号　
        '旧購入依頼ではwork_listというシートにTM_die_shelfの棚番号読み込んでコンボで選択させていたようだが
        '名前の範囲壊れていて、コンボの選択肢一切表示されていない状態のためマスタ参照せず手入力している。
        'なので新購入依頼でも自由入力とした。
        '鈴木MはTM_die_shelfの存在すら知らなかったので、誰も管理していない模様。
        '------------------------------------------------------------------------------
        'strSql = "select die_shelf from TM_die_shelf"
        'Dim DT_TanaNo As DataTable = dCon.DataSet(strSql, "DT").Tables(0)


        ''DGVに現在存在している棚番号列と今作成したDataGridViewComboBoxColumnを入れ替える
        'col棚番号.DataPropertyName = DGV購入品入力.Columns("棚番号").DataPropertyName
        'col棚番号.DataSource = DT_TanaNo
        'col棚番号.ValueMember = "die_shelf"
        'col棚番号.DisplayMember = "die_shelf"
        'DGV購入品入力.Columns.Insert(DGV購入品入力.Columns("棚番号").Index, col棚番号)
        'DGV購入品入力.Columns("棚番号").Visible = False
        'col棚番号.Name = "棚番号"
        ''col棚番号.DataPropertyName = "整理番号"


        '------------------------------------------------------------------------------
        '決裁状況　
        '------------------------------------------------------------------------------
        Col決裁.Items.Add(Select_Form._氏名)
        Col決裁.Items.Add("否認")
        Col決裁.Items.Add("保留")
        Col決裁.Items.Add("稟議承認")

        DGV購入品入力.Columns.Insert(DGV購入品入力.Columns("決裁").Index + 1, Col決裁)
        Col決裁.Name = "決裁候補"
        'Col決裁.HeaderText = ""
        Col決裁.DropDownWidth = 100


        ''------------------------------------------------------------------------------
        ''コピー
        ''------------------------------------------------------------------------------
        'Colコピー.UseColumnTextForButtonValue = True
        'Colコピー.Text = "コピー"
        'DGV購入品入力.Columns.Add(Colコピー)
        'Colコピー.Name = "コピー納入"

        '------------------------------------------------------------------------------
        '予算額　(予算単価　×　数量)　…DGV上で計算すると時間が掛かるため、SQL上で列を作り出すよう指示があったためコメントアウト
        '------------------------------------------------------------------------------
        '予算単価の右隣に表示
        'DGV購入品入力.Columns.Insert(DGV購入品入力.Columns("予算単価").Index + 1, Col予算額)
        'Col予算額.Name = "予算額"

        '------------------------------------------------------------------------------
        '金額　(税抜単価　×　数量)　…DGV上で計算すると時間が掛かるため、SQL上で列を作り出すよう指示があったためコメントアウト
        '------------------------------------------------------------------------------
        '税抜単価の右隣に表示
        'DGV購入品入力.Columns.Insert(DGV購入品入力.Columns("税抜単価").Index + 1, Col金額)
        'Col金額.Name = "金額"

        DirectCast(DGV購入品入力.Columns("購入者"), DataGridViewTextBoxColumn).MaxInputLength = 40
        DirectCast(DGV購入品入力.Columns("品名"), DataGridViewTextBoxColumn).MaxInputLength = 125
        DirectCast(DGV購入品入力.Columns("型式"), DataGridViewTextBoxColumn).MaxInputLength = 125
        DirectCast(DGV購入品入力.Columns("仕入先"), DataGridViewTextBoxColumn).MaxInputLength = 50
        DirectCast(DGV購入品入力.Columns("棚番号"), DataGridViewTextBoxColumn).MaxInputLength = 50
        DirectCast(DGV購入品入力.Columns("整理番号"), DataGridViewTextBoxColumn).MaxInputLength = 50
        DirectCast(DGV購入品入力.Columns("改訂番号"), DataGridViewTextBoxColumn).MaxInputLength = 10
        DirectCast(DGV購入品入力.Columns("購入理由"), DataGridViewTextBoxColumn).MaxInputLength = 120
        DirectCast(DGV購入品入力.Columns("備考1"), DataGridViewTextBoxColumn).MaxInputLength = 50
    End Sub
    Private Sub DGV購入品入力_表示条件()
        DGV購入品入力.Columns("id").Visible = False 'UID非表示
        DGV購入品入力.Columns("group_name").Visible = False
        'DGV購入品入力.Columns("仕入先ID").Visible = False '2022/02/10 経理で使っているので表示することにした
        DGV購入品入力.Columns("部門ID").Visible = False
        DGV購入品入力.Columns("承認済").Visible = False
        'DGV購入品入力.Columns("updateStamp").Visible = False '2022/02/22 せかっく記録しているのに隠す意味ないとのことで表示することになった
        '部署ごとに表示を分岐すると複雑になるので全て表示していいとのことだったが、ワークコード変えると部門も変わってしまうため、ワークコードは生技の時しか表示しないようにしている。
        '技術専用の項目もいくつかあるがそれは全部署表示していいとのこと(部署ごとに分岐するの手間だから)
        If Select_Form.i部署ID <> 81 Then
            DGV購入品入力.Columns("ワークコード").Visible = False
            DGV購入品入力.Columns("設備名").Visible = False
        End If
        '検索用の部門と科目名は非表示としたい
        DGV購入品入力.Columns("科目名").Visible = False
        DGV購入品入力.Columns("部門名").Visible = False

        ''UID,発行NOでウィンドウ枠固定する
        'DGV購入.Columns("発行NO").Frozen = True
        '2022/11/2 ekawai del Start ReadOnlyは使わず、全てCellBeginEditで対応するように指示があったためコメントアウト
        'DGV購入品入力.Columns("購入ID").ReadOnly = True
        'DGV購入品入力.Columns("決裁").ReadOnly = True
        'DGV購入品入力.Columns("決裁候補").ReadOnly = True
        'DGV購入品入力.Columns("予算額").ReadOnly = True
        'DGV購入品入力.Columns("金額").ReadOnly = True
        'DGV購入品入力.Columns("入荷ID").ReadOnly = True
        'DGV購入品入力.Columns("経理検査").ReadOnly = True
        'DGV購入品入力.Columns("設備名").ReadOnly = True
        'DGV購入品入力.Columns("仕入先ID").ReadOnly = True
        'DGV購入品入力.Columns("insertStamp").ReadOnly = True
        'DGV購入品入力.Columns("作成者").ReadOnly = True
        'DGV購入品入力.Columns("作成PC名").ReadOnly = True
        'DGV購入品入力.Columns("作成バージョン").ReadOnly = True
        'DGV購入品入力.Columns("updateStamp").ReadOnly = True
        'DGV購入品入力.Columns("更新者").ReadOnly = True
        'DGV購入品入力.Columns("更新PC名").ReadOnly = True
        'DGV購入品入力.Columns("更新バージョン").ReadOnly = True
        '2022/03/04 ekawai add S 生技部だったら部門はロック…seigi_work_listマスタの部門を正としたいから
        'If Select_Form.s部署 = "生技部" Then
        '    DGV購入品入力.Columns("部門").ReadOnly = True
        'End If
        '2022/03/04 ekawai add E
        '2022/11/02 ekawai del E

        '表示している列は全てサイズ指定したほうが良いそう
        DGV購入品入力.Columns("購入ID").Width = 55
        'DGV購入品入力.Columns("承認済").Width = 30
        DGV購入品入力.Columns("決裁候補").Width = 20
        DGV購入品入力.Columns("決裁").Width = 60
        DGV購入品入力.Columns("登録番号").Width = 30
        DGV購入品入力.Columns("発注日").Width = 70
        DGV購入品入力.Columns("入荷日").Width = 70
        DGV購入品入力.Columns("印刷日").Width = 80
        DGV購入品入力.Columns("部門").Width = 100
        DGV購入品入力.Columns("購入者候補").Width = 20
        DGV購入品入力.Columns("購入者").Width = 60
        DGV購入品入力.Columns("品名").Width = 120
        DGV購入品入力.Columns("型式").Width = 120
        DGV購入品入力.Columns("仕入先ID").Width = 40
        DGV購入品入力.Columns("仕入先候補").Width = 20
        DGV購入品入力.Columns("仕入先").Width = 80
        DGV購入品入力.Columns("改訂番号").Width = 30
        'DGV購入品入力.Columns("営業所").Width = 40　生技部の過去約1年の購入依頼全て見たが実績なし。今は誰も使っていないと判断して削除した。
        DGV購入品入力.Columns("支払区分").Width = 50
        DGV購入品入力.Columns("数量").Width = 40
        DGV購入品入力.Columns("予算単価").Width = 60
        DGV購入品入力.Columns("予算額").Width = 60
        DGV購入品入力.Columns("税抜単価").Width = 60
        DGV購入品入力.Columns("金額").Width = 60
        DGV購入品入力.Columns("科目番号").Width = 40
        DGV購入品入力.Columns("科目").Width = 140
        DGV購入品入力.Columns("購入理由").Width = 130
        DGV購入品入力.Columns("ワークコード").Width = 80
        DGV購入品入力.Columns("整理番号").Width = 60
        DGV購入品入力.Columns("棚番号").Width = 80
        DGV購入品入力.Columns("入荷ID").Width = 50
        DGV購入品入力.Columns("経理検査").Width = 30
        DGV購入品入力.Columns("備考1").Width = 80
        DGV購入品入力.Columns("設備名").Width = 80
        '折り返さないで表示できる列幅にすればよいとのこと
        DGV購入品入力.Columns("insertStamp").Width = 100
        DGV購入品入力.Columns("作成者").Width = 80
        DGV購入品入力.Columns("作成PC名").Width = 70
        DGV購入品入力.Columns("作成バージョン").Width = 60
        DGV購入品入力.Columns("updateStamp").Width = 100
        DGV購入品入力.Columns("更新者").Width = 80
        DGV購入品入力.Columns("更新PC名").Width = 70
        DGV購入品入力.Columns("更新バージョン").Width = 60



        '承認レイアウト　列幅0にすると微妙に枠が残るので、出来る限りVisible=Falseにしている。
        If Select_Form._MODE = "決裁" Then
            DGV購入品入力.Columns("登録番号").Visible = False
            DGV購入品入力.Columns("発注日").Visible = False
            DGV購入品入力.Columns("入荷日").Visible = False
            DGV購入品入力.Columns("印刷日").Width = 0 '承認後印刷する人もいるから表示できるようにしたほうがよいそうなので、あえてVisible=Falseにしていない
            DGV購入品入力.Columns("改訂番号").Visible = False
            DGV購入品入力.Columns("仕入先ID").Visible = False
            'DGV購入品入力.Columns("科目番号").Visible = False 'DGV購入品入力_CurrentCellDirtyStateChangedでエラーが出るからVisible=Falseにできない
            DGV購入品入力.Columns("科目番号").Width = 0
            DGV購入品入力.Columns("税抜単価").Visible = False
            DGV購入品入力.Columns("金額").Visible = False
        End If

        'セルのテキストを折り返して表示する
        DGV購入品入力.Columns("決裁").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        DGV購入品入力.Columns("購入者").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        DGV購入品入力.Columns("品名").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        DGV購入品入力.Columns("型式").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        DGV購入品入力.Columns("仕入先").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        DGV購入品入力.Columns("購入理由").DefaultCellStyle.WrapMode = DataGridViewTriState.True
        DGV購入品入力.Columns("備考1").DefaultCellStyle.WrapMode = DataGridViewTriState.True



    End Sub

    Private Sub DGV購入品入力_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DGV購入品入力.CellBeginEdit
        'Dim dt As DataTable = DGV購入品入力.DataSource
        'For Each ROW As DataRow In dt.Rows
        '    Debug.Print("cellbeginedit")
        '    Debug.Print(ROW("品名"))
        '    Debug.Print(ROW("品名", DataRowVersion.Original))
        '    Debug.Print(ROW("品名", DataRowVersion.Current))
        '    Debug.Print(ROW.RowState)
        'Next

        'CellEditStart = True　’使ってなさそうだからコメントアウト
        '2022/11/2 ekawai add


        If Select_Form.経理Flg = False And Select_Form.更新Flg = False And Select_Form.承認Flg = False Then
            '閲覧権限しかない場合は、キャンセルする
            e.Cancel = True
        Else
            If kc.nz(DGV購入品入力("経理検査", e.RowIndex).Value) = False Then
                '経理検査済みでなければ条件によって細かく分岐が必要
                Select Case DGV購入品入力.Columns(e.ColumnIndex).HeaderText
                    Case "購入ID", "決裁", "予算額", "金額", "入荷ID", "設備名", "仕入先ID", "insertStamp", "作成者", "作成PC名", "作成バージョン", "updateStamp", "更新者", "更新PC名", "更新バージョン", "コピーフラグ"
                        e.Cancel = True
                    Case "品名", "登録番号", "予算単価"
                        If kc.nz(DGV購入品入力("承認済", e.RowIndex).Value) = False Then
                            '承認済みでなければ編集可能
                        Else
                            If Select_Form.経理Flg = True Then
                                '経理だったら承認済みでも編集可能とする
                            Else
                                '経理以外で承認済みだったらキャンセル
                                e.Cancel = True
                            End If

                        End If
                    Case "数量", "型式"
                        If kc.nz(DGV購入品入力("承認済", e.RowIndex).Value) = False Then
                            '承認済みでなければ編集可能
                        Else
                            '承認済みだったら
                            If kc.nz(DGV購入品入力("コピーフラグ", e.RowIndex).Value) = True Or Select_Form.経理Flg = True Then
                                'コピーフラグがTrueまたは経理だったら編集可能
                            Else
                                'コピーフラグがTrueでなければキャンセル
                                e.Cancel = True
                            End If
                        End If
                    Case "決裁候補"
                        If Select_Form.承認Flg = True Then
                            '承認者だったら選択可能とする
                        Else
                            '承認者以外ならキャンセル
                            e.Cancel = True
                        End If
                    Case "経理検査"
                        If Select_Form.経理Flg = True Then
                            '経理なら選択可能とする
                        Else
                            '経理以外ならキャンセル
                            e.Cancel = True
                        End If
                    Case "部門"
                        If Select_Form.s部署 <> "生産技術部" Then '2024/05/14 ekawai 組織変更に伴い部署名変更
                            '生技部以外なら選択可能とする
                        Else
                            '生技部だったらキャンセル
                            e.Cancel = True
                        End If
     

                End Select

            Else
                '経理検査済みだったら
                Select Case DGV購入品入力.Columns(e.ColumnIndex).HeaderText
                    Case "経理検査"
                        If Select_Form.経理Flg = True Then
                            '経理検査は経理のみ編集可能
                        Else
                            e.Cancel = True

                        End If
                    Case "決裁候補"
                        If Select_Form.承認Flg = True Then
                            '承認者だったら決裁候補のみ編集可能(Excel購入依頼の仕様を引き継いでいるだけなので、このような仕様にした理由は不明)
                        Else
                            e.Cancel = True
                        End If
                    Case Else
                        e.Cancel = True
                End Select


            End If





        End If
        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim str編集列名 As String = dgv.CurrentCell.OwningColumn.Name
        If str編集列名 = "印刷日" Or str編集列名 = "発注日" Or str編集列名 = "入荷日" Then
            c日付セル = dgv.CurrentCell 'c日付セルフラグで、日付かそれ以外かでCellValidatingでの処理を分岐させている
        Else
            c日付セル = Nothing
            '仕入先セル触ったら仕入先IDリセットする
            'If str編集列名 = "仕入先" Then
            '    dgv.Rows(e.RowIndex).Cells("仕入先ID").Value = DBNull.Value
            'End If
        End If
    End Sub
    '行ごとにコピーボタンつけるつもりだったが、他のボタンと統一したほうがいいとのことだったので使わなくなった
    'Private Sub DGV購入品入力_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV購入品入力.CellContentClick
    '    'DGVのボタンがクリックされた時の処理
    '    Dim strSql As String
    '    Dim dgv As DataGridView = CType(sender, DataGridView)
    '    Dim row As DataGridViewRow
    '    If e.RowIndex > 0 Then
    '        row = dgv.Rows(e.RowIndex)
    '        If dgv.Columns(e.ColumnIndex).Name = "コピー納入" Then
    '            'バインド元 TD_poに行を追加する

    '            Dim dt As DataTable = DGV購入品入力.DataSource
    '            Dim tran As System.Data.SqlClient.SqlTransaction = Nothing

    '            tran = dCon.Connection.BeginTransaction

    '            strSql = "INSERT INTO"
    '            strSql &= vbCrLf & "TD_po("
    '            strSql &= vbCrLf & "group_name"
    '            strSql &= vbCrLf & ",part_name"
    '            strSql &= vbCrLf & ",order_date"
    '            strSql &= vbCrLf & ",arrival_date"
    '            strSql &= vbCrLf & ",印刷日"
    '            strSql &= vbCrLf & ",division_id"
    '            strSql &= vbCrLf & ",employee"
    '            strSql &= vbCrLf & ",part_number"
    '            strSql &= vbCrLf & ",unit_price"
    '            strSql &= vbCrLf & ",number"
    '            strSql &= vbCrLf & ",kamoku_id"
    '            strSql &= vbCrLf & ",vendor_id"
    '            strSql &= vbCrLf & ",vendor"
    '            strSql &= vbCrLf & ",pay_code"
    '            strSql &= vbCrLf & ",remark"
    '            strSql &= vbCrLf & ",work_code"
    '            strSql &= vbCrLf & ",tana_bango"
    '            strSql &= vbCrLf & ",承認済"
    '            strSql &= vbCrLf & ",決裁"
    '            strSql &= vbCrLf & ",登録番号"
    '            strSql &= vbCrLf & ",予算単価"
    '            strSql &= vbCrLf & ",整理番号"
    '            strSql &= vbCrLf & ",改訂番号"
    '            'strSql &= vbCrLf & ",営業所"
    '            strSql &= vbCrLf & ",購入ID"
    '            strSql &= vbCrLf & ",insertStamp"
    '            strSql &= vbCrLf & ") VALUES ("
    '            strSql &= vbCrLf & SQ(Select_Form.s部署)
    '            strSql &= vbCrLf & "," & SQ(row.Cells("品名").Value)
    '            strSql &= vbCrLf & "," & nr(row.Cells("発注日").Value)
    '            strSql &= vbCrLf & "," & nr(row.Cells("入荷日").Value)
    '            strSql &= vbCrLf & "," & nr(row.Cells("印刷日").Value)
    '            strSql &= vbCrLf & "," & SQ(row.Cells("部門ID").Value)
    '            strSql &= vbCrLf & "," & SQ(row.Cells("購入者").Value)
    '            strSql &= vbCrLf & "," & nr(row.Cells("型式").Value)
    '            strSql &= vbCrLf & "," & nr(row.Cells("税抜単価").Value)
    '            strSql &= vbCrLf & ",0"
    '            strSql &= vbCrLf & "," & row.Cells("科目ID").Value
    '            strSql &= vbCrLf & "," & nr(row.Cells("仕入先ID").Value)
    '            strSql &= vbCrLf & "," & SQ(row.Cells("仕入先").Value)
    '            strSql &= vbCrLf & "," & SQ(row.Cells("pay_code").Value)
    '            strSql &= vbCrLf & "," & nr(row.Cells("購入理由").Value)
    '            strSql &= vbCrLf & "," & nr(row.Cells("ワークコード").Value)
    '            strSql &= vbCrLf & "," & nr(row.Cells("棚番号").Value)
    '            strSql &= vbCrLf & "," & nr(row.Cells("承認済").Value)
    '            strSql &= vbCrLf & "," & nr(row.Cells("決裁").Value)
    '            strSql &= vbCrLf & "," & nr(row.Cells("登録番号").Value)
    '            strSql &= vbCrLf & "," & nz(row.Cells("予算単価").Value)
    '            strSql &= vbCrLf & "," & nr(row.Cells("整理番号").Value)
    '            strSql &= vbCrLf & "," & nr(row.Cells("改訂番号").Value)
    '            'strSql &= vbCrLf & "," & nr(row.Cells("営業所").Value)
    '            strSql &= vbCrLf & ",'" & row.Cells("購入ID").Value
    '            strSql &= vbCrLf & "," & SQ(Now)
    '            strSql &= vbCrLf & ")"


    '            If dCon.ExecuteSqlMW(tran, strSql) = False Then
    '                tran.Rollback()
    '                MsgBox("更新失敗")
    '                Exit Sub
    '            End If
    '            tran.Commit()
    '            tran.Dispose()
    '            DGV購入品入力_Read()
    '        End If
    '    End If
    'End Sub


    Private Sub DGV購入品入力_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DGV購入品入力.CellEndEdit
        'Dim dt As DataTable = DGV購入品入力.DataSource
        'For Each ROW As DataRow In dt.Rows
        '    Debug.Print("cellendedit")
        '    Debug.Print(ROW("品名"))
        '    Debug.Print(ROW("品名", DataRowVersion.Original))
        '    Debug.Print(ROW("品名", DataRowVersion.Current))
        '    Debug.Print(ROW.RowState)
        'Next

        DGVテキストボックス変更時処理(e.RowIndex, e.ColumnIndex)
        'Dim DR As DataRow()

        ''編集したら保存メッセージを表示(カーソル入れただけでも発生するけど)
        'SwEditSave = True
        'Select Case DGV購入品入力.Columns(e.ColumnIndex).Name
        '    Case "数量", "税抜単価", "予算単価"
        '        If nz(DGV購入品入力.Rows(e.RowIndex).Cells("数量").Value) > 0 Then

        '            If nz(DGV購入品入力.Rows(e.RowIndex).Cells("税抜単価").Value) <> 0 Then
        '                自動計算(e.RowIndex, False)
        '            Else
        '                DGV購入品入力.Rows(e.RowIndex).Cells("金額").Value = ""
        '            End If


        '            If nz(DGV購入品入力.Rows(e.RowIndex).Cells("予算単価").Value) <> 0 Then
        '                自動計算(e.RowIndex, True)
        '            Else
        '                DGV購入品入力.Rows(e.RowIndex).Cells("予算額").Value = ""
        '            End If


        '        Else
        '            DGV購入品入力.Rows(e.RowIndex).Cells("予算額").Value = ""
        '            DGV購入品入力.Rows(e.RowIndex).Cells("金額").Value = ""
        '        End If
        '    Case "科目番号"
        '        If ns(DGV購入品入力(e.ColumnIndex, e.RowIndex).Value) <> "" Then
        '            Dim i科目ID As Integer
        '            '科目をリセットする
        '            DGV購入品入力("科目", e.RowIndex).Value = DBNull.Value
        '            i科目ID = nz(DGV購入品入力.Rows(e.RowIndex).Cells("科目番号").Value)
        '            '科目番号に応じた科目コンボを選択する
        '            DR = DT現科目.Select("id = " & i科目ID)
        '            If DR.Length = 1 Then
        '                DGV購入品入力("科目", e.RowIndex).Value = DR(0)("id")
        '            Else
        '                DGV購入品入力(e.ColumnIndex, e.RowIndex).Value = DBNull.Value
        '            End If
        '        End If

        '        '仕入先マスタに存在する仕入先名を入力したら、仕入先IDと支払区分が自動で入るようにする
        '    Case "仕入先"
        '        Dim str仕入先名 As String
        '        str仕入先名 = ns(DGV購入品入力("仕入先", e.RowIndex).Value)
        '        GET仕入先ID(str仕入先名, e.RowIndex)
        '    Case "登録番号"
        '        '部門と科目がすぐに反映されるようにするため
        '        DGV購入品入力.Refresh()

        'End Select

        'コンボボックスを閉じるたびに発生する。コンボボックスを閉じずに選択肢を連続的に変える場合はコンボが閉じられるまではここを通らない
        If Not (Me.dataGridViewComboBox Is Nothing) Then
            '2021/04/12 ekawai del S↓------------------
            'RemoveHandler Me.dataGridViewComboBox.SelectedIndexChanged, _
            '    AddressOf dataGridViewComboBox_SelectedIndexChanged
            'Me.dataGridViewComboBox = Nothing
            '2021/04/12 ekawai del E↑------------------
            コンボイベントハンドラ削除() '2021/04/12 ekawai add S
            EventFlg = False
        End If

    End Sub

    Sub DGVテキストボックス変更時処理(iRow As Integer, iCol As Integer)

        Dim DR As DataRow()

        '編集したら保存メッセージを表示(カーソル入れただけでも発生するけど)
        'SwEditSave = True  →　DGV購入品入力_CurrentCellDirtyStateChangedイベントに移動
        Select Case DGV購入品入力.Columns(iCol).Name
            'Case "数量", "税抜単価", "予算単価"
            '    If nz(DGV購入品入力.Rows(iRow).Cells("数量").Value) > 0 Then

            '        If nz(DGV購入品入力.Rows(iRow).Cells("税抜単価").Value) <> 0 Then
            '            自動計算(iRow, False)
            '        Else
            '            DGV購入品入力.Rows(iRow).Cells("金額").Value = ""
            '        End If

            '        If nz(DGV購入品入力.Rows(iRow).Cells("予算単価").Value) <> 0 Then
            '            自動計算(iRow, True)
            '        Else
            '            DGV購入品入力.Rows(iRow).Cells("予算額").Value = ""
            '        End If

            '    Else
            '        DGV購入品入力.Rows(iRow).Cells("予算額").Value = ""
            '        DGV購入品入力.Rows(iRow).Cells("金額").Value = ""
            '    End If
            '
            Case "科目番号"
                If kc.nz(DGV購入品入力(iCol, iRow).Value) <> 0 Then
                    Dim i科目ID As Integer
                    '科目をリセットする
                    DGV購入品入力("科目", iRow).Value = DBNull.Value
                    i科目ID = kc.nz(DGV購入品入力.Rows(iRow).Cells("科目番号").Value)
                    '科目番号に応じた科目コンボを選択する
                    DR = DT現科目.Select("id = " & i科目ID)
                    If DR.Length = 1 Then
                        DGV購入品入力("科目", iRow).Value = DR(0)("id")
                    Else
                        DGV購入品入力(iCol, iRow).Value = DBNull.Value
                    End If
                End If

                '仕入先マスタに存在する仕入先名を入力したら、仕入先IDと支払区分が自動で入る
                '製品マスタに登録している行を購入品入力にコピーした場合、製品マスタに登録されている支払区分がコピーした時点では入るが
                'あとで仕入先のセルを触ると仕入先マスタに登録されている支払区分に書き変わってしまうという指摘があったのでITで方針を話し合った
                '→今のままでいい。1つの仕入先に支払区分が複数あるはずない。経理が全ての仕入先に支払区分登録すればよいとのこと　2022.02.17
            Case "仕入先"
                Dim str仕入先名 As String
                str仕入先名 = kc.ns(DGV購入品入力("仕入先", iRow).Value)
                GET仕入先ID(str仕入先名, iRow)
            Case "登録番号"
                '部門と科目がすぐに反映されるようにするため
                DGV購入品入力.Refresh()

        End Select


    End Sub
    Private Sub dataGridViewComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim s選択値 As String
        Dim i As Integer
        Dim strSql As String
        'コンボボックス閉じずに選択肢を変えたときに反映されなくなるからコメントアウト→CASE文の中に移動…コンボイベントハンドラ削除
        'If EventFlg = True Then
        '    RemoveHandler Me.dataGridViewComboBox.SelectedIndexChanged, _
        '        AddressOf dataGridViewComboBox_SelectedIndexChanged
        '    EventFlg = False
        'End If

        'オブジェクト型からDataGridViewComboBoxEditingControlに変換
        Dim cb As DataGridViewComboBoxEditingControl = _
            CType(sender, DataGridViewComboBoxEditingControl)
        If cb.SelectedIndex = -1 Then
            Exit Sub
        End If
        s選択値 = cb.SelectedItem.ToString
        i = DGV購入品入力.CurrentCell.RowIndex

        Select Case DGV購入品入力.CurrentCell.OwningColumn.Name

            Case "仕入先候補"
                コンボイベントハンドラ削除()
                '×から始まる仕入先名を選んだら
                If s選択値.Substring(0, 1) = "×" Then
                    MsgBox("現在使われていない項目です")
                    cb.SelectedIndex = 変更前index
                Else
                    DGV購入品入力.Rows(i).Cells("仕入先").Value = s選択値
                    GET仕入先ID(s選択値, i)
                End If
                コンボイベントハンドラ追加()
            Case "購入者候補"

                コンボイベントハンドラ削除()
                DGV購入品入力.Rows(i).Cells("購入者").Value = s選択値
                'addHandlerしないとコンボを閉じずに連続で選択肢を変えた場合にテキストボックスDGV購入品入力.Rows(i).Cells("購入者").Valueが変わらなくなる不具合の対策
                '同じイベントハンドラが同じテキストボックスのイベントに何回も追加されないよう、コンボを閉じる時にDGV購入品入力_CellEndEditでremoveHandlerしている
                コンボイベントハンドラ追加()

            Case "決裁候補"
                コンボイベントハンドラ削除()
                DGV購入品入力.Rows(i).Cells("決裁").Value = s選択値
                コンボイベントハンドラ追加()
                '何故か○○候補コンボの場合は大丈夫だがワークコードだとs選択値がsystem.data.datarowviewになってしまう問題が発生したので、イベントを変更している
                'Case "ワークコード"
                '    'Excel購入依頼の仕様を引き継いで、ワークコード入れたらTM_seigi_work_listに登録されている部門名と設備名が自動で入るようにしている
                '    strSql = "SELECT TM_division.id,TM_seigi_work_list.facility"
                '    strSql &= vbCrLf & "FROM TM_seigi_work_list INNER JOIN"
                '    strSql &= vbCrLf & "TM_division ON TM_seigi_work_list.class1  = TM_division.division"
                '    strSql &= vbCrLf & "WHERE TM_seigi_work_list.work_code = " & SQ(s選択値)
                '    Dim DT生技部門 As DataTable = dCon.DataSet(strSql, "DT").Tables(0)
                '    'マスタに該当するIDがなければ空白を入れる
                '    If DT生技部門.Rows.Count = 0 Then
                '        DGV購入品入力.Rows(i).Cells("部門").Value = ""
                '    Else
                '        DGV購入品入力.Rows(i).Cells("部門").Value = DT生技部門.Rows(0)("id")
                '        DGV購入品入力.Rows(i).Cells("設備名").Value = DT生技部門.Rows(0)("facility")
                '    End If

                '    Dim sTarget As String = cb.EditingControlFormattedValue
                '    If sTarget.Substring(0, 1) = "×" Then
                '        MsgBox("現在使われていない項目です")
                '        cb.SelectedIndex = 変更前index

                'End If
            Case "科目", "部門"
                'カレントセルを取得する時こうやって指定しないとコンボの背景が黒くなるため
                Dim sTarget = cb.EditingControlFormattedValue
                If sTarget.Substring(0, 1) = "×" Then
                    MsgBox("現在使われていない項目です")
                    コンボイベントハンドラ削除()
                    cb.SelectedIndex = 変更前index
                    コンボイベントハンドラ追加()
                End If

        End Select

        'Me.dataGridViewComboBox = Nothing
    End Sub
    Sub コンボイベントハンドラ削除() 'SelectedIndexChangedが何回も発生してStackOverflowでプログラムが落ちる対策
        RemoveHandler Me.dataGridViewComboBox.SelectedIndexChanged, AddressOf dataGridViewComboBox_SelectedIndexChanged
    End Sub
    Sub コンボイベントハンドラ追加()
        AddHandler Me.dataGridViewComboBox.SelectedIndexChanged, AddressOf dataGridViewComboBox_SelectedIndexChanged
    End Sub


    Private Sub dataGridViewComboBox_dropdownclosed(ByVal sender As Object, ByVal e As EventArgs)

    End Sub

    '----------------------------------------------------------------------------------------------------------
    'DataGridViewに表示されているテキストボックスのKeyPressイベントハンドラ 日付ＯＮＬＹ
    '----------------------------------------------------------------------------------------------------------
    Private Sub dataGridViewTextBox_KeyPressIsDate(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        '数字しか入力できないようにする
        If Asc(e.KeyChar) = 8 Then Exit Sub 'バックスペース
        If e.KeyChar = "." Then Exit Sub
        If e.KeyChar = "/" Then Exit Sub
        If e.KeyChar < "0"c Or e.KeyChar > "9"c Then
            If e.KeyChar <> ""c Then
                '指定の文字以外の場合はキーイベントをキャンセル
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

    Private Sub DGV購入品入力_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DGV購入品入力.CurrentCellDirtyStateChanged, DGV購入品入力.Sorted

        '元々はCellEndEditでSwEditSave=Trueにしていたが、カーソルを入れて隣のセルにタブで移動しただけ(unchangedの行)でもフラグが立ってしまう問題があった
        'このイベントは、何かしら値を入れないと発生しないのでSwEditSave=Trueにするタイミングとしてよりふさわしいと考え、ここでフラグを立てることにした
        SwEditSave = True

        'Dim dt As DataTable = DGV購入品入力.DataSource
        'For Each ROW As DataRow In dt.Rows
        '    Debug.Print("currentcelldirtystatechanged")
        '    Debug.Print(ROW("品名"))
        '    Debug.Print(ROW("品名", DataRowVersion.Original))
        '    Debug.Print(ROW("品名", DataRowVersion.Current))
        '    Debug.Print(ROW.RowState)
        'Next

        Dim dgv As DataGridView = CType(sender, DataGridView)
        '科目と部門と支払い区分の時は、コンボ選択したら自動で隣のセルに移動するようにする
        If dgv.CurrentCell.OwningColumn.Name = "科目" Then
            If dgv.IsCurrentCellDirty Then
                dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)

            End If

            dgv.CurrentCell = dgv("購入理由", dgv.CurrentRow.Index)
        End If

        If dgv.CurrentCell.OwningColumn.Name = "部門" Then
            If dgv.IsCurrentCellDirty Then
                dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)

            End If

            dgv.CurrentCell = dgv("購入者", dgv.CurrentRow.Index)
        End If

        If dgv.CurrentCell.OwningColumn.Name = "支払区分" Then
            If dgv.IsCurrentCellDirty Then
                dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)

            End If

            dgv.CurrentCell = dgv("科目番号", dgv.CurrentRow.Index)
        End If

        If dgv.CurrentCell.OwningColumn.Name = "ワークコード" Then
            If dgv.IsCurrentCellDirty Then
                dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)

            End If

            dgv.CurrentCell = dgv("購入者", dgv.CurrentRow.Index)
        End If

    End Sub

    'ユーザーが指定した DataGridView.CellValueChanged 値がコミットされたときに発生します。これは通常、フォーカスがセルから離れるときに発生します。
    Private Sub DGV購入品入力_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DGV購入品入力.CellValueChanged
        Dim str入力値 As String
        If DGV購入品入力.Rows(e.RowIndex).Cells(e.ColumnIndex).Value IsNot DBNull.Value Then
            str入力値 = DGV購入品入力.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
        Else
            str入力値 = ""
        End If

        'Debug.Print(str入力値)
        'Debug.Print(DGV購入品入力.CurrentCell.Value)
        'Debug.Print(DGV購入品入力.CurrentCell.EditedFormattedValue)
        Dim str編集列名 As String = DGV購入品入力.Columns(e.ColumnIndex).Name
        'ヘッダーが変更された時はe.Rowindexが-1になるからその時は除外しないとインデックスが有効範囲にないエラーが出る
        If e.RowIndex >= 0 Then
            Select Case str編集列名
                Case "ワークコード"
                    'Excel購入依頼の仕様を引き継いで、ワークコード入れたらTM_seigi_work_listに登録されている部門名と設備名が自動で入るようにしている
                    Dim strSql As String
                    strSql = "SELECT TM_division.id,TM_seigi_work_list.facility"
                    strSql &= vbCrLf & "FROM TM_seigi_work_list INNER JOIN"
                    strSql &= vbCrLf & "TM_division ON TM_seigi_work_list.class1  = TM_division.division"
                    strSql &= vbCrLf & "WHERE TM_seigi_work_list.work_code = " & kc.SQ(DGV購入品入力.Rows(e.RowIndex).Cells("ワークコード").Value)
                    Dim DT生技部門 As DataTable = dCon.DataSet(strSql, "DT").Tables(0)
                    'マスタに該当するIDがなければ空白を入れる
                    If DT生技部門.Rows.Count = 0 Then
                        RemoveHandler DGV購入品入力.CellValueChanged, AddressOf DGV購入品入力_CellValueChanged
                        DGV購入品入力.Rows(e.RowIndex).Cells("部門").Value = ""
                        DGV購入品入力.Rows(e.RowIndex).Cells("設備名").Value = ""
                        AddHandler DGV購入品入力.CellValueChanged, AddressOf DGV購入品入力_CellValueChanged

                    Else
                        RemoveHandler DGV購入品入力.CellValueChanged, AddressOf DGV購入品入力_CellValueChanged
                        DGV購入品入力.Rows(e.RowIndex).Cells("部門").Value = DT生技部門.Rows(0)("id")
                        DGV購入品入力.Rows(e.RowIndex).Cells("設備名").Value = DT生技部門.Rows(0)("facility")
                        AddHandler DGV購入品入力.CellValueChanged, AddressOf DGV購入品入力_CellValueChanged

                    End If

                    Dim rows As DataRow()
                    'str入力値にはコンボのValueMemberが入っているのでDispllayMemberをDTワークコードから探して表示する
                    rows = DTワークコード.Select("work_code = " & kc.SQ(str入力値))
                    If rows(0)("WK").Substring(0, 1) = "×" Then
                        MsgBox("現在使われていない項目です")
                        RemoveHandler DGV購入品入力.CellValueChanged, AddressOf DGV購入品入力_CellValueChanged
                        DGV購入品入力.Rows(e.RowIndex).Cells("ワークコード").Value = 変更前Value
                        AddHandler DGV購入品入力.CellValueChanged, AddressOf DGV購入品入力_CellValueChanged
                    End If

                Case "印刷日", "発注日", "入荷日"
                    If IsDate(str入力値) = False And str入力値 <> "" Then
                        MsgBox("入力日付エラー")
                        Exit Sub

                    End If
                Case "登録番号" '承認済でない、かつ経理検査済でない
                    If kc.nz(DGV購入品入力.Rows(e.RowIndex).Cells("承認済").Value) <> 1 And kc.nz(DGV購入品入力.Rows(e.RowIndex).Cells("経理検査").Value) <> 1 Then
                        'Get_製品登録情報(e.RowIndex, DGV購入品入力.CurrentCell.EditedFormattedValue.ToString) '貼り付けボタン使用時に貼り付け箇所の直前のセルの値が取得されてしまっていたためコメントアウト
                        Get_製品登録情報(e.RowIndex, DGV購入品入力.Rows(e.RowIndex).Cells("登録番号").EditedFormattedValue)
                    End If
            End Select


            '2022/01/24 ekawai del S
            'CellValueChanged使うように変更した結果、製品からコピーした時に金額が自動計算されなくなってしまったため処理を見直すことになった
            'If nz(DGV購入品入力.Rows(e.RowIndex).Cells("数量").Value) > 0 Then
            '    Select Case str編集列名
            '        Case "税抜単価"
            '            If nz(DGV購入品入力.Rows(e.RowIndex).Cells("税抜単価").Value) <> 0 Then
            '                自動計算(e.RowIndex, False)
            '            Else
            '                DGV購入品入力.Rows(e.RowIndex).Cells("金額").Value = DBNull.Value
            '            End If

            '        Case "予算単価"
            '            If nz(DGV購入品入力.Rows(e.RowIndex).Cells("予算単価").Value) <> 0 Then
            '                自動計算(e.RowIndex, True)
            '            Else
            '                DGV購入品入力.Rows(e.RowIndex).Cells("予算額").Value = DBNull.Value
            '            End If
            '        Case "数量"
            '            If nz(DGV購入品入力.Rows(e.RowIndex).Cells("税抜単価").Value) <> 0 Then
            '                自動計算(e.RowIndex, False)
            '            Else
            '                DGV購入品入力.Rows(e.RowIndex).Cells("金額").Value = DBNull.Value
            '            End If

            '            If nz(DGV購入品入力.Rows(e.RowIndex).Cells("予算単価").Value) <> 0 Then
            '                自動計算(e.RowIndex, True)
            '            Else
            '                DGV購入品入力.Rows(e.RowIndex).Cells("予算額").Value = DBNull.Value
            '            End If

            '    End Select
            'Else
            '    '数量が0だと更新できないので警告する
            '    DGV購入品入力.Rows(e.RowIndex).Cells("金額").Value = DBNull.Value
            '    DGV購入品入力.Rows(e.RowIndex).Cells("予算額").Value = DBNull.Value

            '2022/01/24 ekawai del E
            Select Case str編集列名
                Case "予算単価"
                    自動計算(e.RowIndex, True)
                Case "税抜単価"
                    自動計算(e.RowIndex, False)
                Case "数量"
                    自動計算(e.RowIndex, True)
                    自動計算(e.RowIndex, False)
            End Select
        End If
    End Sub

    'Private Sub DGV購入品入力_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DGV購入品入力.CellFormatting
    '    Dim str編集列名 As String = DGV購入品入力.Columns(e.ColumnIndex).Name
    'CellValueChangedへ移動。
    '計算方法が変わってしまうことに不安感があったので毎回CelFormattingで計算していたが、
    '時間がかかるので過去データはSQLで、DGV上で変更したものは自動計算プロシージャで計算すればよいそう。
    'SQLと.NETで計算結果変わることよっぽどないそう。違ってもデータベースには入れていない仮の行だからあまり問題にならないらしい。
    'If nz(DGV購入品入力.Rows(e.RowIndex).Cells("数量").Value) > 0 Then
    '    Select Case str編集列名
    '        Case "金額"
    '            If nz(DGV購入品入力.Rows(e.RowIndex).Cells("税抜単価").Value) <> 0 Then
    '                自動計算(e.RowIndex, False)
    '            Else
    '                DGV購入品入力.Rows(e.RowIndex).Cells("金額").Value = ""
    '            End If

    '        Case "予算額"
    '            If nz(DGV購入品入力.Rows(e.RowIndex).Cells("予算単価").Value) <> 0 Then
    '                自動計算(e.RowIndex, True)
    '            Else
    '                DGV購入品入力.Rows(e.RowIndex).Cells("予算額").Value = ""
    '            End If

    '    End Select
    'Else
    '    '数量が0の時にここを通るとエラーになる
    '    'DGV購入品入力.Rows(e.RowIndex).Cells("予算額").Value = ""
    '    'DGV購入品入力.Rows(e.RowIndex).Cells("金額").Value = ""
    'End If

    '入力の手間をなくすようログインアカウントの名前が自動で出る入力補助機能をつけていたが、ただの親切で付けているなら不要とのことでコメントアウト
    'If DGV購入品入力.Rows(e.RowIndex).Cells("部門ID").Value IsNot DBNull.Value And DGV購入品入力.Rows(e.RowIndex).Cells("購入者").Value Is DBNull.Value Then
    '    DGV購入品入力.Rows(e.RowIndex).Cells("購入者").Value = Select_Form._氏名
    'End If


    'If (str編集列名 = "印刷日" Or str編集列名 = "発注日" Or str編集列名 = "入荷日") And IsDate(e.Value) Then
    'e.Value = Format(DateTime.Parse(e.Value), "yyyy/MM/dd")
    'End If
    'ロックされている行が一目で分かるように色を変えていたが、エクセルの購入依頼もロック行見分けつかないから色分けは不要とのことでコメントアウト
    'If DGV購入品入力(e.ColumnIndex, e.RowIndex).ReadOnly Then
    '    DGV購入品入力(e.ColumnIndex, e.RowIndex).Style.ForeColor = Color.Gray
    'End If
    'End Sub


    'セルが入力フォーカスを失い、内容の検証が有効になった場合に発生します
    Private Sub DGV購入品入力_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DGV購入品入力.CellValidating
        If Flg再表示 = False Then
            'Dim dt As DataTable = DGV購入品入力.DataSource
            'For Each ROW As DataRow In dt.Rows
            'Next
            Dim dgv As DataGridView = CType(sender, DataGridView)
            Dim str編集列名 As String = dgv.CurrentCell.OwningColumn.Name

            Dim temp As String = kc.ns(e.FormattedValue)
            If temp = "" Then Exit Sub
            If c日付セル IsNot Nothing Then
                'CellValueChangedに移動
                'If DGV購入品入力.Columns(e.ColumnIndex).HeaderText = "登録番号" And _
                '    nz(DGV購入品入力.Rows(e.RowIndex).Cells("承認済").Value) <> 1 And nz(DGV購入品入力.Rows(e.RowIndex).Cells("経理検査").Value) <> 1 Then
                '    Get_製品登録情報(e.RowIndex, dgv.CurrentCell.EditedFormattedValue.ToString)
                'End If

                If (str編集列名 = "印刷日" Or str編集列名 = "発注日" Or str編集列名 = "入荷日") And IsDate(temp) = False Then
                    MsgBox("入力日付エラー")


                    'If dgv.Rows(e.RowIndex).Cells(e.ColumnIndex).ReadOnly = False Then '2022/11/02 ekawai del ReadOnly使わないように指示があったため
                    '2022/1/26
                    'Cell_ValidatingイベントをCellValueChangedに統合するよう指示があったが
                    '日付以外の文字が入った場合、CellValueChangedを通る前に型エラーが出てしまうので、このイベントでエラー処理するしかないという結論になった
                    e.Cancel = True

                    'End If


                    Exit Sub
                    '2021/09/06 ここでFormatしても反映されなかったからCellFormattingでやることにした
                    'Else
                    '    dgv.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Format(DateTime.Parse(temp), "yyyy/MM/dd")

                End If

            End If
        End If
    End Sub
    '2022/01/26 ekawai CellValueChangeに集約したほうがいいとのことでコメントアウト

    'Private Sub DGV購入品入力_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DGV購入品入力.CellValueChanged
    '    Dim dgv As DataGridView = CType(sender, DataGridView)
    '    Dim str編集列名 As String = dgv.CurrentCell.OwningColumn.Name
    '    Dim i金額 As Integer
    '    Dim i数量 As Integer

    '    i金額 = nz(dgv("予算単価", e.RowIndex).Value)
    '    i数量 = nz(dgv("数量", e.RowIndex).Value)

    '    Select Case str編集列名
    '        Case "数量"
    '            RemoveHandler DGV購入品入力.CellValueChanged, AddressOf DGV購入品入力_CellValueChanged
    '            dgv("予算額", e.RowIndex).Value = i金額 * i数量
    '            AddHandler DGV購入品入力.CellValueChanged, AddressOf DGV購入品入力_CellValueChanged
    '        Case "予算単価"

    '        Case "税抜単価"


    '    End Select
    'End Sub
    Private Sub DGV購入品入力_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DGV購入品入力.CellEnter
        'Dim dt As DataTable = DGV購入品入力.DataSource
        'For Each ROW As DataRow In dt.Rows
        '    Debug.Print("cellenter")
        '    Debug.Print(ROW("品名"))
        '    Debug.Print(ROW("品名", DataRowVersion.Original))
        '    Debug.Print(ROW("品名", DataRowVersion.Current))
        '    Debug.Print(ROW.RowState)
        'Next

        Dim dgv As DataGridView = CType(sender, DataGridView)

        If TypeOf dgv.Columns(e.ColumnIndex) Is DataGridViewComboBoxColumn Then
            'コンボボックスのドロップダウンリストが一回のクリックで表示されるようにする
            SendKeys.Send("{F4}")
            If dgv.Columns(e.ColumnIndex).Name = "科目" Or dgv.Columns(e.ColumnIndex).Name = "部門" Or dgv.Columns(e.ColumnIndex).Name = "ワークコード" Then
                変更前Value = dgv.Rows(e.RowIndex).Cells(e.ColumnIndex).EditedFormattedValue

            End If
        End If
  
    End Sub


    Private Sub DGV購入品入力_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DGV購入品入力.EditingControlShowing
        'Dim dt As DataTable = DGV購入品入力.DataSource
        'For Each ROW As DataRow In dt.Rows
        '    Debug.Print("editingcontrolshowing")
        '    Debug.Print(ROW("品名"))
        '    Debug.Print(ROW("品名", DataRowVersion.Original))
        '    Debug.Print(ROW("品名", DataRowVersion.Current))
        '    Debug.Print(ROW.RowState)
        'Next


        Dim dgv As DataGridView = CType(sender, DataGridView)
        Dim str編集列名 As String = dgv.CurrentCell.OwningColumn.Name

        'If str編集列名 = "決裁候補" Or str編集列名 = "仕入先候補" Or str編集列名 = "購入者候補" Or str編集列名 = "ワークコード" Or str編集列名 = "科目" Or str編集列名 = "部門" Then　'2021/04/12 ekawai del 
        '支払区分以外のコンボボックスだったら
        If TypeOf e.Control Is DataGridViewComboBoxEditingControl And str編集列名 <> "支払区分" Then '2021/04/12 ekawai add S
            '編集のために表示されているコントロールを取得
            Me.dataGridViewComboBox = CType(e.Control, DataGridViewComboBoxEditingControl)
            'If str編集列名 = "科目" Or str編集列名 = "部門" Or str編集列名 = "ワークコード" Then '2022/04/12 ekawai del
            If str編集列名 = "科目" Or str編集列名 = "部門" Or str編集列名 = "ワークコード" Or str編集列名 = "仕入先候補" Then '2021/04/12 ekawai add S
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
            コンボイベントハンドラ追加()

            EventFlg = True
        Else

            If TypeOf e.Control Is DataGridViewTextBoxEditingControl Then
                'DGrVTBEC = CType(e.Control, DataGridViewTextBoxEditingControl)
                '編集のために表示されているコントロールを取得
                Dim tb As DataGridViewTextBoxEditingControl = _
                    CType(e.Control, DataGridViewTextBoxEditingControl)

                '数字(プラスのみ)
                RemoveHandler tb.KeyPress, AddressOf dataGridViewTextBox_KeyPress
                If str編集列名 = "科目番号" Then 'TD_poでもTM_kamokuでもint型だから正の整数しか入らないように制御
                    'KeyPressイベントハンドラを追加
                    AddHandler tb.KeyPress, AddressOf dataGridViewTextBox_KeyPress

                End If
                '数字(マイナス許可)
                RemoveHandler tb.KeyPress, AddressOf dataGridViewTextBox_KeyPressMinus
                If str編集列名 = "数量" Or str編集列名 = "税抜単価" Or str編集列名 = "予算単価" Then
                    '数字の時はCtrl+Vできないように制御しているので右クリックの貼り付けもできないようにした
                    tb.ContextMenu = New ContextMenu
                    'KeyPressイベントハンドラを追加
                    AddHandler tb.KeyPress, AddressOf dataGridViewTextBox_KeyPressMinus
                End If
                '日付
                RemoveHandler tb.KeyPress, AddressOf dataGridViewTextBox_KeyPressIsDate
                If str編集列名 = "印刷日" Or str編集列名 = "発注日" Or str編集列名 = "入荷日" Then
                    'KeyPressイベントハンドラを追加
                    AddHandler tb.KeyPress, AddressOf dataGridViewTextBox_KeyPressIsDate
                End If

            End If

        End If
    End Sub


    'Private Sub DGV購入_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DGV購入.CellValidating
    '    Dim dgv As DataGridView = CType(sender, DataGridView)
    '    If (dgv.Columns(e.ColumnIndex).Name = "仕入先" Or dgv.Columns(e.ColumnIndex).Name = "") AndAlso _
    '        TypeOf dgv.Columns(e.ColumnIndex) Is DataGridViewComboBoxColumn Then
    '        Dim cbc As DataGridViewComboBoxColumn = _
    '            CType(dgv.Columns(e.ColumnIndex), DataGridViewComboBoxColumn)
    '        If Not cbc.Items.Contains(e.FormattedValue) Then
    '            cbc.Items.Add(e.FormattedValue)
    '        End If
    '        dgv(e.ColumnIndex, e.RowIndex).Value = e.FormattedValue
    '    End If
    'End Sub

    'Private Sub DGV購入_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DGV購入.EditingControlShowing

    '    If TypeOf e.Control Is DataGridViewComboBoxEditingControl Then
    '        Dim dgv As DataGridView = CType(sender, DataGridView)
    '        If dgv.CurrentCell.OwningColumn.Name = "仕入先" Then
    '            Dim cmb As DataGridViewComboBoxEditingControl
    '            cmb = CType(e.Control, DataGridViewComboBoxEditingControl)
    '            cmb.DropDownStyle = ComboBoxStyle.DropDown
    '        End If

    '    End If
    'End Sub

    Private Sub Get_製品登録情報(i行番号 As Integer, No As String)
        Dim strSql As String = ""

        strSql = "select * from TM_製品"
        strSql &= vbCrLf & "where 部署 = '" & Select_Form.s部署 & "'"
        strSql &= vbCrLf & "and 登録番号 = '" & No & "'"
        Dim DT製品 As DataTable = dCon.DataSet(strSql, "DT").Tables(0)

        If DT製品.Rows.Count = 1 Then
            DGV購入品入力.Rows(i行番号).Cells("部門ID").Value = DT製品.Rows(0)("部門ID")
            DGV購入品入力.Rows(i行番号).Cells("購入者").Value = DT製品.Rows(0)("購入者")
            DGV購入品入力.Rows(i行番号).Cells("品名").Value = DT製品.Rows(0)("品名")
            DGV購入品入力.Rows(i行番号).Cells("型式").Value = DT製品.Rows(0)("型式")
            DGV購入品入力.Rows(i行番号).Cells("仕入先ID").Value = DT製品.Rows(0)("仕入先ID")
            DGV購入品入力.Rows(i行番号).Cells("仕入先").Value = DT製品.Rows(0)("仕入先")
            DGV購入品入力.Rows(i行番号).Cells("予算単価").Value = DT製品.Rows(0)("予算単価")
            DGV購入品入力.Rows(i行番号).Cells("数量").Value = DT製品.Rows(0)("数量")
            DGV購入品入力.Rows(i行番号).Cells("支払区分").Value = DT製品.Rows(0)("支払区分")
            DGV購入品入力.Rows(i行番号).Cells("科目ID").Value = DT製品.Rows(0)("科目ID")
            DGV購入品入力.Rows(i行番号).Cells("購入理由").Value = DT製品.Rows(0)("購入理由")
            DGV購入品入力.Rows(i行番号).Cells("購入ID").Value = getNext購入ID()
        End If
    End Sub
    Private Function メイン更新() As Boolean
        Dim 更新対象Flg As Boolean
        Dim 同時更新Flg As Boolean

        Dim strSql As String = ""
        Dim iエラー行 As Integer
        Dim dv As DataRowVersion

        Dim dt As DataTable = DGV購入品入力.DataSource
        Dim tran As System.Data.SqlClient.SqlTransaction = Nothing
        Dim s購入ID As String
        Dim DT同時更新 As New DataTable
        '空のテーブルを作る(列を先に入れておく)
        For Each col As DataColumn In dt.Columns
            DT同時更新.Columns.Add(col.ColumnName)
        Next


        更新対象Flg = False
        同時更新Flg = False
        If dt.Rows.Count > 0 Then
            If dt.Rows.Count = DGV購入品入力.Rows.Count Then
                '空行がdtの行としてカウントされてしまうことがある。仕入先のセル3回クリックした時とか。
                '更新チェックの対象にしたくないし後のループにも入れたくないので削除する
                dt.Rows(dt.Rows.Count - 1).Delete()
            End If

            '更新前のエラーチェック
            For Each row As DataRow In dt.Rows

                dv = DataRowVersion.Current
                Select Case row.RowState
                    Case DataRowState.Unchanged
                        '変更なし⇒次へ進む 
                        Continue For
                    Case DataRowState.Modified, DataRowState.Added
                        '更新 or 追加　⇒エラーチェック
                        iエラー行 = 必須項目Check(dt.Rows.IndexOf(row) + 1, row, dv, "購入品入力", "")
                        If iエラー行 > 0 Then
                            'エラー行にスクロール
                            DGV購入品入力.FirstDisplayedScrollingRowIndex = iエラー行 - 1
                            DGV購入品入力.Rows(iエラー行 - 1).Selected = True
                            'TD_poへの書き込みを中止する
                            Return False
                        End If

                End Select

            Next

            tran = dCon.Connection.BeginTransaction
            'DataRowStateで変更か新規か判断したいので、DGVではなくDatatableをループしている
            For Each row As DataRow In dt.Rows
                dv = DataRowVersion.Current
                Try
                    Select Case row.RowState


                        Case DataRowState.Unchanged
                            '次へ進む 
                            Continue For
                        Case DataRowState.Modified
                            '更新の場合
                            '同時編集対策
                            'Select Case CheckID(TargetID)
                            '    Case "削除行", "重複行"
                            If CheckID(row("id", dv)) <> "更新対象" Then
                                Dim DR As DataRow
                                同時更新Flg = 1
                                DR = DT同時更新.NewRow
                                '重複更新行の情報をテーブルに格納
                                For Each col As DataColumn In dt.Columns
                                    DR(col.ColumnName) = row(col.ColumnName, dv)
                                Next
                                DT同時更新.Rows.Add(DR)
                                '次の行へ
                                Continue For
                                'Case "更新対象"
                            Else

                                'Check_DGV(row, dv, "購入品入力")
                                strSql = "UPDATE TD_po"
                                strSql &= vbCrLf & "SET"
                                '部署名：必須
                                strSql &= vbCrLf & "group_name = " & kc.SQ(Select_Form.s部署)
                                '品名：必須
                                strSql &= vbCrLf & ",part_name = " & kc.SQ(row("品名", dv))
                                ''発注日：空白可
                                strSql &= vbCrLf & ",order_date = " & kc.nn(row("発注日", dv))
                                '入荷日:空白可★まとめ時は必須
                                strSql &= vbCrLf & ", arrival_date = " & kc.nn(row("入荷日", dv))
                                '印刷日：空白可 再印刷したい時ユーザーが手動で消す可能性あるからnullを書き込める必要がある
                                strSql &= vbCrLf & ", 印刷日 = " & kc.nn(row("印刷日", dv))
                                '部門ID：★必須
                                strSql &= vbCrLf & ", division_id = " & kc.SQ(row("部門ID", dv))
                                '購入者：★必須
                                strSql &= vbCrLf & ", employee = " & kc.SQ(row("購入者", dv))
                                '型式：空白可
                                '改行コードを取る 旧まとめブックの仕様を継承 なぜ型式だけ改行取るのかは謎
                                If row("型式", dv) Is DBNull.Value Then
                                    strSql &= vbCrLf & ", part_number = NULL"
                                Else
                                    strSql &= vbCrLf & ", part_number = " & kc.SQ(row("型式", dv)).Replace(Chr(13), "").Replace(Chr(10), "")
                                End If
                                '税抜き単価：空白可★まとめ時は必須
                                strSql &= vbCrLf & ", unit_price = " & kc.nn(row("税抜単価", dv))
                                '数量：★必須
                                strSql &= vbCrLf & ", number = " & row("数量", dv)
                                '科目ID：★必須
                                strSql &= vbCrLf & ", kamoku_id = " & row("科目ID", dv)
                                '仕入先ID:空白可
                                strSql &= vbCrLf & ", vendor_id = " & kc.nn(row("仕入先ID", dv))
                                '仕入先：★必須
                                strSql &= vbCrLf & ", vendor = " & kc.SQ(row("仕入先", dv))
                                '支払区分：★必須
                                strSql &= vbCrLf & ", pay_code = " & kc.SQ(row("pay_code", dv))
                                '購入理由：空白可
                                strSql &= vbCrLf & ", remark = " & kc.nn(row("購入理由", dv))
                                'ワークコード：空白可
                                strSql &= vbCrLf & ", work_code = " & kc.nn(row("ワークコード", dv))
                                '整理番号：空白可
                                strSql &= vbCrLf & ", 整理番号 = " & kc.nn(row("整理番号", dv))
                                '棚番号：空白可  
                                strSql &= vbCrLf & ", tana_bango = " & kc.nn(row("棚番号", dv))
                                '決裁:空白可　保留と否認以外なら承認済をTRUEにする
                                If row("決裁", dv) Is DBNull.Value Then
                                    strSql &= vbCrLf & ", 決裁 = NULL"
                                Else
                                    strSql &= vbCrLf & ", 決裁 = " & kc.SQ(row("決裁", dv))
                                    '決裁が空白でないということは承認済の可能性があるのでチェック
                                    If row("決裁", dv) <> "保留" And row("決裁", dv) <> "否認" Then
                                        'strSql &= vbCrLf & ", 承認済 = 1"
                                        strSql &= vbCrLf & ", 承認済 = 'True'"
                                    Else
                                        'strSql &= vbCrLf & ", 承認済 = 0"
                                        strSql &= vbCrLf & ", 承認済 = 'False'"
                                    End If


                                End If
                                '登録番号：空白可
                                strSql &= vbCrLf & ", 登録番号 = " & kc.nn(row("登録番号", dv))
                                '予算単価：★必須
                                strSql &= vbCrLf & ", 予算単価 = " & row("予算単価", dv)

                                If row("経理検査", dv) Is DBNull.Value Then
                                    'strSql &= vbCrLf & ", 経理検査 = 0"
                                    strSql &= vbCrLf & ", 経理検査 = 'False'"
                                Else
                                    strSql &= vbCrLf & ", 経理検査 = " & kc.SQ(row("経理検査", dv))
                                    'If row("経理検査", dv) = True Then
                                    '    strSql &= vbCrLf & ", 経理検査 = 1"
                                    'Else
                                    '    strSql &= vbCrLf & ", 経理検査 = 0"
                                    'End If

                                End If

                                '改訂番号：空白可
                                strSql &= vbCrLf & ", 改訂番号 = " & kc.nn(row("改訂番号", dv))
                                '備考1：空白可
                                strSql &= vbCrLf & ", 備考1 = " & kc.nn(row("備考1", dv))
                                'updateStamp
                                strSql &= vbCrLf & ", updateStamp = " & kc.SQ(Now)
                                strSql &= vbCrLf & ", 更新者 = " & kc.SQ(Select_Form._UserName)
                                strSql &= vbCrLf & ", 更新PC名 = " & kc.SQ(Select_Form._PC名)
                                strSql &= vbCrLf & ", 更新バージョン = " & kc.SQ(cVer)
                                '2022/11/02 ekawai add S
                                '入力者が行を編集しなかったときはフラグがFalseにならない(つまり数量と型式が編集できる)が問題ないか確認したところ、その場合は仕方ないとのことだった
                                '承認済みと経理検査の行でなければ後から編集できても特に問題ないし、最終的に経理が経理検査フラグつけるので、翌月初までにはコピーフラグは全てFalseになるはず
                                If kc.nz(row("コピーフラグ", dv)) = True Then
                                    strSql &= vbCrLf & ", コピーフラグ = 'False'"
                                End If
                                '2022/11/02 ekawai add E
                                strSql &= vbCrLf & "WHERE id = " & row("id", dv)

                                'tran = dCon.Connection.BeginTransaction
                                If dCon.ExecuteSqlMW(tran, strSql) = False Then
                                    tran.Rollback()
                                    'エラー行にスクロール
                                    DGV購入品入力.FirstDisplayedScrollingRowIndex = dt.Rows.IndexOf(row)
                                    DGV購入品入力.Rows(dt.Rows.IndexOf(row)).Selected = True

                                    MsgBox(dt.Rows.IndexOf(row) + 1 & "行目 UPDATE失敗")
                                    Return False
                                Else
                                    更新対象Flg = True
                                End If
                                'End Select
                                End If
                        Case DataRowState.Added

                                '新規行の場合


                                '追加前のチェック()

                                'Check_DGV(row, dv, "購入品入力")
                                'チェック問題なければ購入IDを取得する
                                s購入ID = getNext購入ID()
                                'Dim id As New String("0"c, 8 - num.Length)
                                's購入ID = id & num

                                strSql = "INSERT INTO"
                                strSql &= vbCrLf & "TD_po("
                                strSql &= vbCrLf & "group_name"
                                strSql &= vbCrLf & ",part_name"
                                strSql &= vbCrLf & ",order_date"
                                strSql &= vbCrLf & ",arrival_date"
                                strSql &= vbCrLf & ",印刷日"
                                strSql &= vbCrLf & ",division_id"
                                strSql &= vbCrLf & ",employee"
                                strSql &= vbCrLf & ",part_number"
                                strSql &= vbCrLf & ",unit_price"
                                strSql &= vbCrLf & ",number"
                                strSql &= vbCrLf & ",kamoku_id"
                                strSql &= vbCrLf & ",vendor_id"
                                strSql &= vbCrLf & ",vendor"
                                strSql &= vbCrLf & ",pay_code"
                                strSql &= vbCrLf & ",remark"
                                strSql &= vbCrLf & ",work_code"
                                strSql &= vbCrLf & ",tana_bango"
                                strSql &= vbCrLf & ",承認済"
                                strSql &= vbCrLf & ",決裁"
                                strSql &= vbCrLf & ",登録番号"
                                strSql &= vbCrLf & ",予算単価"
                                strSql &= vbCrLf & ",経理検査"
                                strSql &= vbCrLf & ",整理番号"
                                strSql &= vbCrLf & ",改訂番号"
                                'strSql &= vbCrLf & ",営業所"
                                strSql &= vbCrLf & ",購入ID"
                                strSql &= vbCrLf & ",insertStamp"
                                strSql &= vbCrLf & ",備考1"
                                strSql &= vbCrLf & ",作成者"
                                strSql &= vbCrLf & ",作成PC名"
                                strSql &= vbCrLf & ",作成バージョン"
                                strSql &= vbCrLf & ") VALUES ("
                                strSql &= vbCrLf & kc.SQ(Select_Form.s部署)
                                strSql &= vbCrLf & "," & kc.SQ(row("品名", dv))
                                strSql &= vbCrLf & "," & kc.nn(row("発注日", dv))
                                strSql &= vbCrLf & "," & kc.nn(row("入荷日", dv))
                                strSql &= vbCrLf & "," & kc.nn(row("印刷日", dv))
                                strSql &= vbCrLf & "," & kc.SQ(row("部門ID", dv))
                                strSql &= vbCrLf & "," & kc.SQ(row("購入者", dv))
                                strSql &= vbCrLf & "," & kc.nn(row("型式", dv)).Replace(Chr(13), "").Replace(Chr(10), "")
                                strSql &= vbCrLf & "," & kc.nn(row("税抜単価", dv))
                                strSql &= vbCrLf & "," & row("数量", dv)
                                strSql &= vbCrLf & "," & row("科目ID", dv)
                                '現金の時は仕入先IDを空白にする
                                'If kc.ns(row("仕入先ID", dv)) <> "" Then
                                '    If kc.SQ(row("pay_code", dv)) = "現金" Then
                                '        strSql &= vbCrLf & ", null"
                                '    Else
                                '        strSql &= vbCrLf & ", " & kc.SQ(row("仕入先ID", dv))
                                '    End If
                                'Else
                                '    strSql &= vbCrLf & ", null"
                                'End If
                                ''2021/04/12 ekawai del S↓------------------ 現金の時は仕入先IDあってもNULL入れていたけどやめることにした→Main_Form起動時に仕入先IDが空白の行にID振る処理追加したから矛盾するため
                                '仕入先IDが空白の時はNULLを入れる…旧まとめブックの仕様を継承。理由は不明。過去データがそうやって入っているから合わせているだけ。
                                'If row("仕入先ID", dv) Is DBNull.Value Then
                                '    strSql &= vbCrLf & ", null"
                                'Else
                                '    '空白じゃない時は、支払区分が現金だったらNULLを入れる
                                '    If kc.SQ(row("pay_code", dv)) = "現金" Then
                                '        strSql &= vbCrLf & ", null"
                                '    Else
                                '        strSql &= vbCrLf & ", " & kc.SQ(row("仕入先ID", dv))
                                '    End If
                                'End If
                                '2021/04/12 ekawai del E↑------------------
                                strSql &= vbCrLf & "," & kc.nn(row("仕入先ID", dv)) '2021/04/12 ekawai add 
                                strSql &= vbCrLf & "," & kc.SQ(row("仕入先", dv))
                                strSql &= vbCrLf & "," & kc.SQ(row("pay_code", dv))
                                strSql &= vbCrLf & "," & kc.nn(row("購入理由", dv))
                                strSql &= vbCrLf & "," & kc.nn(row("ワークコード", dv))
                                strSql &= vbCrLf & "," & kc.nn(row("棚番号", dv))
                                'If row("決裁", dv) IsNot DBNull.Value Then
                                '    If row("決裁", dv) <> "否認" And row("決裁", dv) <> "保留" Then
                                '        strSql &= vbCrLf & ",'True'"
                                '    Else
                                '        strSql &= vbCrLf & ",'False'"
                                '    End If
                                'Else
                                '    strSql &= vbCrLf & ",'False'"
                                'End If

                                If row("決裁", dv) Is DBNull.Value Then
                                    '決裁列が空白だったら承認済列はFalse
                                    strSql &= vbCrLf & ",'False'"
                                Else
                                    '決裁列が空白じゃなかったら
                                    If row("決裁", dv) <> "否認" And row("決裁", dv) <> "保留" Then
                                        '否認と保留以外(稟議承認or人の名前)ならTRUE
                                        strSql &= vbCrLf & ",'True'"
                                    Else
                                        '否認と保留だったらFALSE
                                        strSql &= vbCrLf & ",'False'"
                                    End If

                                End If


                                strSql &= vbCrLf & "," & kc.nn(row("決裁", dv))
                                strSql &= vbCrLf & "," & kc.nn(row("登録番号", dv))
                                strSql &= vbCrLf & "," & row("予算単価", dv)
                                'strSql &= vbCrLf & "," & kc.nn(row("経理検査", dv))
                                'If row("経理検査", dv) IsNot DBNull.Value Then
                                '    strSql &= vbCrLf & "," & kc.SQ(row("経理検査", dv))
                                'Else
                                '    strSql &= vbCrLf & ",'False'"
                                'End If
                                If row("経理検査", dv) Is DBNull.Value Then
                                    strSql &= vbCrLf & ",'False'"
                                Else
                                    strSql &= vbCrLf & "," & kc.SQ(row("経理検査", dv))
                                End If

                                strSql &= vbCrLf & "," & kc.nn(row("整理番号", dv))
                                strSql &= vbCrLf & "," & kc.nn(row("改訂番号", dv))
                                'strSql &= vbCrLf & "," & nr(row("営業所", dv))
                                strSql &= vbCrLf & "," & kc.SQ(s購入ID)
                                strSql &= vbCrLf & "," & kc.SQ(Now)
                                strSql &= vbCrLf & "," & kc.nn(row("備考1", dv))
                                strSql &= vbCrLf & "," & kc.SQ(Select_Form._UserName)
                                strSql &= vbCrLf & "," & kc.SQ(Select_Form._PC名)
                                strSql &= vbCrLf & "," & kc.SQ(cVer)

                                strSql &= vbCrLf & ")"

                                'tran = dCon.Connection.BeginTransaction
                                If dCon.ExecuteSqlMW(tran, strSql) = False Then
                                    tran.Rollback()
                                    'エラー行にスクロール
                                    DGV購入品入力.FirstDisplayedScrollingRowIndex = dt.Rows.IndexOf(row)
                                    DGV購入品入力.Rows(dt.Rows.IndexOf(row)).Selected = True
                                    MsgBox(dt.Rows.IndexOf(row) + 1 & "行目 INSERT失敗")
                                    Return False
                                End If
                                更新対象Flg = True
                                '削除行
                        Case DataRowState.Deleted
                                '2022/02/08 ekawai del S
                                'If CheckID(row("id", DataRowVersion.Original)) <> "更新対象" Then
                                '    Dim DR As DataRow
                                '    同時更新Flg = 1
                                '    DR = DT同時更新.NewRow
                                '    '重複更新行の情報をテーブルに格納
                                '    For Each col As DataColumn In dt.Columns
                                '        DR(col.ColumnName) = row(col.ColumnName, DataRowVersion.Original)
                                '    Next
                                '    DT同時更新.Rows.Add(DR)
                                '    '次の行へ
                                '    Continue For
                                '    'Case "更新対象"
                                'Else

                                '    更新対象Flg = True
                                '    'TD_poにidがあれば削除する
                                '    'Deleted 行には Current 行バージョンがないため、列値にアクセスするときに DataRowVersion.Original を渡す必要があります。
                                '    If nz(row("id", DataRowVersion.Original)) <> 0 Then
                                '        strSql = "DELETE FROM TD_po"
                                '        strSql &= vbCrLf & "WHERE id = " & row("id", DataRowVersion.Original)

                                '        If dCon.ExecuteSqlMW(tran, strSql) = False Then
                                '            tran.Rollback()
                                '            MsgBox("id=" & row("id", DataRowVersion.Original) & " 削除失敗")
                                '            Return False
                                '        Else
                                '        End If

                                '    Else
                                '        'なければ画面更新したら勝手に消えるから何もしない
                                '    End If


                                'End If
                                '2022/02/08 ekawai del E
                    End Select
                Catch ex As Exception
                    MsgBox(ex.Message & "更新できませんでした。" & vbCrLf & ex.Message)
                    If tran IsNot Nothing Then
                        tran.Rollback()
                    End If
                    Return False
                End Try

                's購入ID = Integer.Parse(s購入ID) + 1
            Next


            If 更新対象Flg = True Then
                tran.Commit()
                dt.AcceptChanges() 'Rowstateをunchangedに更新する

                DGV購入品入力_表示()
                DGV購入品入力_セル設定()
                If 同時更新Flg Then
                    MsgBox("他のユーザーが変更したため、一部更新できない行がありました")
                    Excel出力(DT同時更新) '自分が変更した行が削除した行か区別つかないが問題ないとのこと
                Else
                    MsgBox("更新成功")
                End If


            Else
                tran.Rollback()
                If 同時更新Flg Then
                    MsgBox("他のユーザーが変更したため、更新できませんでした")
                    Excel出力(DT同時更新)
                Else
                    MsgBox("更新対象がありません")
                End If

            End If
            SwEditSave = False
            tran.Dispose()
            Return True '更新対象ない場合も失敗したわけではないからTRUEを返す
        Else
            'DGVが0行の時　更新成功Flg = Falseのまま
        End If



    End Function
    Sub Excel出力(DT As DataTable)
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
        With DT
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
    '自分がフォームを最後に更新したよりあとに、他の人が同じ行に変更を加えていたら更新できないよう制御
    Function CheckID(id As String) As String
        Dim mCon As New merrweth_init_DbConnection
        Dim tran As System.Data.SqlClient.SqlTransaction = Nothing
        Dim strSql As String
        Dim DT_最新情報 As DataTable
        strSql = ""
        strSql = strSql & vbCrLf & "SELECT * FROM TD_po"
        strSql = strSql & vbCrLf & "WHERE id =" & id
        'データベース接続、トランザクション開始（行ロックする）
        tran = mCon.Connection.BeginTransaction


        DT_最新情報 = mCon.GetSqlDataTable(tran, strSql)
        tran.Commit()

        '新規登録後変更がない行にはupdateStampがないので比較できないから最後にフォームを更新した時間(最終更新時刻)を使っている
        If DT_最新情報.Rows.Count = 1 Then
            'UpdateStampを取得する
            If DT_最新情報.Rows(0)("updateStamp") IsNot DBNull.Value Then
                If 最終更新時刻 < DT_最新情報.Rows(0)("updateStamp") Then
                    Return "重複行"
                Else
                    Return "更新対象"
                End If
            Else
                Return "更新対象"

            End If
        Else
            Return "削除行"
        End If
    End Function

    Private Sub btn更新_Click(sender As Object, e As EventArgs) Handles btn更新.Click
        '開発途中で画面書き換える前に必ず更新する仕様に変えるよう指示があったので
        '更新失敗か成功か判断できるようにファンクションにすることにした
        'このプロシージャでは成功か失敗かで処理は分岐しない。他のところで使う。
        If メイン更新() Then
            CloseFlg = True '右上の閉じるボタン押した時に使うフラグ
        Else
            CloseFlg = False
        End If
    End Sub

    'Public Function 製品更新(dt As DataTable)
    '    'テーブルスタイルの取得

    '    'DataRowViewを使いDataGridViewの現在の行（または任意の行）からソース元のDataTableのDataRowを取得します。
    '    'Dim dgr As System.Windows.Forms.DataGridViewRow = Me.DGV購入.CurrentRow
    '    'Dim drv As System.Data.DataRowView = CType(dgr.DataBoundItem, System.Data.DataRowView)
    '    'Dim dr As System.Data.DataRow = CType(drv.Row, System.Data.DataRow)
    '    'Dim tran As System.Data.SqlClient.SqlTransaction = Nothing
    '    Dim iエラー行 As Integer
    '    Dim 更新対象Flg As String
    '    Dim strSql As String = ""

    '    Dim dv As DataRowVersion

    '    Dim tran As System.Data.SqlClient.SqlTransaction = Nothing

    '    更新対象Flg = False
    '    '更新前のエラーチェック
    '    For Each row As DataRow In dt.Rows
    '        dv = DataRowVersion.Current
    '        Select Case row.RowState
    '            Case DataRowState.Unchanged
    '                '変更なし⇒次へ進む 
    '                Continue For
    '            Case DataRowState.Modified, DataRowState.Added
    '                '更新 or 追加　⇒エラーチェック
    '                iエラー行 = 必須項目Check(dt.Rows.IndexOf(row) + 1, row, dv, "製品")
    '                If iエラー行 > 0 Then
    '                    'エラー行にスクロール
    '                    DGV製品.FirstDisplayedScrollingRowIndex = iエラー行
    '                    'TD_poへの書き込みを中止する
    '                    Return False
    '                End If

    '        End Select

    '    Next

    '    tran = dCon.Connection.BeginTransaction
    '    For Each row As DataRow In dt.Rows
    '        dv = DataRowVersion.Current
    '        Try

    '            Select Case row.RowState
    '                Case DataRowState.Unchanged
    '                    '次へ進む 
    '                    Continue For
    '                Case DataRowState.Modified
    '                    '更新の場合
    '                    '更新前のチェック
    '                    'Check_DGV(row, dv, "製品")
    '                    strSql = "UPDATE TM_製品"
    '                    strSql &= vbCrLf & "SET"
    '                    strSql &= vbCrLf & "部署 = " & SQ(Select_Form.s部署)
    '                    strSql &= vbCrLf & ",品名 = " & SQ(ns(row("品名", dv)))
    '                    strSql &= vbCrLf & ", 部門ID = " & SQ(ns(row("部門ID", dv)))
    '                    strSql &= vbCrLf & ", 購入者 = " & SQ(ns(row("購入者", dv)))
    '                    strSql &= vbCrLf & ", 型式 = " & SQ(ns(row("型式", dv)))
    '                    strSql &= vbCrLf & ", 予算単価 = " & nz(row("予算単価", dv))
    '                    strSql &= vbCrLf & ", 数量 = " & nz(row("数量", dv))
    '                    '科目ID(int型)
    '                    If row("科目ID", dv) IsNot DBNull.Value Then
    '                        strSql &= vbCrLf & ", 科目ID = " & row("科目ID", dv)
    '                    End If
    '                    strSql &= vbCrLf & ", 仕入先ID = " & SQ(ns(row("仕入先ID", dv)))
    '                    strSql &= vbCrLf & ", 仕入先 = " & SQ(ns(row("仕入先", dv)))
    '                    strSql &= vbCrLf & ", 支払区分 = " & SQ(ns(row("支払区分", dv)))
    '                    strSql &= vbCrLf & ", 購入理由 = " & SQ(ns(row("購入理由", dv)))
    '                    If row("登録番号", dv) IsNot DBNull.Value Then
    '                        strSql &= vbCrLf & ", 登録番号 = " & ns(SQ(row("登録番号", dv)))
    '                    End If
    '                    strSql &= vbCrLf & ", 備考 = " & SQ(ns(row("備考", dv)))

    '                    strSql &= vbCrLf & "WHERE UID = " & row("UID", dv)

    '                    If dCon.ExecuteSqlMW(tran, strSql) = False Then
    '                        tran.Rollback()
    '                        MsgBox("UPDATE失敗")
    '                        Return False
    '                    Else
    '                        更新対象Flg = True
    '                    End If

    '                Case DataRowState.Added

    '                    '新規行の場合


    '                    '追加前のチェック
    '                    'Check_DGV(row, dv, "製品")
    '                    strSql = "INSERT INTO"
    '                    strSql &= vbCrLf & "TM_製品("
    '                    strSql &= vbCrLf & "部署"
    '                    strSql &= vbCrLf & ",品名"
    '                    strSql &= vbCrLf & ",部門ID"
    '                    strSql &= vbCrLf & ",購入者"
    '                    strSql &= vbCrLf & ",型式"
    '                    strSql &= vbCrLf & ",予算単価"
    '                    strSql &= vbCrLf & ",数量"
    '                    strSql &= vbCrLf & ",科目ID"
    '                    strSql &= vbCrLf & ",仕入先ID"
    '                    strSql &= vbCrLf & ",仕入先"
    '                    strSql &= vbCrLf & ",支払区分"
    '                    strSql &= vbCrLf & ",購入理由"
    '                    strSql &= vbCrLf & ",登録番号"
    '                    strSql &= vbCrLf & ",備考"
    '                    strSql &= vbCrLf & ") VALUES ("
    '                    strSql &= vbCrLf & SQ(Select_Form.s部署)
    '                    strSql &= vbCrLf & "," & SQ(row("品名", dv))
    '                    strSql &= vbCrLf & "," & nr(row("部門ID", dv))
    '                    strSql &= vbCrLf & "," & nr(row("購入者", dv))
    '                    strSql &= vbCrLf & "," & nr(row("型式", dv))
    '                    strSql &= vbCrLf & "," & nr(row("予算単価", dv))
    '                    strSql &= vbCrLf & "," & nz(row("数量", dv))
    '                    strSql &= vbCrLf & "," & nz(row("科目ID", dv))
    '                    strSql &= vbCrLf & "," & nr(row("仕入先ID", dv))
    '                    strSql &= vbCrLf & "," & nr(row("仕入先", dv))
    '                    strSql &= vbCrLf & "," & nr(row("支払区分", dv))
    '                    strSql &= vbCrLf & "," & nr(row("購入理由", dv))
    '                    strSql &= vbCrLf & "," & nr(row("登録番号", dv))
    '                    strSql &= vbCrLf & "," & nr(row("備考", dv))
    '                    strSql &= vbCrLf & ")"


    '                    dCon.ExecuteSqlMW(tran, strSql)
    '                    If dCon.ExecuteSqlMW(tran, strSql) = False Then
    '                        tran.Rollback()
    '                        MsgBox("INSERT失敗")
    '                        Return False
    '                    End If
    '                    更新対象Flg = True

    '            End Select
    '        Catch ex As Exception
    '            MsgBox(ex.Message & "更新失敗")
    '            If tran IsNot Nothing Then
    '                tran.Rollback()
    '            End If
    '            Return False
    '        End Try

    '        'End If

    '    Next
    '    If 更新対象Flg = True Then
    '        tran.Commit()
    '        tran.Dispose()
    '        dt.AcceptChanges() 'Rowstateをunchangedに更新する

    '        MsgBox("更新成功")
    '        Return True
    '    Else
    '        MsgBox("更新対象がありません")
    '        Return True
    '    End If
    'End Function

    '各フォーム共通
    '元々はフォームごとにチェック分けていたが、変更があった時にバラバラのフォームにコードが存在すると管理が手間なので、統一した
    'しかし2021/10/7にユーザーテストで単価と数量は空白許可してほしいという意見があり、製品と見積に関しては単価と数量の空白を認めることになったのでSELECT分で分岐している。
    '他のフォームへのコピー前にもこのチェックを通るように指示があったので、コピー先という変数を増やした。
    Public Function 必須項目Check(i行番号 As Integer, r As DataRow, dv As DataRowVersion, form As String, コピー先 As String) As Integer
        '複数行エラーの時全てメッセージ出すときりがないので
        '上からチェックして最初にエラーが見つかった行番号だけ教える


        '初回入力必須項目 
        'If ns(r("部門ID", dv)) = "" Then
        '    MsgBox(i行番号 & "行目に部門を入力してください")
        '    Return i行番号
        'End If

        'If ns(r("購入者", dv)) = "" Then
        '    MsgBox(i行番号 & "行目に購入者を入力してください")
        '    Return i行番号
        'End If

        If kc.ns(r("品名", dv)) = "" Then
            MsgBox(i行番号 & "行目に品名を入力してください")
            Return i行番号
        End If

        If kc.ns(r("仕入先", dv)) = "" Then
            MsgBox(i行番号 & "行目に仕入先を入力してください")
            Return i行番号
        End If

        'If nz(r("科目ID", dv)) = 0 Then
        'MsgBox(i行番号 & "行目に科目IDを入力してください")
        'Return i行番号
        'End If

        '必須条件をフォームごとに変えた結果、フォームごとに分岐が必要になった。

        Select Case form


            Case "製品"
                If kc.ns(r("支払区分", dv)) = "" Then
                    MsgBox(i行番号 & "行目に支払区分を入力してください")
                    Return i行番号
                End If

                If kc.nz(r("科目ID", dv)) = 0 Then
                    MsgBox(i行番号 & "行目に科目IDを入力してください")
                    Return i行番号
                End If

                Select Case コピー先
                    'INSERTしているので必須項目の漏れは許されない
                    Case "見積"
                        If kc.ns(r("購入者", dv)) = "" Then
                            MsgBox(i行番号 & "行目に購入者を入力してください")
                            Return i行番号
                        End If

                        'Case "購入品入力"
                        '    If ns(r("支払区分", dv)) = "" Then
                        '        MsgBox(i行番号 & "行目に支払区分を入力してください")
                        '        Return i行番号
                        '    End If
                End Select

            Case "見積"
                If kc.ns(r("購入者", dv)) = "" Then
                    MsgBox(i行番号 & "行目に購入者を入力してください")
                    Return i行番号
                End If


                Select Case コピー先
                    'INSERTしているので必須項目の漏れは許されない
                    Case "製品"

                        If kc.nz(r("科目ID", dv)) = 0 Then
                            MsgBox(i行番号 & "行目に科目IDを入力してください")
                            Return i行番号
                        End If

                        If kc.ns(r("支払区分", dv)) = "" Then
                            MsgBox(i行番号 & "行目に支払区分を入力してください")
                            Return i行番号
                        End If

                    Case "購入品入力"
                        '見積から購入品入力はINSERTしているので、必須項目の漏れは許されない
                        If kc.nz(r("数量", dv)) = 0 Then
                            MsgBox(i行番号 & "行目に数量を入力してください")
                            Return i行番号
                        End If
                        If kc.nz(r("見積単価", dv)) = 0 Then
                            MsgBox(i行番号 & "行目に見積単価を入力してください")
                            Return i行番号
                        End If

                        If kc.nz(r("科目ID", dv)) = 0 Then
                            MsgBox(i行番号 & "行目に科目IDを入力してください")
                            Return i行番号
                        End If

                        If kc.ns(r("部門ID", dv)) = "" Then
                            MsgBox(i行番号 & "行目に部門を入力してください")
                            Return i行番号
                        End If

                        If kc.ns(r("支払区分", dv)) = "" Then
                            MsgBox(i行番号 & "行目に支払区分を入力してください")
                            Return i行番号
                        End If

                    Case Else
                        '何もしない
                End Select



            Case "購入品入力"
                If kc.ns(r("部門ID", dv)) = "" Then
                    MsgBox(i行番号 & "行目に部門を入力してください")
                    Return i行番号
                End If

                If kc.ns(r("購入者", dv)) = "" Then
                    MsgBox(i行番号 & "行目に購入者を入力してください")
                    Return i行番号
                End If

                If kc.nz(r("数量", dv)) = 0 Then
                    MsgBox(i行番号 & "行目に数量を入力してください")
                    Return i行番号
                End If
                If kc.nz(r("予算単価", dv)) = 0 Then
                    MsgBox(i行番号 & "行目に予算単価を入力してください")
                    Return i行番号
                End If
                If kc.ns(r("pay_code", dv)) = "" Then
                    MsgBox(i行番号 & "行目に支払区分を入力してください")
                    Return i行番号
                End If

                If kc.nz(r("科目ID", dv)) = 0 Then
                    MsgBox(i行番号 & "行目に科目IDを入力してください")
                    Return i行番号
                End If

        End Select

        Return -1


    End Function


    'Private Sub Check_DGV(r As DataRow, dv As DataRowVersion, form As String)
    '    '初回入力必須項目
    '    If ns(r("部門ID", dv)) = "" Then
    '        MsgBox("部門を入力してください")
    '        Throw New Exception
    '    End If

    '    If ns(r("購入者", dv)) = "" Then
    '        MsgBox("購入者を入力してください")
    '        Throw New Exception
    '    End If

    '    If ns(r("品名", dv)) = "" Then
    '        MsgBox("品名を入力してください")
    '        Throw New Exception
    '    End If

    '    If ns(r("仕入先", dv)) = "" Then
    '        MsgBox("仕入先を入力してください")
    '        Throw New Exception
    '    End If

    '    If nz(r("数量", dv)) = 0 Then
    '        MsgBox("数量を入力してください")
    '        Throw New Exception
    '    End If

    '    If nz(r("予算単価", dv)) = 0 Then
    '        MsgBox("予算単価を入力してください")
    '        Throw New Exception
    '    End If
    '    If form = "購入品入力" Then
    '        If ns(r("pay_code", dv)) = "" Then
    '            MsgBox("支払区分を入力してください")
    '            Throw New Exception
    '        End If
    '    Else
    '        If ns(r("支払区分", dv)) = "" Then
    '            MsgBox("支払区分を入力してください")
    '            Throw New Exception
    '        End If

    '    End If

    '    If nz(r("科目ID", dv)) = 0 Then
    '        MsgBox("科目IDを入力してください")
    '        Throw New Exception
    '    End If


    'End Sub

    '----------------------------------------------------------------------------------------------------------
    'ＮＳ
    '----------------------------------------------------------------------------------------------------------
    'Public Function ns(ByVal inData)
    '    Dim t As String
    '    'If inData Is System.DBNull.Value Then
    '    '    Return ""
    '    'End If

    '    If inData Is Nothing Then
    '        Return ""
    '    End If

    '    If inData Is System.DBNull.Value Then
    '        Return ""
    '    End If

    '    If inData.Equals(vbNull) Then
    '        t = ""
    '    Else
    '        If IsDBNull(inData) Then
    '            t = ""
    '        Else
    '            t = inData
    '        End If
    '    End If
    '    Return t
    'End Function
    ''----------------------------------------------------------------------------------------------------------
    ''ＮＺ
    ''----------------------------------------------------------------------------------------------------------
    'Public Function nz(ByVal inData)
    '    Dim t

    '    If inData Is System.DBNull.Value Then
    '        Return 0
    '    End If

    '    If inData Is Nothing Then
    '        Return 0
    '    End If

    '    If inData.Equals(vbNull) Then
    '        t = 0
    '    Else
    '        If IsDBNull(inData) Then
    '            t = 0
    '        Else
    '            If IsNumeric(inData) = False Then
    '                t = 0
    '            Else
    '                t = inData
    '            End If
    '        End If
    '    End If
    '    Return t
    'End Function
    'Public Function SQ(str As String) As String
    '    Return "'" & str & "'"
    ''End Function
    ''NULLの時の処理
    'Public Function nr(ByVal inData) As String
    '    Dim t
    '    If inData Is System.DBNull.Value Then
    '        t = "null"
    '    Else
    '        t = SQ(inData)
    '    End If
    '    Return t
    'End Function






    Private Sub Set製品情報()
        Dim drNew As DataRow
        drNew = DT購入品リスト.NewRow
        drNew("登録番号") = arr製品情報(0)
        drNew("部門ID") = arr製品情報(1)
        drNew("購入者") = arr製品情報(2)
        drNew("品名") = arr製品情報(3)
        drNew("型式") = arr製品情報(4)
        drNew("数量") = arr製品情報(5)
        drNew("予算単価") = arr製品情報(6)
        drNew("仕入先ID") = arr製品情報(7)
        drNew("仕入先") = arr製品情報(8)
        drNew("pay_code") = arr製品情報(9)
        drNew("科目ID") = arr製品情報(10)
        drNew("購入理由") = arr製品情報(11)
        drNew("購入ID") = getNext購入ID()
        DT購入品リスト.Rows.Add(drNew)
        SwEditSave = True


    End Sub

    'Private Sub Set見積情報()
    '    Dim strSql As String
    '    Dim tran As System.Data.SqlClient.SqlTransaction = Nothing

    '    'Dim drNew As DataRow
    '    'drNew = DT購入品リスト.NewRow
    '    'drNew("購入者") = arr見積情報(0)
    '    'drNew("品名") = arr見積情報(1)
    '    'drNew("型式") = arr見積情報(2)
    '    'drNew("数量") = arr見積情報(3)
    '    'drNew("予算単価") = arr見積情報(4)
    '    'drNew("仕入先ID") = arr見積情報(5)
    '    'drNew("仕入先") = arr見積情報(6)
    '    'drNew("pay_code") = arr見積情報(7)
    '    'drNew("科目ID") = arr見積情報(8)
    '    'drNew("購入理由") = arr見積情報(9)
    '    'drNew("部門ID") = arr見積情報(11)

    '    'DT購入品リスト.Rows.Add(drNew)
    '    strSql = "INSERT INTO"
    '    strSql &= vbCrLf & "TD_po("
    '    strSql &= vbCrLf & "group_name"
    '    strSql &= vbCrLf & ",employee"
    '    strSql &= vbCrLf & ",part_name"
    '    strSql &= vbCrLf & ",part_number"
    '    strSql &= vbCrLf & ",number"
    '    strSql &= vbCrLf & ",予算単価"
    '    strSql &= vbCrLf & ",vendor_id"
    '    strSql &= vbCrLf & ",vendor"
    '    strSql &= vbCrLf & ",pay_code"
    '    strSql &= vbCrLf & ",kamoku_id"
    '    strSql &= vbCrLf & ",remark"
    '    strSql &= vbCrLf & ",division_id"
    '    strSql &= vbCrLf & ",購入ID"
    '    strSql &= vbCrLf & ",insertStamp"
    '    strSql &= vbCrLf & ") VALUES ("
    '    strSql &= vbCrLf & SQ(Select_Form.s部署)
    '    strSql &= vbCrLf & "," & SQ(arr見積情報(0)) '購入者
    '    strSql &= vbCrLf & "," & SQ(arr見積情報(1)) '品名
    '    strSql &= vbCrLf & "," & nr(arr見積情報(2)).Replace(Chr(13), "").Replace(Chr(10), "") '型式
    '    strSql &= vbCrLf & "," & arr見積情報(3) '数量
    '    strSql &= vbCrLf & "," & arr見積情報(4) '予算単価
    '    If SQ(arr見積情報(6)) <> "" Then
    '        If SQ(arr見積情報(7)) = "現金" Then
    '            strSql &= vbCrLf & ", null"
    '        Else
    '            strSql &= vbCrLf & "," & nr(arr見積情報(5)) '仕入先ID
    '        End If
    '    Else
    '        strSql &= vbCrLf & ", null"
    '    End If
    '    strSql &= vbCrLf & "," & SQ(arr見積情報(6)) '仕入先名
    '    strSql &= vbCrLf & "," & SQ(arr見積情報(7)) '支払区分
    '    strSql &= vbCrLf & "," & SQ(arr見積情報(8)) '科目ID
    '    strSql &= vbCrLf & "," & nr(arr見積情報(9)) '購入理由
    '    strSql &= vbCrLf & "," & SQ(arr見積情報(11)) '科目ID
    '    strSql &= vbCrLf & "," & SQ(getNext購入ID())
    '    strSql &= vbCrLf & "," & SQ(Now)
    '    strSql &= vbCrLf & ")"
    '    'データベース接続、トランザクション開始（行ロックする）
    '    tran = dCon.Connection.BeginTransaction
    '    If dCon.ExecuteSqlMW(tran, strSql) = False Then
    '        tran.Rollback()
    '        MsgBox("TD_poへの書き込みに失敗しました")
    '        Exit Sub
    '    Else

    '    End If

    '    strSql = "delete from TD_見積"
    '    strSql &= vbCrLf & "where UID = " & arr見積情報(10)


    '    If dCon.ExecuteSqlMW(tran, strSql) = False Then
    '        tran.Rollback()
    '        MsgBox("TD_見積の削除に失敗しました")
    '        Exit Sub
    '    Else
    '        tran.Commit()
    '        tran.Dispose()
    '    End If


    '    DGV購入品入力_表示()
    'End Sub

    Private Sub chk未処理_CheckedChanged(sender As Object, e As EventArgs) Handles chk未処理.CheckedChanged

        ''矛盾チェック　未処理のみのチェックが入った状態で入荷日の絞り込みはできない
        'If chk未処理.Checked = True And cmb検索項目日付.SelectedValue = "arrival_date" Then
        '    '開始日と終了日のいずれかにチェックが入っていたら終了
        '    If dtp開始日.Checked = True Or dtp終了日.Checked = True Then
        '        MsgBox("「未処理のみ」にチェックを入れた状態で「入荷日」の絞り込みはできません")
        '        Exit Sub
        '    End If
        'End If

        If SwEditSave Then
            'If MsgBox("保存せずに画面を切り替えますか？", vbYesNo + vbDefaultButton1) = vbNo Then
            '    '元に戻す
            '    If chk未処理.Checked = True Then
            '        RemoveHandler Me.chk未処理.CheckedChanged, AddressOf chk未処理_CheckedChanged
            '        chk未処理.Checked = False
            '        AddHandler Me.chk未処理.CheckedChanged, AddressOf chk未処理_CheckedChanged
            '    Else
            '        RemoveHandler Me.chk未処理.CheckedChanged, AddressOf chk未処理_CheckedChanged
            '        chk未処理.Checked = True
            '        AddHandler Me.chk未処理.CheckedChanged, AddressOf chk未処理_CheckedChanged
            '    End If

            'Else
            '    SwEditSave = False
            '    DGV購入品入力_表示()
            'End If
            If MsgBox("保存してよろしいですか?" & vbCrLf & "「いいえ」を押すと更新していないデータは消えます", vbYesNo + vbDefaultButton1) = vbYes Then
                If メイン更新() = False Then
                    '更新失敗したら画面を書き換えずに終了
                    Exit Sub
                End If
            Else
                '本来は更新成功しないとSqEditSave = Trueにならないが、この場合は強制的にFalseにする(そうしないと更新するまで画面切り替える度、永遠に更新しますか？というメッセージが出てしまうため)
                SwEditSave = False

            End If

        End If

        DGV購入品入力_表示()
        DGV購入品入力_セル設定()
    End Sub

    Private Sub btn検索_Click(sender As Object, e As EventArgs) Handles btn検索.Click
        '検索のときは金額と予算額の計算がされないようにするため
        Disp検索結果()
    End Sub
    Sub Disp検索結果()
        Dim i As Integer
        Dim sValue列 As String
        Dim strSql As String
        Dim str検索型 As String

        検索条件 = ""
        日付検索条件 = ""
        検索Flg = False
        日付検索Flg = False

        For i = 1 To 3
            sValue列 = CType(Me.Controls("cmb検索項目" & i), ComboBox).SelectedValue
            If CType(Me.Controls("cmb検索項目" & i), ComboBox).SelectedIndex >= 0 _
                And CType(Me.Controls("cmb記号" & i), ComboBox).SelectedIndex >= 0 And CType(Me.Controls("txt検索条件" & i), TextBox).Text <> "" Then
                検索Flg = True
                strSql = ""
                strSql = "SELECT 型 FROM TM_列名変換 WHERE データ列 = '" & CType(Me.Controls("cmb検索項目" & i), ComboBox).SelectedValue & "'"
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
        '日付種類のコンボが空白でない時のみ日付を検索条件に加える

        If dtp開始日.Checked = True Then
            If cmb検索項目日付.SelectedValue Is Nothing Then
                MsgBox("日付で検索する場合は、日付種類を選択してください")
                Exit Sub
            Else
                日付検索Flg = True
                日付検索条件 &= vbCrLf & "AND " & cmb検索項目日付.SelectedValue & ">= '" & dtp開始日.Value.ToString("yyyy/MM/dd") & "'"

            End If
        End If
        If dtp終了日.Checked = True Then
            If cmb検索項目日付.SelectedValue Is Nothing Then
                MsgBox("日付で検索する場合は、日付種類を選択してください")
                Exit Sub
            Else

                日付検索Flg = True
                日付検索条件 &= vbCrLf & "AND " & cmb検索項目日付.SelectedValue & "<= '" & dtp終了日.Value.ToString("yyyy/MM/dd") & "'"
            End If
        End If


        If SwEditSave Then

            If MsgBox("保存してよろしいですか?" & vbCrLf & "「いいえ」を押すと更新していないデータは消えます", vbYesNo + vbDefaultButton1) = vbYes Then
                If メイン更新() = False Then
                    '更新失敗したら画面を書き換えずに終了
                    Exit Sub
                End If
            Else
                '本来は更新成功しないとSqEditSave = Trueにならないが、この場合は強制的にFalseにする(そうしないと更新するまで画面切り替える度、永遠に更新しますか？というメッセージが出てしまうため)
                SwEditSave = False
            End If

        End If
        DGV購入品入力_表示()
        DGV購入品入力_セル設定()
    End Sub
    'Function 文字種判定(検索項目名) As String
    '    Select Case 検索項目名
    '        Case "入荷日", "発注日", "申請日"
    '            文字種判定 = "日付型"
    '        Case "品名", "型式", "仕入先名"
    '            文字種判定 = "文字列型"
    '        Case "数量", "金額", "予算単価"
    '            文字種判定 = "数値型"

    '    End Select

    'End Function

    'Private Sub cmb検索項目1_TextChanged(sender As Object, e As EventArgs) Handles cmb検索項目1.TextChanged
    '    Dim strSql As String
    '    If cmb検索項目1.SelectedIndex >= 0 Then
    '        strSql = ""
    '        strSql = "SELECT 型 FROM TM_列名変換 WHERE データ列 = '" & cmb記号1.SelectedValue & "'"
    '        Dim DT As DataTable = dCon.DataSet(strSql, "DT").Tables(0)
    '        If DT.Rows.Count = 1 Then
    '            s型 = DT.Rows(0)("型")
    '            If s型 = "数値" Then
    '                txt検索条件1.ImeMode = Windows.Forms.ImeMode.Alpha
    '            End If

    '        Else
    '            MsgBox("TM_列名変換テーブルに型が登録されていません")
    '            cmb検索項目1.SelectedIndex = -1
    '        End If
    '    End If
    'End Sub




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
            strSql = "SELECT * FROM TM_列名変換 WHERE データ列 = '" & cmb項目.SelectedValue & "'"
            Dim DT As DataTable = dCon.DataSet(strSql, "DT").Tables(0)
            If DT.Rows.Count = 1 Then
                s型 = DT.Rows(0)("型")
                CType(Me.Controls("cmb記号" & i), ComboBox).Items.Clear()
                CType(Me.Controls("cmb記号" & i), ComboBox).Items.Add("=")
                CType(Me.Controls("cmb記号" & i), ComboBox).Items.Add("<>")

                Select Case s型
                    Case "数値"
                        CType(Me.Controls("cmb記号" & i), ComboBox).Items.Add(">=")
                        CType(Me.Controls("cmb記号" & i), ComboBox).Items.Add("<=")
                        CType(Me.Controls("cmb記号" & i), ComboBox).Items.Add(">")
                        CType(Me.Controls("cmb記号" & i), ComboBox).Items.Add("<")
                        CType(Me.Controls("txt検索条件" & i), TextBox).ImeMode = Windows.Forms.ImeMode.Disable

                    Case "文字"
                        CType(Me.Controls("cmb記号" & i), ComboBox).Items.Add("を含む")
                        CType(Me.Controls("txt検索条件" & i), TextBox).ImeMode = Windows.Forms.ImeMode.Hiragana

                        If DT.Rows(0)("表示列") = "支払区分" Then '支払区分だけはグリッドの表示名(pay_code)と表示列の名前が一致しないから特別処理が必要
                            dt選択肢 = makeSelectList("pay_code")
                            For Each r In dt選択肢.Rows
                                CType(Me.Controls("cmb検索条件選択" & i), ComboBox).Items.Add(r("pay_code"))
                            Next

                        Else

                            dt選択肢 = makeSelectList(DT.Rows(0)("表示列"))
                            For Each r In dt選択肢.Rows
                                CType(Me.Controls("cmb検索条件選択" & i), ComboBox).Items.Add(r(DT.Rows(0)("表示列")))
                            Next

                        End If

                End Select

                CType(Me.Controls("cmb記号" & i), ComboBox).SelectedIndex = -1
                CType(Me.Controls("txt検索条件" & i), TextBox).Enabled = True '位置0に行がありませんというエラーの対策
                CType(Me.Controls("txt検索条件" & i), TextBox).Text = ""
            Else
                MsgBox("TM_列名変換テーブルに型が登録されていません")
                cmb項目.SelectedIndex = -1
            End If
        End If

    End Sub




    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        For i = 1 To 3
            CType(Me.Controls("cmb検索項目" & i), ComboBox).SelectedIndex = -1
            CType(Me.Controls("cmb記号" & i), ComboBox).SelectedIndex = -1
            CType(Me.Controls("txt検索条件" & i), TextBox).Text = ""
            CType(Me.Controls("txt検索条件" & i), TextBox).Enabled = False
        Next

        rbn未承認.Checked = False
        rbn承認済.Checked = False
        chk未印刷.Checked = False
        chk未処理.Checked = False
        dtp開始日.Checked = False
        dtp開始日.Checked = False
        cmb検索項目日付.Text = ""
        日付検索条件 = ""
        検索Flg = False
        日付検索Flg = True
        日付検索条件 = "AND insertStamp >= " & Today.AddYears(-2) 'Form_Load時の日付条件にする
        If SwEditSave Then

            'If MsgBox("保存せずに画面を切り替えますか？", vbYesNo + vbDefaultButton1) = vbNo Then
            '    '何もしない
            'Else
            '    For i = 1 To 3
            '        CType(Me.Controls("cmb検索項目" & i), ComboBox).Text = ""
            '        CType(Me.Controls("cmb記号" & i), ComboBox).Text = ""
            '        CType(Me.Controls("txt検索条件" & i), TextBox).Text = ""
            '    Next
            '    chk未処理.Checked = False
            '    rbn未承認.Checked = False
            '    rbn承認済.Checked = False
            '    検索Flg = False
            '    SwEditSave = False
            '    DGV購入品入力_表示()

            'End If
            If メイン更新() = False Then
                '更新失敗したら画面を書き換えずに終了
                Exit Sub
            End If

        End If

        DGV購入品入力_表示()
        DGV購入品入力_セル設定()
    End Sub


    '「KeyPress」は、キーボード上の文字・数字とテンキーを押した時に1文字ずつ発生する。「tab」「shift」などを押されてもイベントは発生しません。
    'ただし「Enter」「BackSpace」「Esc」が押されたときはイベントが発生します。
    Private Sub txt検索条件_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt検索条件1.KeyPress, txt検索条件3.KeyPress, txt検索条件2.KeyPress
        Dim strSql As String
        Dim str検索型 As String
        Dim i As Integer
        i = Strings.Right(CType(sender, TextBox).Name, 1)
        strSql = ""
        strSql = "SELECT 型 FROM TM_列名変換 WHERE データ列 = '" & CType(Me.Controls("cmb検索項目" & i), ComboBox).SelectedValue & "'"
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
    Private Function makeSelectList(ColName As String) As DataTable
        Dim dt検索用リスト As DataTable
        Dim dtView As DataView
        dtView = New DataView(DGV購入品入力.DataSource)
        dtView.Sort = ColName & " ASC"
        dt検索用リスト = dtView.ToTable(True, ColName)
        Return dt検索用リスト
    End Function
    Private Sub 自動計算(TargetRow As Integer, flg予算 As Boolean)

        'Dim i単価 As Integer
        'Dim i数量 As Integer
        'int(32,767)より最大値が大きいLong(2,147,483,647)を使うことにした
        '単価が小数のものがあると分かったのでDecimal型を使うことにした　±79,228,162,514,264,337,593,543,950,335まで表せる
        'Dim l単価 As Long
        Dim d単価 As Decimal
        Dim l数量 As Long
        'ありえなさそうな数字が入っていたら警告する
        If DGV購入品入力.Rows(TargetRow).Cells("数量").Value Is DBNull.Value Or DGV購入品入力.Rows(TargetRow).Cells("数量").Value Is Nothing Then
            l数量 = 0
        Else
            l数量 = CLng(DGV購入品入力.Rows(TargetRow).Cells("数量").Value)
        End If
        '数量が0以上の時だけ計算する
        If l数量 > 0 Then
            If flg予算 Then
                If DGV購入品入力.Rows(TargetRow).Cells("予算単価").Value Is DBNull.Value Or DGV購入品入力.Rows(TargetRow).Cells("予算単価").Value Is Nothing Then
                    d単価 = 0

                Else
                    d単価 = CDbl(DGV購入品入力.Rows(TargetRow).Cells("予算単価").Value)
                End If
                If d単価 > 1000000000 Or l数量 > 1000000000 Then
                    '数量か単価が10億超えていたら確認メッセージを出す
                    MsgBox("大きすぎる数字が入力されています。確認してください。")
                Else
                    If d単価 > 0 Then
                        RemoveHandler DGV購入品入力.CellValueChanged, AddressOf DGV購入品入力_CellValueChanged
                        DGV購入品入力.Rows(TargetRow).Cells("予算額").Value = l数量 * d単価
                        AddHandler DGV購入品入力.CellValueChanged, AddressOf DGV購入品入力_CellValueChanged
                    Else
                        '単価が0以下なら予算額は空白
                        DGV購入品入力.Rows(TargetRow).Cells("予算額").Value = DBNull.Value
                    End If

                End If
            Else
                If DGV購入品入力.Rows(TargetRow).Cells("税抜単価").Value Is DBNull.Value Or DGV購入品入力.Rows(TargetRow).Cells("税抜単価").Value Is Nothing Then
                    d単価 = 0
                Else
                    d単価 = CDbl(DGV購入品入力.Rows(TargetRow).Cells("税抜単価").Value)
                End If

                If d単価 > 1000000000 Or l数量 > 1000000000 Then
                    MsgBox("大きすぎる数字が入力されています。確認してください。")
                Else
                    If d単価 > 0 Then
                        RemoveHandler DGV購入品入力.CellValueChanged, AddressOf DGV購入品入力_CellValueChanged
                        DGV購入品入力.Rows(TargetRow).Cells("金額").Value = l数量 * d単価
                        AddHandler DGV購入品入力.CellValueChanged, AddressOf DGV購入品入力_CellValueChanged
                    Else
                        '単価が0より小さかったら金額は空白
                        DGV購入品入力.Rows(TargetRow).Cells("金額").Value = DBNull.Value
                    End If

                End If
            End If
        Else
            '数量が0以下だったらどちらも空白にする
            RemoveHandler DGV購入品入力.CellValueChanged, AddressOf DGV購入品入力_CellValueChanged
            DGV購入品入力.Rows(TargetRow).Cells("金額").Value = DBNull.Value
            DGV購入品入力.Rows(TargetRow).Cells("予算額").Value = DBNull.Value
            AddHandler DGV購入品入力.CellValueChanged, AddressOf DGV購入品入力_CellValueChanged

        End If
    End Sub

    '列の幅が決まっていたほうがいいと言われたので、購入IDはあえて0埋めのstringにしている。計算するときは数値(Intは桁数足りないから×)にしないといけないから要注意!
    Public Function getNext購入ID() As String
        'DataSetの中の処理でエラーが出てしまっていたのでコネクションを別で宣言することにした(本当はあまりよくないがITで話し合いこれしかないという結論になった)
        Dim mCon As New merrweth_init_DbConnection
        '番号管理から番号取得
        Dim strSql As String = ""
        Dim ID As String
        Dim Number As String

        Dim tran As System.Data.SqlClient.SqlTransaction = Nothing
        strSql = "select * from TM_ミヤマ番号管理"
        strSql &= vbCrLf & "with( UPDLOCK, ROWLOCK ) "
        strSql &= vbCrLf & "where 区分 = '購入ID'"
        Dim dataSet = mCon.DataSet(strSql, "dt購入ID")
        ID = kc.nz(dataSet.Tables(0).Rows(0)("番号"))
        Number = Long.Parse(ID) + 1 '数字に変換して1を足す
        Dim Zero As New String("0"c, 8 - Number.Length)
        getNext購入ID = Zero & Number

        strSql = "update TM_ミヤマ番号管理 set 番号 = 番号 + 1 where 区分 = '購入ID'"

        'データベース接続、トランザクション開始（行ロックする）
        tran = mCon.Connection.BeginTransaction


        If mCon.ExecuteSqlMW(tran, strSql) = False Then
            tran.Rollback()

        Else
            tran.Commit()

        End If
        tran.Dispose()
    End Function

    Private Sub INPUT削除履歴(intpoID As Integer)
        Dim mCon As New merrweth_init_DbConnection
        Dim tran As System.Data.SqlClient.SqlTransaction = Nothing
        Dim strSql As String = ""
        Dim flgINPUT成功 As Boolean

        strSql = ""
        strSql &= vbCrLf & "INSERT INTO"
        strSql &= vbCrLf & "TD_購入依頼削除履歴"
        strSql &= vbCrLf & "("
        strSql &= vbCrLf & "po_id"
        strSql &= vbCrLf & ",order_date"
        strSql &= vbCrLf & ",arrival_date"
        strSql &= vbCrLf & ",division_id"
        strSql &= vbCrLf & ",group_name"
        strSql &= vbCrLf & ",employee"
        strSql &= vbCrLf & ",part_name"
        strSql &= vbCrLf & ",part_number"
        strSql &= vbCrLf & ",unit_price"
        strSql &= vbCrLf & ",number"
        strSql &= vbCrLf & ",kamoku_id"
        strSql &= vbCrLf & ",vendor_id"
        strSql &= vbCrLf & ",vendor"
        strSql &= vbCrLf & ",pay_code"
        strSql &= vbCrLf & ",remark"
        strSql &= vbCrLf & ",work_code"
        strSql &= vbCrLf & ",tana_bango"
        strSql &= vbCrLf & ",no_tax"
        strSql &= vbCrLf & ",承認済"
        strSql &= vbCrLf & ",決裁"
        strSql &= vbCrLf & ",登録番号"
        strSql &= vbCrLf & ",予算単価"
        strSql &= vbCrLf & ",整理番号"
        strSql &= vbCrLf & ",経理検査"
        strSql &= vbCrLf & ",改訂番号"
        strSql &= vbCrLf & ",入荷ID"
        strSql &= vbCrLf & ",備考1"
        strSql &= vbCrLf & ",購入ID"
        strSql &= vbCrLf & ",insertStamp"
        strSql &= vbCrLf & ",updateStamp"
        strSql &= vbCrLf & ",印刷日"
        strSql &= vbCrLf & ",CSVNo"
        strSql &= vbCrLf & ",作成者"
        strSql &= vbCrLf & ",更新者"
        strSql &= vbCrLf & ",作成PC名"
        strSql &= vbCrLf & ",更新PC名"
        strSql &= vbCrLf & ",作成バージョン"
        strSql &= vbCrLf & ",更新バージョン"
        strSql &= vbCrLf & ")"
        strSql &= vbCrLf & "SELECT"
        strSql &= vbCrLf & "id"
        strSql &= vbCrLf & ",order_date"
        strSql &= vbCrLf & ",arrival_date"
        strSql &= vbCrLf & ",division_id"
        strSql &= vbCrLf & ",group_name"
        strSql &= vbCrLf & ",employee"
        strSql &= vbCrLf & ",part_name"
        strSql &= vbCrLf & ",part_number"
        strSql &= vbCrLf & ",unit_price"
        strSql &= vbCrLf & ",number"
        strSql &= vbCrLf & ",kamoku_id"
        strSql &= vbCrLf & ",vendor_id"
        strSql &= vbCrLf & ",vendor"
        strSql &= vbCrLf & ",pay_code"
        strSql &= vbCrLf & ",remark"
        strSql &= vbCrLf & ",work_code"
        strSql &= vbCrLf & ",tana_bango"
        strSql &= vbCrLf & ",no_tax"
        strSql &= vbCrLf & ",承認済"
        strSql &= vbCrLf & ",決裁"
        strSql &= vbCrLf & ",登録番号"
        strSql &= vbCrLf & ",予算単価"
        strSql &= vbCrLf & ",整理番号"
        strSql &= vbCrLf & ",経理検査"
        strSql &= vbCrLf & ",改訂番号"
        strSql &= vbCrLf & ",入荷ID"
        strSql &= vbCrLf & ",備考1"
        strSql &= vbCrLf & ",購入ID"
        strSql &= vbCrLf & ",insertStamp"
        strSql &= vbCrLf & ",updateStamp"
        strSql &= vbCrLf & ",印刷日"
        strSql &= vbCrLf & ",CSVNo"
        strSql &= vbCrLf & ",作成者"
        strSql &= vbCrLf & ",更新者"
        strSql &= vbCrLf & ",作成PC名"
        strSql &= vbCrLf & ",更新PC名"
        strSql &= vbCrLf & ",作成バージョン"
        strSql &= vbCrLf & ",更新バージョン"
        strSql &= vbCrLf & "FROM TD_po"
        strSql &= vbCrLf & "WHERE id =" & intpoID

        tran = mCon.Connection.BeginTransaction

        If mCon.ExecuteSqlMW(tran, strSql) = False Then
            tran.Rollback()
            flgINPUT成功 = False
            MsgBox("削除履歴のINSERTに失敗しました")
        Else
            tran.Commit()
            flgINPUT成功 = True
        End If
        tran.Dispose()

        If flgINPUT成功 Then
            strSql = ""
            strSql &= vbCrLf & "UPDATE"
            strSql &= vbCrLf & "TD_購入依頼削除履歴"
            strSql &= vbCrLf & "SET"
            strSql &= vbCrLf & "削除者 = " & kc.SQ(Select_Form._UserName)
            strSql &= vbCrLf & ",削除PC名 = " & kc.SQ(Select_Form._PC名)
            strSql &= vbCrLf & ",削除バージョン = " & kc.SQ(cVer)
            strSql &= vbCrLf & ",DeleteStamp = " & kc.SQ(Now())
            strSql &= vbCrLf & "WHERE po_id = " & intpoID

            tran = mCon.Connection.BeginTransaction
            If mCon.ExecuteSqlMW(tran, strSql) = False Then
                tran.Rollback()
                flgINPUT成功 = False
                MsgBox("削除履歴のUPDATEに失敗しました")
            Else
                tran.Commit()
                flgINPUT成功 = True
            End If
            tran.Dispose()

        End If
    End Sub
    Private Sub btn送料_Click(sender As Object, e As EventArgs) Handles btn送料.Click, btn手数料.Click, btn値引き.Click
        Dim iコピー行No As Integer = DGV購入品入力.SelectedCells(0).RowIndex
        Dim row As DataGridViewRow
        Dim i科目ID As Integer
        Dim strSql As String
        Dim tran As System.Data.SqlClient.SqlTransaction = Nothing
        Dim button = CType(sender, Button)
        '2022/02/15 ekawai add S 
        If DGV購入品入力.GetCellCount(DataGridViewElementStates.Selected) > 1 Then
            MsgBox("複数セルを選択している場合は実行できません")
            Exit Sub
        End If
        '2022/02/15 ekawai add E

        If kc.ns(DGV購入品入力.Rows(iコピー行No).Cells("品名").Value) = "" Then
            MsgBox("送料を入れたい商品の行を選択してください。" & vbCrLf & "送料はその下に追加されます")
            Exit Sub
        End If

        If SwEditSave Then
            If MsgBox("保存してよろしいですか？", vbYesNo + vbDefaultButton1) = vbYes Then
                If メイン更新() = False Then
                    '更新失敗したら画面を書き換えずに終了
                    Exit Sub
                End If
            Else
                '本来は更新成功しないとSqEditSave = Trueにならないが、この場合は強制的にFalseにする(そうしないと更新するまで画面切り替える度、永遠に更新しますか？というメッセージが出てしまうため)
                SwEditSave = False
            End If
        End If


        row = DGV購入品入力.Rows(iコピー行No)
        '2022/11/02 ekawai del S ReadOnly使わないように指示があったため削除
        '理由は謎だが旧購入依頼でロック解除していたのでそれを継承
        'If row.Cells("型式").ReadOnly = True Then
        '    row.Cells("型式").ReadOnly = False
        'End If
        '2022/11/02 ekawai del E 
        '部門によって科目を変更する(旧購入依頼の仕様)
        'If Cells(intRow + 1, INPUT_DIVISION).Value = "一般管理販売部門" _
        'Or Cells(intRow + 1, INPUT_DIVISION).Value = "営業部門" _
        'Or Cells(intRow + 1, INPUT_DIVISION).Value = "情報システム部門" _
        'Or Cells(intRow + 1, INPUT_DIVISION).Value = "役員部門" Then
        '    Cells(intRow + 1, INPUT_KAMOKU_ID).Value = 690
        'Else
        '    Cells(intRow + 1, INPUT_KAMOKU_ID).Value = 545
        'End If

        Dim val As String = InputBox("金額を入力してください")
        Dim d金額 As Decimal '行コピーの時に単価が小数の場合もあり得るからDecimalにしている

        If INPUTBOX_入力値チェック(val) = True Then

            d金額 = val


            If row.Cells("部門ID").Value = "1" Then '1=営業部門 division_idは文字列
                i科目ID = 690
            Else
                i科目ID = 545
            End If

            Select Case button.Text

                Case "送料"
                    strSql = Input_コピー(iコピー行No, 1, "送料", d金額, d金額, "掛け", i科目ID, "", row.Cells("購入ID").Value)
                Case "手数料"
                    strSql = Input_コピー(iコピー行No, 1, "手数料", d金額, d金額, "掛け", i科目ID, "", row.Cells("購入ID").Value)
                Case "値引き"

                    If row.Cells("予算単価").Value - d金額 <= 0 Then
                        MsgBox("元の予算単価より大きい数は値引きできません")
                        Exit Sub
                    Else
                        strSql = Input_コピー(iコピー行No, 1, "値引き", -d金額, -d金額, "掛け", row.Cells("科目ID").Value, "", row.Cells("購入ID").Value)
                    End If
                Case Else
                    strSql = ""
            End Select

            tran = dCon.Connection.BeginTransaction
            If dCon.ExecuteSqlMW(tran, strSql) = False Then
                tran.Rollback()
                MsgBox("更新失敗")
                SwEditSave = True
            Else
                tran.Commit()
                tran.Dispose()
                SwEditSave = False
                DGV購入品入力_表示()
                DGV購入品入力_セル設定()
            End If

            'dCon.ConnectionClose() グリッド表示するところで ConnectionString プロパティは初期化されていませんと出るのでコメントアウト


        End If
    End Sub

    Private Sub btnCOPY_Click(sender As Object, e As EventArgs) Handles btnCOPY.Click
        Dim iコピー行No As Integer = DGV購入品入力.SelectedCells(0).RowIndex
        Dim strSql As String
        Dim s複写元購入ID As String
        Dim row As DataGridViewRow
        'バインド元に行を追加する
        Dim dt As DataTable = DGV購入品入力.DataSource
        Dim tran As System.Data.SqlClient.SqlTransaction = Nothing

        'If ns(DGV購入品入力.CurrentRow.Cells("品名").Value) = "" Then
        '    MsgBox("コピーしたい商品の行を選択してください。")
        '    Exit Sub
        'End If
        '2022/02/15 ekawai add S
        If DGV購入品入力.GetCellCount(DataGridViewElementStates.Selected) > 1 Then
            MsgBox("複数セルを選択している場合はコピーできません")
            Exit Sub
        End If
        '2022/02/15 ekawai add E
        If SwEditSave Then
            If MsgBox("保存してよろしいですか？", vbYesNo + vbDefaultButton1) = vbYes Then
                If メイン更新() = False Then
                    '更新失敗したら画面を書き換えずに終了
                    MsgBox("コピーを中止します")
                    Exit Sub
                Else

                End If
            Else
                '本来は更新成功しないとSqEditSave = Trueにならないが、この場合は強制的にFalseにする(そうしないと更新するまで画面切り替える度、永遠に更新しますか？というメッセージが出てしまうため)
                SwEditSave = False
            End If
        End If


        '旧購入依頼で承認済みの時型式がロックされているが、コピー時は解除していたのでそれを継承している。数量もロック解除すればよいそう。
        'If row.Cells("型式").ReadOnly = True Then
        '    row.Cells("型式").ReadOnly = False
        'End If
        'If row.Cells("数量").ReadOnly = True Then
        '    row.Cells("数量").ReadOnly = False
        'End If


        'Dim num As String = InputBox("コピー納入数を入力してください")
        'Dim i分納数 As Integer


        'If INPUTBOX_入力値チェック(num) = True Then
        '    If row.Cells("数量").Value - num <= 0 Then
        '        MsgBox("元のデータより大きい数にはコピーできません")
        '        Exit Sub
        '    Else
        '        i分納数 = num
        '        '////コピー元////
        '        strSql = "UPDATE TD_po"
        '        strSql &= vbCrLf & "SET"
        '        strSql &= vbCrLf & "number = " & row.Cells("数量").Value - i分納数
        '        strSql &= vbCrLf & "WHERE id = " & row.Cells("id").Value
        '        'コピー元のDBの数量をアップデート
        '        If dCon.ExecuteSqlMW(tran, strSql) = False Then
        '            tran.Rollback()
        '            tran.Dispose()
        '            MsgBox("更新失敗")
        '            SwEditSave = True
        '            Exit Sub
        '        Else

        'コピー元の更新が成功したら新しい行をINSERTする
        '////コピー先////
        'strSql = Input_コピー(iコピー行No, i分納数, row.Cells("品名").Value, row.Cells("予算単価").Value, nz(row.Cells("税抜単価").Value), row.Cells("pay_code").Value, row.Cells("科目ID").Value, ns(row.Cells("型式").Value))
        If DGV購入品入力.RowCount = iコピー行No + 1 Then
            MsgBox("この行はコピーできません")
            Exit Sub
        Else
            row = DGV購入品入力.Rows(iコピー行No)
        End If

        s複写元購入ID = row.Cells("購入ID").Value

        strSql = Input_コピー(iコピー行No, row.Cells("数量").Value, row.Cells("品名").Value, row.Cells("予算単価").Value, kc.nz(row.Cells("税抜単価").Value), row.Cells("pay_code").Value, row.Cells("科目ID").Value, kc.ns(row.Cells("型式").Value), s複写元購入ID)
        tran = dCon.Connection.BeginTransaction
        If dCon.ExecuteSqlMW(tran, strSql) = False Then
            tran.Rollback()
            tran.Dispose()
            MsgBox("更新失敗")
            SwEditSave = True
            Exit Sub
        Else
            tran.Commit()
            tran.Dispose()
            SwEditSave = False
            DGV購入品入力_表示()
            'DGV購入品入力_セル設定(s複写元購入ID)
            DGV購入品入力_セル設定()

            tran = dCon.Connection.BeginTransaction
            strSql = "UPDATE TD_po SET コピーフラグ = 'True' WHERE 購入ID = " & kc.SQ(s複写元購入ID)
            If dCon.ExecuteSqlMW(tran, strSql) = False Then
                tran.Rollback()
                tran.Dispose()
                MsgBox("更新失敗")
                SwEditSave = True
                Exit Sub

            Else
                tran.Commit()
                tran.Dispose()
                SwEditSave = False
                DGV購入品入力_表示()
                DGV購入品入力_セル設定()

            End If

        End If


        'End If


        'End If
        'End If

    End Sub
    '送料や値引きボタンと共用
    Function Input_コピー(iRow As Integer, iSuryo As Integer, sHinName As String, dYTanka As Decimal, dZTanka As Decimal, sPayCode As String, iKamokuID As Integer, sKata As String, sID As String)
        '決裁済みの行もコピーできてしまうので大丈夫か念のため確認したが、旧購入依頼からそういう仕様だから問題ないとのこと。分納の時にコピーできないと困るから。
        Dim strSql As String
        Dim row As DataGridViewRow
        row = DGV購入品入力.Rows(iRow)
        strSql = "INSERT INTO"
        strSql &= vbCrLf & "TD_po("
        strSql &= vbCrLf & "group_name"
        strSql &= vbCrLf & ",part_name"
        strSql &= vbCrLf & ",order_date"
        strSql &= vbCrLf & ",arrival_date"
        strSql &= vbCrLf & ",印刷日"
        strSql &= vbCrLf & ",division_id"
        strSql &= vbCrLf & ",employee"
        strSql &= vbCrLf & ",part_number"
        strSql &= vbCrLf & ",unit_price"
        strSql &= vbCrLf & ",number"
        strSql &= vbCrLf & ",kamoku_id"
        strSql &= vbCrLf & ",vendor_id"
        strSql &= vbCrLf & ",vendor"
        strSql &= vbCrLf & ",pay_code"
        strSql &= vbCrLf & ",remark"
        strSql &= vbCrLf & ",work_code"
        strSql &= vbCrLf & ",tana_bango"
        strSql &= vbCrLf & ",承認済"
        strSql &= vbCrLf & ",決裁"
        strSql &= vbCrLf & ",登録番号"
        strSql &= vbCrLf & ",予算単価"
        strSql &= vbCrLf & ",整理番号"
        strSql &= vbCrLf & ",改訂番号"
        strSql &= vbCrLf & ",備考1"
        strSql &= vbCrLf & ",購入ID"
        strSql &= vbCrLf & ",insertStamp"
        strSql &= vbCrLf & ",作成者"
        strSql &= vbCrLf & ",作成PC名"
        strSql &= vbCrLf & ",作成バージョン"

        strSql &= vbCrLf & ") VALUES ("
        strSql &= vbCrLf & kc.SQ(Select_Form.s部署)
        strSql &= vbCrLf & "," & kc.SQ(sHinName)
        strSql &= vbCrLf & "," & kc.nn(row.Cells("発注日").Value)
        strSql &= vbCrLf & "," & kc.nn(row.Cells("入荷日").Value)
        strSql &= vbCrLf & "," & kc.nn(row.Cells("印刷日").Value)
        strSql &= vbCrLf & "," & kc.SQ(row.Cells("部門ID").Value)
        strSql &= vbCrLf & "," & kc.SQ(row.Cells("購入者").Value)
        strSql &= vbCrLf & "," & kc.SQ(sKata)
        '税抜単価が空白だったらnullで更新
        If dZTanka = 0 Then
            strSql &= vbCrLf & ",null"
        Else
            strSql &= vbCrLf & "," & dZTanka
        End If
        strSql &= vbCrLf & "," & iSuryo
        strSql &= vbCrLf & "," & iKamokuID
        strSql &= vbCrLf & "," & kc.nn(row.Cells("仕入先ID").Value)
        strSql &= vbCrLf & "," & kc.SQ(row.Cells("仕入先").Value)
        strSql &= vbCrLf & "," & kc.SQ(sPayCode)
        strSql &= vbCrLf & "," & kc.nn(row.Cells("購入理由").Value)
        strSql &= vbCrLf & "," & kc.nn(row.Cells("ワークコード").Value)
        strSql &= vbCrLf & "," & kc.nn(row.Cells("棚番号").Value)
        strSql &= vbCrLf & "," & kc.nn(row.Cells("承認済").Value)
        strSql &= vbCrLf & "," & kc.nn(row.Cells("決裁").Value)
        strSql &= vbCrLf & "," & kc.nn(row.Cells("登録番号").Value)
        strSql &= vbCrLf & "," & dYTanka
        strSql &= vbCrLf & "," & kc.nn(row.Cells("整理番号").Value)
        strSql &= vbCrLf & "," & kc.nn(row.Cells("改訂番号").Value)
        strSql &= vbCrLf & "," & kc.nn(row.Cells("備考1").Value)
        strSql &= vbCrLf & "," & kc.SQ(sID)
        strSql &= vbCrLf & "," & kc.SQ(Now)
        strSql &= vbCrLf & "," & kc.SQ(Select_Form._UserName)
        strSql &= vbCrLf & "," & kc.SQ(Select_Form._PC名)
        strSql &= vbCrLf & "," & kc.SQ(cVer)


        strSql &= vbCrLf & ")"

        Return strSql

    End Function
    Private Function INPUTBOX_入力値チェック(s As String)
        If s = "" Then
            'VB.NETのINPUTBOXは、キャンセル、右上の閉じるボタン、空白の状態で「OK」どれも""が返ってくる仕様。
            'そのためボタンごとにメッセージを分けれないことを説明したら、最終的にキャンセルできれば別に細かいこと気にしないでいいとのことだったので
            'どのボタンを押しても「空白です」というメッセージを出している。
            MsgBox("空白です")
            Return False
        End If

        If Not IsNumeric(s) Then
            MsgBox("数字を入力してください")
            Return False
        End If

        If s <= 0 Then
            MsgBox("0以下の数字は入力できません")
            Return False
        Else
            Return True
        End If

    End Function

    'Deleteキーでの削除はユーザーに馴染みがないからやめたほうがいいから、既存の購入依頼のようにボタンからしか削除できないようにしている
    '複数選択可能にすると行コピーの時にどこを選択しているのか判定できなくなるので、複数行削除は不可としている
    '削除は1行なのに対して更新はDGV全体に対して行うから辻褄が合わなくなってしまった
    'そのためDeleteボタンは本当の削除ではなく見た目上の削除にすることに決まった　2021/10/13
    '削除ボタン押すと、実際の選択行と見た目の一番下の行が一緒にDELETEされていることが分かった。。
    'もはやなぜ2021/10/13に仮削除を採用したか忘れたし、デメリットないと思うから、即削除すればよいということになった。
    Private Sub btn削除_Click(sender As Object, e As EventArgs) Handles btn削除.Click
        Dim i削除行No As Integer = DGV購入品入力.SelectedCells(0).RowIndex
        Dim strSql As String = ""
        Dim tran As System.Data.SqlClient.SqlTransaction = Nothing
        Dim Response As String
        Dim strmsg As String = ""

        'Dim DT同時更新 As New DataTable

        ''空のテーブルを作る(列を先に入れておく)　…削除の場合はExcel出力いらないそう。1行だからDGV上で分かる。
        'For Each col As DataGridViewColumn In DGV購入品入力.Columns
        '    DT同時更新.Columns.Add(col.Name)
        'Next

        'If SwEditSave = True Then
        '    If MsgBox("保存してよろしいですか?" & vbCrLf & "「いいえ」を押すと更新していないデータは消えます", vbYesNo + vbDefaultButton1) = vbYes Then
        '        If Main更新(i削除行No) = False Then
        '            Exit Sub
        '        End If
        '    Else
        '        '本来は更新成功しないとSqEditSave = Trueにならないが、この場合は強制的にFalseにする(そうしないと更新するまで画面切り替える度、永遠に更新しますか？というメッセージが出てしまうため)
        '        SwEditSave = False
        '    End If
        'End If

        '最終行をRemoveAtで消そうとすると「コミットされていない新しい行を削除することはできません。」というエラーが出てしまった。(Deleteキーの場合だと最終行削除できないよう制御されている)
        'Datatableの行数とDataGridviewの行数で最終行かどうか判定することにした 
        '※一度変数に入れてdt.rows.countとした時とDGV○○.rows.countとした時で結果変わるので注意。
        '　検証したところdatatabele型の変数に入れた場合は空白行なしの行数がrows.countの結果、DGV○○とした場合は空白行も含む行数になった。つまりDGV○○とした時のほうが1行多い。

        '選択されているセルの数だけを取得したいのであれば、「DataGridView1.SelectedCells.Count」のように取得するのではなく、DataGridView.GetCellCountメソッドを使用したほうが効率的らしい。
        '2022/02/14 ekawai add S
        If DGV購入品入力.GetCellCount(DataGridViewElementStates.Selected) > 1 Then
            MsgBox("複数セルを選択している場合は削除できません")
            Exit Sub
        End If
        '2022/02/14 ekawai add E
        If i削除行No + 1 = DGV購入品入力.Rows.Count Then
            'DataGridViewの一番下の新規行(自動で作られる行)は削除できない
            MsgBox("新規行は削除できません")
            Exit Sub

        End If

        If kc.nz(DGV購入品入力.Rows(i削除行No).Cells("経理検査").Value) Then
            MsgBox("経理検査済の行は削除できません")
            Exit Sub
        End If

        '削除行を選択して分かりやすくする
        DGV購入品入力.Rows(i削除行No).Selected = True
        '2022/03/03 ekawai add S
        strmsg = "選択中の行を削除してよろしいですか？" & _
    vbCrLf & "※「はい」を押すと、即データベースから消えます。" & vbCrLf & _
    vbCrLf & "購入ID: " & (DGV購入品入力.Rows(i削除行No).Cells("購入ID").Value) & _
    vbCrLf & "購入者: " & (DGV購入品入力.Rows(i削除行No).Cells("購入者").Value) & _
    vbCrLf & "品名: " & (DGV購入品入力.Rows(i削除行No).Cells("品名").Value) & _
    vbCrLf & "型式: " & (DGV購入品入力.Rows(i削除行No).Cells("型式").Value) & _
    vbCrLf & "数量: " & (DGV購入品入力.Rows(i削除行No).Cells("数量").Value) & _
    vbCrLf & "予算単価: " & (DGV購入品入力.Rows(i削除行No).Cells("予算単価").Value) & _
    vbCrLf & "仕入先: " & (DGV購入品入力.Rows(i削除行No).Cells("仕入先").Value) & _
    vbCrLf & "支払区分: " & (DGV購入品入力.Rows(i削除行No).Cells("pay_code").Value) & _
    vbCrLf & "購入理由: " & (DGV購入品入力.Rows(i削除行No).Cells("購入理由").Value)

        If kc.nz(DGV購入品入力.Rows(i削除行No).Cells("id").Value) = 0 Then
            '新規行の場合
            'まだデータベースに登録されていないから見た目上行を消したら終わり
            Response = MsgBox(strmsg, MsgBoxStyle.YesNo, "確認")
            If Response = vbYes Then
                DGV購入品入力.Rows.RemoveAt(i削除行No)
            End If
        Else
            '2022/03/03 ekawai add E
            '既存行の場合
            If CheckID(kc.nz(DGV購入品入力.Rows(i削除行No).Cells("id").Value)) <> "更新対象" Then
                '元々はメイン更新で削除していたため複数同時更新対象があるからExcel出力していたが、btn削除では1行しか削除できないのでわざわざExcel出力しないでいいそう
                'Dim DR As DataRow
                'DR = DT同時更新.NewRow
                ''重複更新行の情報をテーブルに格納
                'For Each col As DataColumn In DGV購入品入力.Columns
                '    DR(col.ColumnName) = DGV購入品入力.Rows(i削除行No).Cells(col.ColumnName).Value
                'Next
                'DT同時更新.Rows.Add(DR)
                MsgBox("他のユーザーによって変更されたため、削除できません")
                'Excel出力(DT同時更新)　
            Else
                '2022/03/03 ekawai del
                '    strmsg = "選択中の行を削除してよろしいですか？" & _
                'vbCrLf & "※「はい」を押すと、即データベースから消えます。" & vbCrLf & _
                'vbCrLf & "購入ID: " & (DGV購入品入力.Rows(i削除行No).Cells("購入ID").Value) & _
                'vbCrLf & "購入者: " & (DGV購入品入力.Rows(i削除行No).Cells("購入者").Value) & _
                'vbCrLf & "品名: " & (DGV購入品入力.Rows(i削除行No).Cells("品名").Value) & _
                'vbCrLf & "型式: " & (DGV購入品入力.Rows(i削除行No).Cells("型式").Value) & _
                'vbCrLf & "数量: " & (DGV購入品入力.Rows(i削除行No).Cells("数量").Value) & _
                'vbCrLf & "予算単価: " & (DGV購入品入力.Rows(i削除行No).Cells("予算単価").Value) & _
                'vbCrLf & "仕入先: " & (DGV購入品入力.Rows(i削除行No).Cells("仕入先").Value) & _
                'vbCrLf & "支払区分: " & (DGV購入品入力.Rows(i削除行No).Cells("pay_code").Value) & _
                'vbCrLf & "購入理由: " & (DGV購入品入力.Rows(i削除行No).Cells("購入理由").Value)

                'If kc.nz(DGV購入品入力.Rows(i削除行No).Cells("id").Value) = 0 Then
                '    'まだデータベースに登録されていないから見た目上行を消したら終わり
                '    Response = MsgBox(strmsg, MsgBoxStyle.YesNo, "確認")
                '    If Response = vbYes Then
                '        DGV購入品入力.Rows.RemoveAt(i削除行No)
                '    End If
                'Else
                '2022/03/03 ekawai del 新規入力業の削除はCheckIDの前に実施しないといけなかった
                'メッセージの内容を詳しくして間違い削除防止
                Response = MsgBox(strmsg, MsgBoxStyle.YesNo, "確認")

                If Response = vbYes Then

                    'If nz(DGV購入品入力.Rows(i削除行No).Cells("id").Value) = 0 Then

                    'DGV購入品入力.Rows.RemoveAt(i削除行No)　2022/02/08 ekawai コメントアウト

                    'SwEditSave = True
                    'Else
                    '本当は成功した場合だけ履歴残したいけど、削除後だとSELECTできないから、配列入れる必要があるが面倒なので、
                    '削除成功しても失敗しても履歴テーブルには書き込めばよいとのこと　2022/02/07 
                    INPUT削除履歴(DGV購入品入力.Rows(i削除行No).Cells("id").Value)

                    '2022/02/08 ekawai コメントアウトを復活
                    strSql = "DELETE FROM TD_po"
                    strSql &= vbCrLf & "WHERE id = " & DGV購入品入力.Rows(i削除行No).Cells("id").Value
                    tran = dCon.Connection.BeginTransaction
                    If dCon.ExecuteSqlMW(tran, strSql) = False Then
                        tran.Rollback()
                        tran.Dispose()
                        'SwEditSave = True
                        MsgBox("削除失敗")
                        Exit Sub

                    End If

                    tran.Commit()
                    tran.Dispose()

                    DGV購入品入力_表示()
                    DGV購入品入力_セル設定()
                    '2022/02/08 ekawai
                    MsgBox("削除成功")

                    'End If

                End If

            End If


        End If



        SwEditSave = False
    End Sub



    '未印刷チェックがオンの時のみ押せるようにする
    '未印刷かどうかは印刷日に日付が入っているかどうかで判定。申請日という名前はやめて印刷日にする。
    '購入品精算書も同じボタンで処理する(未承認印刷の時は出さない)
    Private Sub 印刷準備(b未承認印刷 As Boolean)

        Dim xlsTemplate As String = ""
        xlsTemplate = "\\FS1\060System\Excelテンプレート\購入依頼_注文書・精算書.xlsx"
        Dim oApp As Excel.Application
        Dim oBooks As Excel.Workbooks
        Dim oBook As Excel.Workbook
        Dim oSheets As Excel.Sheets
        Dim oSheet1 As Excel.Worksheet
        Dim oSheet2 As Excel.Worksheet
        Dim oDialogs As Excel.Dialogs

        'Applicationインスタンスの生成
        oApp = New Excel.Application

        Dim Arr(,) As Object = Nothing

        'oAppが持つWorkbooksプロパティを取得
        oBooks = oApp.Workbooks
        'oBooksに、ExcelFileNameで指定したファイルをオープンし、そのインスタンスをoBookに格納します。
        oBook = oBooks.Open(xlsTemplate)
        'ワークシートを選択
        oSheets = oBook.Worksheets
        oSheet1 = DirectCast(oSheets("注文書"), Excel.Worksheet)
        oSheet2 = DirectCast(oSheets("購入品精算書"), Excel.Worksheet)
        oDialogs = oApp.Dialogs



        'セルの領域を選択
        Dim oCells1 As Excel.Range = Nothing  ' 中継用セル
        Dim rngFrom1 As Excel.Range = Nothing  ' 始点セル指定用
        Dim rngTo1 As Excel.Range = Nothing    ' 終点セル指定用
        Dim rngTarget1 As Excel.Range = Nothing ' 貼付け範囲指定用
        Dim oCells2 As Excel.Range = Nothing  ' 中継用セル
        Dim rngFrom2 As Excel.Range = Nothing  ' 始点セル指定用
        Dim rngTo2 As Excel.Range = Nothing    ' 終点セル指定用
        Dim rngTarget2 As Excel.Range = Nothing ' 貼付け範囲指定用

        Dim rng部署1 As Excel.Range = Nothing
        Dim rng仕入先1 As Excel.Range = Nothing
        Dim rng仕入先2 As Excel.Range = Nothing

        Dim fileName As String = ""


        Dim 注文書明細S As Integer = 6
        Dim 精算書明細S As Integer = 10


        Dim Temp明細START列 As Integer = 1

        'Dim i明細数 As Integer = 10 'テンプレートの行数
        '一時Excelは不要らしいので直接原紙に読み取り専用で書き込むことにした
        'Dim f = DateTime.Now.ToString("yyyyMMddHHmmss")
        'outputExcel = "\\Fs1\d14it\一時フォルダ\Excel\" & f & ".xlsx"
        'System.IO.File.Copy(xlsTemplate, outputExcel, True)
        oApp.Visible = False
        oApp.DisplayAlerts = False

        'Dim dtOriginal As New DataTable
        Dim dtCopy As New DataTable
        Dim i As Integer

        'ヘッダー
        dtCopy.Columns.Add("部署")
        dtCopy.Columns.Add("仕入先")


        '明細
        'dtCopy.Columns.Add("部門")　IT内で話し合って不要という結論になった
        dtCopy.Columns.Add("購入者")
        dtCopy.Columns.Add("品名")
        dtCopy.Columns.Add("型式")
        dtCopy.Columns.Add("数量")
        dtCopy.Columns.Add("予算単価")
        'dtCopy.Columns.Add("科目")　IT内で話し合って不要という結論になった
        dtCopy.Columns.Add("購入理由")
        dtCopy.Columns.Add("決裁")
        dtCopy.Columns.Add("承認済")
        dtCopy.Columns.Add("印刷日")
        dtCopy.Columns.Add("pay_code")
        dtCopy.Columns.Add("id") 'DBのアップデートに必要
        Dim CopyRow As DataRow
        '注文書の部署はループの前に入れておく
        'rng部署1 = oSheet1.Range("B1")
        rng部署1 = oSheet1.Range("F25") '20220214 生技のリクエストを受けてレイアウト変更
        rng部署1.Value = Select_Form.s部署
        MRComObject(rng部署1)

        For i = 0 To DT購入品リスト.Rows.Count - 1
            '現金は除く

            'If ns(Ori("pay_code")) <> "現金" Then
            CopyRow = dtCopy.NewRow
            '印刷前に更新しているから必須項目の漏れはあり得ない
            CopyRow("部署") = DGV購入品入力.Rows(i).Cells("group_name").Value
            CopyRow("仕入先") = DGV購入品入力.Rows(i).Cells("仕入先").Value

            'CopyRow("部門") = ns(DGV購入品入力.Rows(i).Cells("部門").FormattedValue)　IT内で話し合って不要という結論になった
            CopyRow("購入者") = DGV購入品入力.Rows(i).Cells("購入者").Value
            CopyRow("品名") = DGV購入品入力.Rows(i).Cells("品名").Value
            CopyRow("型式") = kc.ns(DGV購入品入力.Rows(i).Cells("型式").Value)
            CopyRow("数量") = DGV購入品入力.Rows(i).Cells("数量").Value
            CopyRow("予算単価") = DGV購入品入力.Rows(i).Cells("予算単価").Value
            'CopyRow("科目") = ns(DGV購入品入力.Rows(i).Cells("科目").FormattedValue)　IT内で話し合って不要という結論になった
            CopyRow("購入理由") = kc.ns(DGV購入品入力.Rows(i).Cells("購入理由").Value)
            CopyRow("決裁") = kc.ns(DGV購入品入力.Rows(i).Cells("決裁").Value)
            CopyRow("承認済") = kc.ns(DGV購入品入力.Rows(i).Cells("承認済").Value)
            CopyRow("印刷日") = kc.ns(DGV購入品入力.Rows(i).Cells("印刷日").Value)
            CopyRow("pay_code") = DGV購入品入力.Rows(i).Cells("支払区分").Value
            CopyRow("id") = DGV購入品入力.Rows(i).Cells("id").Value
            dtCopy.Rows.Add(CopyRow)
            'End If
        Next

        Dim viw As New DataView(dtCopy)
        'dtOriginal = DT購入品リスト
        Dim dt仕入先リスト As DataTable = viw.ToTable(True, "仕入先") '重複を削除する
        dt仕入先リスト.Columns.Add("Count", GetType(Integer))
        'For Each row As DataRow In dt仕入先リスト.Rows
        '    Dim expr As String = String.Format("仕入先 = '{0}'", row("仕入先"))
        '    row("Count") = dtCopy.Compute("COUNT(数量)", expr)
        'Next
        Dim dt仕入先別 As New DataTable
        Dim dt1 As New DataTable
        Dim dt2 As New DataTable

        Dim str精算書ID As String = ""
        Dim str注文書ID As String = ""

        For i = 0 To dt仕入先リスト.Rows.Count - 1
            oApp.Visible = True
            '直接dtに入れると0件の時エラーになるのでカウントが1以上になるのを確認してから入れる
            If dtCopy.Select("仕入先 = '" & dt仕入先リスト.Rows(i)("仕入先") & "'").Count > 0 Then
                If dt仕入先別 IsNot Nothing Then
                    dt仕入先別.Clear()
                End If
                dt仕入先別 = dtCopy.Select("仕入先 = '" & dt仕入先リスト.Rows(i)("仕入先") & "'").CopyToDataTable

                '現金以外の場合(注文書の場合)
                If dt仕入先別.Select("pay_code <> '現金'").Count > 0 Then
                    If dt1 IsNot Nothing Then
                        dt1.Clear()
                    End If
                    dt1 = dt仕入先別.Select("pay_code <> '現金'").CopyToDataTable
                    dt1 = create印刷用DT(b未承認印刷, dt1)
                    If dt1 Is Nothing Then
                        Continue For
                    Else
                        'rng仕入先1 = oSheet1.Range("B1")　生技のリクエストでレイアウト変更
                        rng仕入先1 = oSheet1.Range("A3")
                        rng仕入先1.Value = dt仕入先リスト.Rows(i)("仕入先")
                        MRComObject(rng仕入先1)
                        印刷(dt1, oSheet1, oDialogs)
                    End If

                End If

                '現金の場合
                If dt仕入先別.Select("pay_code = '現金'").Count > 0 Then
                    If dt2 IsNot Nothing Then
                        dt2.Clear()
                    End If
                    dt2 = dt仕入先別.Select("pay_code = '現金'").CopyToDataTable
                    dt2 = create印刷用DT(b未承認印刷, dt2)
                    If dt2 Is Nothing Then
                        Continue For
                    Else

                        'rng仕入先2 = oSheet2.Range("C5")
                        rng仕入先2 = oSheet2.Range("B5") '2022/03/15 セル間違ってた
                        rng仕入先2.Value = dt仕入先リスト.Rows(i)("仕入先")
                        MRComObject(rng仕入先2)
                        印刷(dt2, oSheet2, oDialogs)

                    End If
                End If

            Else
                Continue For
            End If


        Next

        MRComObject(oSheet2)
        MRComObject(oSheet1)
        MRComObject(oSheets)
        oBook.Close(False)
        MRComObject(oBook)
        MRComObject(oBooks)
        MRComObject(oDialogs)

        oApp.Quit()
        oApp.DisplayAlerts = True
        MRComObject(oApp)
        If Flg印刷対象有 Then
            MsgBox("印刷終了")
        Else
            MsgBox("印刷対象がありません")
        End If
        Flg印刷対象有 = False
        'Call ProcessCheck()



    End Sub

    Sub 印刷(dt As DataTable, xlSheet As Excel.Worksheet, xlDialogs As Excel.Dialogs)
        'テンプレートの情報
        Const c明細開始行 As Integer = 10
        Const c明細開始列 As Integer = 1
        Const c明細行数 As Integer = 10


        Dim view As New DataView(dt)
        Dim DT印刷対象 As New DataTable

        Dim xlCell As Excel.Range
        Dim xl始点 As Excel.Range
        Dim xl終点 As Excel.Range
        Dim xl明細範囲 As Excel.Range
        Dim xlDialog As Excel.Dialog
        xlDialog = xlDialogs(Excel.XlBuiltInDialog.xlDialogPrint)

        Dim sep As String = ""
        Dim s印刷日更新対象ID As String
        Dim i明細最終列 As Integer
        Dim Arr(,) As Object = Nothing
        Dim strSql As String

        'DT印刷対象 = view.ToTable(False, {"部門", "購入者", "品名", "型式", "数量", "科目", "購入理由", "決裁"}) 'false = 重複を削除しない
        DT印刷対象 = view.ToTable(False, {"購入者", "品名", "型式", "数量", "予算単価", "購入理由", "決裁"})
        i明細最終列 = DT印刷対象.Columns.Count
        'rng仕入先2 = oSheet2.Range("C5")
        'rng仕入先2.Value = dt仕入先リスト.Rows(i)("仕入先")
        'MRComObject(rng仕入先2)


        ' 配列を明細の範囲に貼り付け
        'テンプレートの行数を上回ったら次の頁へ
        Dim Page As Integer 'ページ数

        '総ページ数
        Dim Total As Integer = Math.Ceiling(DT印刷対象.Rows.Count / 10) - 1

        '印刷ページのループ
        For Page = 0 To Total

            '1頁ごとに印刷済みにアップデートするので初期化している
            s印刷日更新対象ID = ""
            sep = "" '2022/03/15 ekawai add

            '二次元配列データ作成　
            ReDim Arr(c明細行数 - 1, i明細最終列 - 1)
            Dim counter As Integer = 1
            Try
                '行単位のループ
                For iRow As Integer = 0 To c明細行数 - 1
                    If counter <= DT印刷対象.Rows.Count - (c明細行数 * Page) Then
                        s印刷日更新対象ID = s印刷日更新対象ID & sep & kc.SQ(dt.Rows(iRow + (10 * Page))("id"))
                        sep = ","
                    End If

                    For iCol As Integer = 0 To DT印刷対象.Columns.Count - 1
                        'If counter <= DT印刷対象.Rows.Count Then 2022/03/15 ekawai del　2枚以上の印刷ができなかった
                        If counter <= DT印刷対象.Rows.Count - (c明細行数 * Page) Then '2022/03/15 ekawai add 
                            Arr(iRow, iCol) = DT印刷対象.Rows(iRow + (c明細行数 * Page))(iCol)
                        Else
                            '空白を入れる
                            Arr(iRow, iCol) = ""
                        End If

                    Next
                    counter = counter + 1
                Next

            Catch ex As Exception
                Throw
            End Try
            xlCell = xlSheet.Cells
            '始点セル
            xl始点 = DirectCast(xlCell(c明細開始行, c明細開始列), Excel.Range)


            '終点セル
            xl終点 = DirectCast(xlCell(c明細開始行 + c明細行数 - 1, i明細最終列), Excel.Range)


            '貼り付け範囲作成
            xl明細範囲 = xlSheet.Range(xl始点, xl終点)

            MRComObject(xl終点)
            MRComObject(xl始点)
            MRComObject(xlCell)


            xl明細範囲.Value = Arr

            xlSheet.Select()
            '印刷プレビューだとプリント押すとプリンタ選択なしで印刷されてしまうので全部署印刷ダイアログ出すようにした
            Try
                xlDialog.Show()
            Catch ex As System.Runtime.InteropServices.COMException
                'PCFAXのダイヤログ開いた後で「キャンセル」押すとハンドルされていない例外が発生する対策
            End Try
 			Flg印刷対象有 = True　'2022/03/15 ekawai add
            xl明細範囲.ClearContents()


            MRComObject(xl明細範囲)
            'DTに印刷日入れて再印刷できなくする(例えダイアログでキャンセルしていても印刷日は書き込む)

            strSql = ""
            strSql &= vbCrLf & "UPDATE TD_po SET 印刷日 = " & kc.SQ(Now)
            'strSql &= vbCrLf & ",UpdateStamp = " & SQ(Now) 印刷日記録しているからUpdateStamp更新しないことにした。印刷前に変更していた場合に追えなくなるから。
            strSql &= vbCrLf & "WHERE id IN (" & s印刷日更新対象ID & ")"
            Dim tran As System.Data.SqlClient.SqlTransaction = Nothing

            tran = dCon.Connection.BeginTransaction


            If dCon.ExecuteSqlMW(tran, strSql) = False Then
                tran.Rollback()
                MsgBox("印刷日の更新に失敗しました") '2022/03/15 ekawai add
            Else
                tran.Commit()

            End If
            tran.Dispose()
        Next

        MRComObject(xlDialog)
    End Sub
    '≪テスト用≫プロセスチェック
    Private Sub ProcessCheck()
        'タスクマネージャに、Excel.exe が残っていないか確認(テスト環境でのみ使用の事)
        Dim st As Integer = System.Environment.TickCount
        '以前は、Loop しながら5秒間程繰り返し確認していたのだが、その間に解放される場合が
        'ある事が判明したので、下記のように1回きりの確認でもデクリメント処理がキチンと
        '行われていたら解放される事が解ったので下記のように厳密に判定する事にしました。
        System.Threading.Thread.Sleep(1000)
        Application.DoEvents()
        If Process.GetProcessesByName("Excel").Length = 0 Then
            '先にフォームを閉じるとエラーが発生するので
            '必要により表示するなりコメントにして下さい。
            MessageBox.Show(Me, "Excel.EXE は解放されました。")
            Exit Sub
        End If
        If Process.GetProcessesByName("Excel").Length >= 1 Then
            Dim ret As DialogResult
            ret = MessageBox.Show(Me, "まだ Excel.EXE が起動しています。強制終了しますか？", _
                                                            "確認", MessageBoxButtons.YesNo)
            If ret = Windows.Forms.DialogResult.Yes Then
                Dim localByName As Process() = Process.GetProcessesByName("Excel")
                Dim p As Process
                '起動中のExcelを取得
                For Each p In localByName
                    'Windowの無い(表示していない)Excel があれば強制終了させる
                    '画面に表示している Excel は、終了させないので必要なら手動で終了して下さい。
                    If System.String.Compare(p.MainWindowTitle, "", True) = 0 Then
                        'Excel.EXE のプロセスを削除
                        p.Kill()
                    End If
                Next
            End If
        End If
    End Sub

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

    Private Sub chk未印刷_CheckedChanged(sender As Object, e As EventArgs) Handles chk未印刷.CheckedChanged


        If SwEditSave Then
            'If MsgBox("保存せずに画面を切り替えますか？", vbYesNo + vbDefaultButton1) = vbNo Then
            '    '元に戻す
            '    If chk未印刷.Checked = True Then
            '        RemoveHandler Me.chk未印刷.CheckedChanged, AddressOf chk未印刷_CheckedChanged
            '        chk未印刷.Checked = False
            '        AddHandler Me.chk未印刷.CheckedChanged, AddressOf chk未印刷_CheckedChanged
            '    Else
            '        RemoveHandler Me.chk未印刷.CheckedChanged, AddressOf chk未印刷_CheckedChanged
            '        chk未印刷.Checked = True
            '        AddHandler Me.chk未印刷.CheckedChanged, AddressOf chk未印刷_CheckedChanged
            '    End If

            'Else
            '    If chk未印刷.Checked = True And Select_Form.更新Flg = True Then
            '        btn注文書印刷.Visible = True
            '        btn未承認印刷.Visible = True
            '    Else
            '        btn注文書印刷.Visible = False
            '        btn未承認印刷.Visible = False
            '    End If
            '    SwEditSave = False
            '    DGV購入品入力_表示()

            'End If
            If MsgBox("保存してよろしいですか?" & vbCrLf & "「いいえ」を押すと更新していないデータは消えます", vbYesNo + vbDefaultButton1) = vbYes Then
                If メイン更新() = False Then
                    '更新失敗したら画面を書き換えずに終了
                    Exit Sub
                End If
            Else
                '本来は更新成功しないとSqEditSave = Trueにならないが、この場合は強制的にFalseにする(そうしないと更新するまで画面切り替える度、永遠に更新しますか？というメッセージが出てしまうため)
                SwEditSave = False
            End If
        End If

        DGV購入品入力_表示()
        DGV購入品入力_セル設定()

        If chk未印刷.Checked = True And (Select_Form.更新Flg = True Or Select_Form.承認Flg = True) Then
            btn注文書印刷.Visible = True
            btn未承認印刷.Visible = True
        Else
            btn注文書印刷.Visible = False
            btn未承認印刷.Visible = False
        End If

    End Sub

    Private Sub rbn未承認_CheckedChanged(sender As Object, e As EventArgs) Handles rbn未承認.CheckedChanged

        If Select_Form.承認Flg = True And rbn未承認.Checked = True Then
            btn一括承認.Visible = True
            btn一括稟議承認.Visible = True
        Else
            btn一括承認.Visible = False
            btn一括稟議承認.Visible = False
        End If

        If SwEditSave Then

            'If MsgBox("保存せずに画面を切り替えますか？", vbYesNo + vbDefaultButton1) = vbNo Then
            '    '元に戻す
            '    If rbn未承認.Checked = True Then
            '        RemoveHandler Me.rbn未承認.CheckedChanged, AddressOf rbn未承認_CheckedChanged
            '        rbn未承認.Checked = False
            '        AddHandler Me.rbn未承認.CheckedChanged, AddressOf rbn未承認_CheckedChanged
            '    Else
            '        RemoveHandler Me.rbn未承認.CheckedChanged, AddressOf rbn未承認_CheckedChanged
            '        rbn未承認.Checked = True
            '        AddHandler Me.rbn未承認.CheckedChanged, AddressOf rbn未承認_CheckedChanged
            '    End If

            'Else
            '    If rbn未承認.Checked = True And Select_Form.承認Flg = True Then
            '        btn一括承認.Visible = True
            '    Else
            '        btn一括承認.Visible = False
            '    End If
            '    SwEditSave = False
            '    DGV購入品入力_表示()

            'End If

            If MsgBox("保存してよろしいですか?" & vbCrLf & "「いいえ」を押すと更新していないデータは消えます", vbYesNo + vbDefaultButton1) = vbYes Then

                If メイン更新() = False Then
                    '更新失敗したら画面を書き換えずに終了
                    Exit Sub
                End If
            Else
                '本来は更新成功しないとSqEditSave = Trueにならないが、この場合は強制的にFalseにする(そうしないと更新するまで画面切り替える度、永遠に更新しますか？というメッセージが出てしまうため)
                SwEditSave = False
            End If

        End If


        DGV購入品入力_表示()
        DGV購入品入力_セル設定()

    End Sub
    Private Sub rbn承認済_CheckedChanged(sender As Object, e As EventArgs) Handles rbn承認済.CheckedChanged

        btn一括承認.Visible = False
        btn一括稟議承認.Visible = False
        If SwEditSave Then
            'If MsgBox("保存せずに画面を切り替えますか？", vbYesNo + vbDefaultButton1) = vbNo Then
            ''元に戻す
            'If rbn承認済.Checked = True Then
            '    RemoveHandler Me.rbn承認済.CheckedChanged, AddressOf rbn承認済_CheckedChanged
            '    rbn承認済.Checked = False
            '    AddHandler Me.rbn承認済.CheckedChanged, AddressOf rbn承認済_CheckedChanged
            'Else
            '    RemoveHandler Me.rbn承認済.CheckedChanged, AddressOf rbn承認済_CheckedChanged
            '    rbn承認済.Checked = True
            '    AddHandler Me.rbn承認済.CheckedChanged, AddressOf rbn承認済_CheckedChanged
            'End If

            'Else
            '    btn一括承認.Visible = False
            '    SwEditSave = False
            '    DGV購入品入力_表示()

            'End If
            If MsgBox("保存してよろしいですか?" & vbCrLf & "「いいえ」を押すと更新していないデータは消えます", vbYesNo + vbDefaultButton1) = vbYes Then
                If メイン更新() = False Then
                    '更新失敗したら画面を書き換えずに終了
                    Exit Sub
                End If
            Else
                '本来は更新成功しないとSqEditSave = Trueにならないが、この場合は強制的にFalseにする(そうしないと更新するまで画面切り替える度、永遠に更新しますか？というメッセージが出てしまうため)
                SwEditSave = False
            End If

        End If


        DGV購入品入力_表示()
        DGV購入品入力_セル設定()
    End Sub


    Private Sub btn未承認印刷_Click(sender As Object, e As EventArgs) Handles btn未承認印刷.Click
        If SwEditSave Then
            If MsgBox("保存してよろしいですか?" & vbCrLf & "「いいえ」を押すと更新していないデータは消えます", vbYesNo + vbDefaultButton1) = vbYes Then
                If メイン更新() = False Then
                    '更新失敗したら画面を書き換えずに終了
                    Exit Sub
                End If
            Else
                '本来は更新成功しないとSqEditSave = Trueにならないが、この場合は強制的にFalseにする(そうしないと更新するまで画面切り替える度、永遠に更新しますか？というメッセージが出てしまうため)
                SwEditSave = False

            End If
        End If
        Call 印刷準備(True)
        Call DGV購入品入力_表示()
    End Sub

    Private Sub btn注文書印刷_Click(sender As Object, e As EventArgs) Handles btn注文書印刷.Click
        '印刷項目を毎回ユーザーに選ばせる案もあったが、何の単位で選ばせればいいのか判断が難しい。
        '仕入先単位にした場合、1回の印刷で5つ仕入先があったら5回確認画面が出てボタン押すまで印刷できないから、さすがにそれは手間だということになった。(既にダイヤログボックスで1回止めてるから1仕入先につき2回ボタン押すことになる)
        'さらに、印刷って管理者が行う場合もあるが、管理者がこの仕入先はこの項目要らないとか判断できないよねということになり、ITで勝手にテンプレート決めることにした。
        'その結果部門と科目は社外の人には要らないから消して、昔リクエストがあった単価の列を増やし、品名と型式の幅を広げることにした。→運用開始後品質から単価いらないという意見があったので消した
        If SwEditSave Then
            If MsgBox("保存してよろしいですか?" & vbCrLf & "「いいえ」を押すと更新していないデータは消えます", vbYesNo + vbDefaultButton1) = vbYes Then
                If メイン更新() = False Then
                    '更新失敗したら画面を書き換えずに終了
                    Exit Sub
                End If
            Else
                '本来は更新成功しないとSqEditSave = Trueにならないが、この場合は強制的にFalseにする(そうしないと更新するまで画面切り替える度、永遠に更新しますか？というメッセージが出てしまうため)
                SwEditSave = False
            End If
        End If
        Call 印刷準備(False)
        Call DGV購入品入力_表示()
    End Sub

    Private Function create印刷用DT(bl未承認印刷 As Boolean, dt As DataTable) As DataTable
        If bl未承認印刷 Then
            '未承認のみ絞り込む
            If dt.Select("承認済 = 'FALSE'").Count > 0 Then
                dt = dt.Select("承認済 = 'FALSE'").CopyToDataTable
            Else
                Return Nothing

            End If
        Else
            '承認済みしか印刷しない場合は承認済みのみ絞り込む
            If dt.Select("承認済 = 'TRUE'").Count > 0 Then
                dt = dt.Select("承認済 = 'TRUE'").CopyToDataTable

            Else

                Return Nothing
            End If
        End If

        Return dt
    End Function

    Private Sub 配列格納(ByVal dt As DataTable, ByRef strID As String, ByRef arr(,) As Object)
        Dim view As New DataView(dt)
        Dim dtID As DataTable
        Dim sep As String = ""

        dtID = view.ToTable(False, "id")
        dt = view.ToTable(False, {"部門", "購入者", "品名", "型式", "数量", "科目", "購入理由", "決裁"}) '重複を削除しない


        '二次元配列データ作成　
        ReDim arr(dt.Rows.Count - 1, dt.Columns.Count - 1)
        Try
            For iRow As Integer = 0 To dt.Rows.Count - 1
                For iCol As Integer = 0 To dt.Columns.Count - 1
                    arr(iRow, iCol) = dt.Rows(iRow)(iCol)
                Next
            Next

        Catch ex As Exception
            Throw
        End Try

        For Each row In dtID.Rows
            strID = strID & sep & kc.SQ(row("id"))
            sep = ","
        Next

    End Sub

    Private Sub DGV購入品入力_セル設定(Optional ByVal s除外購入ID As String = "")
        '2022/11/02　del S ReadOnly使わないように指示があったため
        'If Select_Form.更新Flg = False And Select_Form.承認Flg = False And Select_Form.経理Flg = False Then
        '    '権限なかったら
        '    DGV購入品入力.ReadOnly = True

        'Else

        '    '承認者だったら決裁列のロックを解除
        '    If Select_Form.承認Flg Then
        '        DGV購入品入力.Columns("決裁候補").ReadOnly = False
        '        'DGV購入品入力.Columns("決裁").ReadOnly = False '候補以外選ばせたくないから編集不可
        '    End If

        '    '経理だったら経理検査列のロックを解除
        '    If Select_Form.経理Flg Then
        '        DGV購入品入力.Columns("経理検査").ReadOnly = False

        '    End If


        '    For Each DGVRow As DataGridViewRow In DGV購入品入力.Rows

        '        If kc.nz(DGVRow.Cells("経理検査").Value) = True Then
        '            '全列ロック
        '            DGVRow.Cells("印刷日").ReadOnly = True
        '            DGVRow.Cells("発注日").ReadOnly = True
        '            DGVRow.Cells("入荷日").ReadOnly = True
        '            DGVRow.Cells("登録番号").ReadOnly = True
        '            DGVRow.Cells("部門").ReadOnly = True
        '            DGVRow.Cells("購入者").ReadOnly = True
        '            DGVRow.Cells("購入者候補").ReadOnly = True
        '            DGVRow.Cells("品名").ReadOnly = True
        '            DGVRow.Cells("型式").ReadOnly = True
        '            DGVRow.Cells("仕入先").ReadOnly = True
        '            DGVRow.Cells("仕入先候補").ReadOnly = True
        '            DGVRow.Cells("改訂番号").ReadOnly = True

        '            DGVRow.Cells("支払区分").ReadOnly = True
        '            DGVRow.Cells("数量").ReadOnly = True
        '            DGVRow.Cells("予算単価").ReadOnly = True
        '            DGVRow.Cells("税抜単価").ReadOnly = True
        '            DGVRow.Cells("科目番号").ReadOnly = True
        '            DGVRow.Cells("科目").ReadOnly = True
        '            DGVRow.Cells("科目ID").ReadOnly = True
        '            DGVRow.Cells("購入理由").ReadOnly = True
        '            DGVRow.Cells("ワークコード").ReadOnly = True
        '            DGVRow.Cells("備考1").ReadOnly = True
        '            DGVRow.Cells("棚番号").ReadOnly = True
        '            DGVRow.Cells("整理番号").ReadOnly = True
        '        Else
        '            '経理検査がFalseの時経理なら一切ロックしない
        '            If Select_Form.経理Flg = True Then

        '            Else
        '                '更新者や承認者だったら承認済みの場合は一部ロック
        '                'If DGVRow.Cells("承認済").Value IsNot DBNull.Value Then
        '                If kc.nz(DGVRow.Cells("承認済").Value) = True Then
        '                    '旧購入依頼ではFixRowという箇所で行っていた処理
        '                    DGVRow.Cells("登録番号").ReadOnly = True
        '                    DGVRow.Cells("品名").ReadOnly = True
        '                    '行コピーの場合型式と数量はロックしない
        '                    If s除外購入ID = DGVRow.Cells("購入ID").Value Then
        '                        DGVRow.Cells("型式").ReadOnly = False
        '                        DGVRow.Cells("数量").ReadOnly = False

        '                    Else
        '                        DGVRow.Cells("型式").ReadOnly = True
        '                        DGVRow.Cells("数量").ReadOnly = True

        '                    End If
        '                    DGVRow.Cells("予算単価").ReadOnly = True
        '                End If
        '                'End If

        '            End If
        '        End If
        '    Next

        'End If
        '2022/11/02 del E
        If DGV購入品入力.Rows.Count > 0 Then
            'グリッドに表示される最初（一番上）の行の取得／設定
            DGV購入品入力.FirstDisplayedScrollingRowIndex = DGV購入品入力.Rows.Count - 1
        End If

    End Sub

    Private Sub DGV購入品入力_Sorted(sender As Object, e As EventArgs) Handles DGV購入品入力.Sorted
        Dim dgv As DataGridView = CType(sender, DataGridView)
        If dgv.Rows.Count > 1 Then
            DGV購入品入力_セル設定()
        End If
    End Sub


    Private Sub btnマスタ_Click(sender As Object, e As EventArgs) Handles btnマスタ.Click
        Dim f As New Master_Form
        'f.ShowDialog(Me)
        'f.Close()
        'f.Dispose()
        'Me.Close()
        f.Show()
    End Sub


    'Private Sub ミスミ照合済みデータ取込()
    '    Dim DT As DataTable
    '    Dim strSql As String

    '    DT = Read_ミスミCSV("\\Fs1\030購入品管理\技術部\照合済.csv")
    '    If DT Is Nothing Then

    '    Else

    '        Dim tran As System.Data.SqlClient.SqlTransaction = Nothing

    '        tran = dCon.Connection.BeginTransaction

    '        For Each DR As DataRow In DT.Rows
    '            strSql = "INSERT INTO TD_po("
    '            strSql &= vbCrLf & "vendor"
    '            strSql &= vbCrLf & ",vendor_id"
    '            strSql &= vbCrLf & ",pay_code"
    '            'strSql &= vbCrLf & ",employee"
    '            strSql &= vbCrLf & ",order_date"
    '            strSql &= vbCrLf & ",arrival_date"
    '            strSql &= vbCrLf & ",part_name"
    '            strSql &= vbCrLf & ",part_number"
    '            strSql &= vbCrLf & ",number"
    '            strSql &= vbCrLf & ",unit_price"
    '            strSql &= vbCrLf & ",remark"
    '            strSql &= vbCrLf & ") VALUES ("
    '            strSql &= vbCrLf & "'ミスミ'"
    '            strSql &= vbCrLf & "'m15'"
    '            strSql &= vbCrLf & ",'掛け'"
    '            'strSql &= vbCrLf & "," & SQ(DR.Item("注文担当者"))
    '            strSql &= vbCrLf & "," & SQ(DR.Item("発注"))
    '            strSql &= vbCrLf & "," & SQ(DR.Item("入荷"))
    '            strSql &= vbCrLf & "," & SQ(DR.Item("品名"))
    '            strSql &= vbCrLf & "," & SQ(DR.Item("型式"))
    '            strSql &= vbCrLf & "," & DR.Item("数量")
    '            strSql &= vbCrLf & "," & DR.Item("税抜単価")
    '            strSql &= vbCrLf & "," & SQ(DR.Item("購入理由"))
    '            strSql &= vbCrLf & ")"

    '            If dCon.ExecuteSqlMW(tran, strSql) = False Then
    '                tran.Rollback()
    '                tran.Dispose()
    '                MsgBox("更新失敗")
    '                Exit Sub
    '            End If



    '        Next

    '        tran.Commit()
    '        tran.Dispose()

    '    End If

    'End Sub
    ''TextFieldParserクラスによるCSVファイルの読み込み
    'Private Function Read_ミスミCSV(strFileName As String) As DataTable
    '    Read_ミスミCSV = Nothing
    '    Dim ミスミDT As New DataTable

    '    'CSVファイルをコンストラクタで指定してインスタンスを作成する
    '    Dim parser As Microsoft.VisualBasic.FileIO.TextFieldParser = Nothing
    '    'Shift-JISエンコードで変換できない場合は「?」文字の設定
    '    Dim encFallBack As System.Text.DecoderReplacementFallback = New System.Text.DecoderReplacementFallback("?")
    '    Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis", System.Text.EncoderFallback.ReplacementFallback, encFallBack)
    '    'TextFieldParserクラス(ファイルパス,文字コード)
    '    parser = New Microsoft.VisualBasic.FileIO.TextFieldParser(strFileName, enc)

    '    '区切りの指定
    '    With parser
    '        .TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
    '        .SetDelimiters(",")
    '        '空白があった場合にTrimしない
    '        .TrimWhiteSpace = False
    '    End With

    '    'CSV読込実行
    '    Dim strArr()() As String = Nothing
    '    Dim nLine As Integer = 0
    '    Dim iLoopCnt As Integer = 0
    '    While Not parser.EndOfData
    '        Try
    '            Dim ArrCSV As String() = parser.ReadFields()
    '            Dim ColCnt = 0
    '            'iCntが0の時は1回目のループだからタイトル行
    '            If iLoopCnt = 0 Then
    '                ミスミDT.Clear()
    '                For Each col As String In ArrCSV
    '                    'ミスミDTに見出しを書き込む
    '                    ミスミDT.Columns.Add(col, GetType(System.String))
    '                Next

    '            Else
    '                Dim mRow As DataRow = ミスミDT.NewRow
    '                Dim rCnt As Integer = 0
    '                For Each col As String In ArrCSV
    '                    mRow.Item(rCnt) = col
    '                    rCnt += 1
    '                Next
    '                ミスミDT.Rows.Add(mRow)
    '            End If

    '        Catch ex As Exception
    '            Return Nothing
    '        End Try
    '        iLoopCnt += 1

    '    End While

    '    Return ミスミDT
    'End Function

    Private Sub btn技術DB_Click(sender As Object, e As EventArgs) Handles btn技術DB.Click
        '仮コピーとDailyInsert(ブック閉じる時のINSERT)の機能を１つにする。期間指定は不要。期間関係なく入荷IDが空白のものを対象とする
        '既存の購入依頼では、月末集計ボタンを押した時にTD_gijutsu_nyukaをアップデートしていたが、
        '後で修正が発生した場合は、購入依頼と技術部access両方修正してもらうように変わった。(間違えたら担当者が修正するのは当然なので、そんなことまで配慮しないでいいらしい)
        '確定コピーボタンの処理ははbtn出荷データ作成_Click
        Dim iRow As Integer
        Dim NEXT入荷ID As Integer
        Dim DT入荷ID As DataTable
        'DataSourceでバインドしている時では初期化
        Dim Kamoku As String '科目用変数

        Dim strSql As String = ""
        Dim tran As System.Data.SqlClient.SqlTransaction = Nothing
        'グリッド編集中にボタンを押してしまう可能性も考えて、強制的にグリッドの更新処理を行うことになった
        btn更新.PerformClick()

        技術処理対象_Read()

        '2021/08/25 ITで話し合った結果
        '本当は更新ボタンと動きを合わせるために、1行でもエラーが発生したらTD_gijutsu_nyukaへのINSERTは全件中止させたかったが
        'TD_gijutsu_nyukaへのINSERTの直後にNyukaIDの最大値を取得して入荷IDをTD_poという必要があるため諦めるしかないという結論になった。
        'TD_poのUPDATEはよっぽどエラーが起きないそうなので、TD_gijutsu_nyukaのINSERTだけ1行ずつコミットする形になった
        If DT購入品リスト.Rows.Count > 0 Then
            For iRow = 0 To DT購入品リスト.Rows.Count - 1
                Kamoku = DGV購入品入力.Rows(iRow).Cells("科目").FormattedValue
                Kamoku = Kamoku.Substring(Kamoku.IndexOf(":") + 2) '科目名のみ切り出す　:の次に半角スペースが入っているから+2している
                strSql = ""
                strSql = "insert into TD_gijutsu_nyuka ("
                strSql = strSql & vbCrLf & "Kubun0,"  '部門 ex)金型保全部門
                strSql = strSql & vbCrLf & "Kubun1,"  '新品・再利用の区分
                strSql = strSql & vbCrLf & "Kubun2," '仕入先番号
                strSql = strSql & vbCrLf & "TanaNo," '棚番号
                strSql = strSql & vbCrLf & "Rev,"    '改訂番号 
                strSql = strSql & vbCrLf & "SeiriNo,"  '整理番号
                strSql = strSql & vbCrLf & "Name1,"  '品名
                strSql = strSql & vbCrLf & "Name2,"  '型番
                strSql = strSql & vbCrLf & "Name3,"  '科目
                strSql = strSql & vbCrLf & "Name4,"
                strSql = strSql & vbCrLf & "Num,"
                strSql = strSql & vbCrLf & "UnitPrice,"
                strSql = strSql & vbCrLf & "OrderDate," '発注日 
                strSql = strSql & vbCrLf & "NyukaDate,"
                strSql = strSql & vbCrLf & "PrintYN,"
                strSql = strSql & vbCrLf & "PointOfOrder,"
                strSql = strSql & vbCrLf & "OrderNum,"
                strSql = strSql & vbCrLf & "erase_inventory"
                strSql = strSql & vbCrLf & ")"
                strSql = strSql & vbCrLf & "values("
                strSql = strSql & vbCrLf & " '" & DGV購入品入力.Rows(iRow).Cells("部門").FormattedValue & "',"
                strSql = strSql & vbCrLf & " '新',"
                strSql = strSql & vbCrLf & " '" & DGV購入品入力.Rows(iRow).Cells("仕入先ID").Value & "',"
                strSql = strSql & vbCrLf & " '" & DGV購入品入力.Rows(iRow).Cells("棚番号").Value & "',"
                strSql = strSql & vbCrLf & " '" & DGV購入品入力.Rows(iRow).Cells("改訂番号").Value & "'," '2020/01/22 ekawai add
                strSql = strSql & vbCrLf & " '" & DGV購入品入力.Rows(iRow).Cells("整理番号").Value & "',"
                strSql = strSql & vbCrLf & " '" & DGV購入品入力.Rows(iRow).Cells("品名").Value & "',"
                strSql = strSql & vbCrLf & " '" & DGV購入品入力.Rows(iRow).Cells("型式").Value & "',"
                'strSql = strSql & vbCrLf & " '" & DGV購入品入力.Rows(iRow).Cells("科目").FormattedValue & "',"　グリッドの見た目のままDBに取り込んでしまっていたのでコメントアウト
                strSql = strSql & vbCrLf & " '" & Kamoku & "',"
                strSql = strSql & vbCrLf & "'',"
                strSql = strSql & vbCrLf & DGV購入品入力.Rows(iRow).Cells("数量").Value & ","
                strSql = strSql & vbCrLf & DGV購入品入力.Rows(iRow).Cells("税抜単価").Value & ","
                strSql = strSql & vbCrLf & "'" & DGV購入品入力.Rows(iRow).Cells("発注日").Value & "'," '2020/01/22 ekawai add
                strSql = strSql & vbCrLf & "'" & DGV購入品入力.Rows(iRow).Cells("入荷日").Value & "',"
                strSql = strSql & vbCrLf & "0,"
                strSql = strSql & vbCrLf & "0,"
                strSql = strSql & vbCrLf & "0,"
                strSql = strSql & vbCrLf & "1"
                strSql = strSql & vbCrLf & ")"

                tran = dCon.Connection.BeginTransaction
                If dCon.ExecuteSqlMW(tran, strSql) = False Then
                    tran.Rollback()
                    MsgBox(iRow & "行目でエラーが発生しました。これ以降の行の処理を中止します。")
                    Exit Sub
                Else
                    tran.Commit()
                    tran.Dispose()

                End If

                'TD_poの入荷ID書き込み＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
                'TD_gijutsu_nyukaは自動採番のため、ミヤマ番号管理を使っている購入IDとは同じように処理できないので、
                '旧購入依頼同様に直前のINSERT終わった時点のMAX(NyukaID)が対象ということにしてTD_poに書き込むしかないという結論になった
                strSql = ""
                strSql = "SELECT MAX(NyukaID) as 技術入荷ID FROM TD_gijutsu_nyuka"
                DT入荷ID = dCon.DataSet(strSql, "技術入荷ID").Tables(0)
                NEXT入荷ID = DT入荷ID.Rows(0)("技術入荷ID")

                strSql = ""
                strSql = "UPDATE TD_po SET"
                strSql = strSql & vbCrLf & "入荷ID = " & NEXT入荷ID
                strSql = strSql & vbCrLf & "WHERE id = '" & DGV購入品入力.Rows(iRow).Cells("id").Value & "'"
                'TD_poのUPDATEはよほど失敗しないそうなのでトランザクション使わないことにした
                dCon.Command() 'tran.disposeしているから、dCon.ExecuteSQL使うためにはもう1回Newしないといけないらしい　
                dCon.ExecuteSQL(strSql)
                '＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
            Next
            MsgBox("技術入荷テーブルへの書き込みが全て終了しました")
        Else
            MsgBox("技術入荷テーブルへの書き込み対象がありません")
        End If
        DGV購入品入力_表示()
        DGV購入品入力_セル設定()
    End Sub
    Private Sub 技術処理対象_Read()
        Flg再表示 = True
        Dim ds購入 As DataSet
        'DataSourceでバインドしている時は初期化
        DGV購入品入力.DataSource = Nothing

        Dim strSql As String = ""

        strSql = strSql & vbCrLf & "SELECT id"
        strSql = strSql & vbCrLf & ", 購入ID"
        strSql = strSql & vbCrLf & ",承認済"
        strSql = strSql & vbCrLf & ",決裁"
        strSql = strSql & vbCrLf & ",group_name"
        strSql = strSql & vbCrLf & ", format(印刷日, 'yyyy/MM/dd') AS 印刷日"
        strSql = strSql & vbCrLf & ", format(order_date, 'yyyy/MM/dd') AS 発注日"
        strSql = strSql & vbCrLf & ", format(arrival_date, 'yyyy/MM/dd') AS 入荷日"
        strSql = strSql & vbCrLf & ",登録番号"
        strSql = strSql & vbCrLf & ", division_id AS 部門ID"
        strSql = strSql & vbCrLf & ", employee AS 購入者"
        strSql = strSql & vbCrLf & ", part_name AS 品名"
        strSql = strSql & vbCrLf & ", part_number AS 型式"
        strSql = strSql & vbCrLf & ", vendor_id　AS 仕入先ID"
        strSql = strSql & vbCrLf & ", vendor as 仕入先"
        strSql = strSql & vbCrLf & ", 改訂番号"
        'strSql = strSql & vbCrLf & ", 営業所"
        strSql = strSql & vbCrLf & ", pay_code"
        strSql = strSql & vbCrLf & ", number AS 数量"
        strSql = strSql & vbCrLf & ", 予算単価"
        strSql = strSql & vbCrLf & ", unit_price AS 税抜単価"
        strSql = strSql & vbCrLf & ", kamoku_id　AS　科目ID"
        strSql = strSql & vbCrLf & ", remark AS 購入理由"
        strSql = strSql & vbCrLf & ", work_code AS ワークコード"
        strSql = strSql & vbCrLf & ", tana_bango AS 棚番号"
        strSql = strSql & vbCrLf & ", 整理番号"
        strSql = strSql & vbCrLf & ", 入荷ID"
        strSql = strSql & vbCrLf & ", 経理検査"
        strSql = strSql & vbCrLf & "FROM TD_po"
        strSql = strSql & vbCrLf & "WHERE group_name = " & kc.SQ(Select_Form.s部署)
        strSql = strSql & vbCrLf & "AND 入荷ID IS NULL"
        strSql = strSql & vbCrLf & "AND order_date IS NOT NULL"
        strSql = strSql & vbCrLf & "AND arrival_date IS NOT NULL"
        strSql = strSql & vbCrLf & "AND vendor_id IS NOT NULL"
        strSql = strSql & vbCrLf & "AND unit_price IS NOT NULL"

        'strSql = strSql & vbCrLf & "ORDER BY id"
        strSql = strSql & vbCrLf & "ORDER BY 購入ID,id"
        ds購入 = dCon.DataSet(strSql, "入荷テーブル更新対象")
        DT購入品リスト = ds購入.Tables("入荷テーブル更新対象")
        DGV購入品入力.DataSource = DT購入品リスト
        If DGV購入品入力.Rows.Count > 0 Then
            DGV購入品入力.FirstDisplayedScrollingRowIndex = DGV購入品入力.Rows.Count - 1
        End If

        DGV購入品入力.RowsDefaultCellStyle.BackColor = Color.FromArgb(221, 235, 247)
        DGV購入品入力.AlternatingRowsDefaultCellStyle.BackColor = Color.White

        Flg再表示 = False
    End Sub
    Private Function Check技術用()
        Dim dt As DataTable = DGV購入品入力.DataSource
        For Each DR As DataRow In dt.Rows
            If DR("発注日").ToString = "" Then
                Return False
            End If
            If DR("入荷日").ToString = "" Then
                Return False
            End If
            If DR("部門ID").ToString = "" Then
                Return False
            End If
            If DR("数量").ToString = "" Then
                Return False
            End If
            If DR("仕入先ID").ToString = "" Then
                Return False
            End If

        Next



    End Function



    'このイベントは通常、セルが編集されたが変更がデータキャッシュにコミットされていない場合、または編集操作がキャンセルされた場合に発生します。
    'このイベント内でDataGridView.CommitEditメソッドを呼び出して値をコミットすると瞬時に値が反映されるようになるらしい


    Private Sub GET仕入先ID(sName As String, iRow As Integer)
        Dim DR As DataRow()
        DGV購入品入力("仕入先ID", iRow).Value = DBNull.Value
        If sName <> "" Then
            DR = DT部署別仕入先.Select("kiban = " & kc.SQ(sName))
            Select Case DR.Length
                Case 0
                    MsgBox(Select_Form.s部署 & "の仕入先マスタに存在しない仕入先が入力されました。")

                Case 1
                    'DGV購入品入力("仕入先ID", iRow).Value = DR(0)("facility_id") '2021/04/12 ekawai del
                    DGV購入品入力("仕入先ID", iRow).Value = DR(0)("vendor_id") '2021/04/12 ekawai add S
                    '支払区分自動入力
                    DGV購入品入力("支払区分", iRow).Value = DR(0)("pay_code")

            End Select
        End If
    End Sub

    'Private Sub btnミスミ_Click(sender As Object, e As EventArgs) Handles btnミスミ.Click
    '    Dim DT As DataTable
    '    Dim i As Integer
    '    DT = Read_ミスミCSV("\\Fs1\030購入品管理\技術部\照合済.csv")

    '    i = DGV購入品入力.Rows.Count - 1
    '    For Each Row As DataRow In DT.Rows

    '        DGV購入品入力.Rows(i).Cells("品名").Value = Row("品名").ToString
    '        DGV購入品入力.Rows(i).Cells("型式").Value = Row("型式").ToString

    '        Exit For
    '    Next

    'End Sub

    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        Dim dsOutput As DataSet     '出力リストのいれもの
        'DataSetに取得
        dsOutput = dCon.DataSet(DGV購入品入力SQL, "t_エクセル出力")

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
        Dim DGV行Cnt As Integer = DGV購入品入力.Rows.Count

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
        '2022/07/25 ekawai add S 複数列貼り付け対応
        Dim cols() As String = lines(0).Split(vbTab)
        Dim colMaxIndex As Integer = cols.GetUpperBound(0)
        Dim rowMaxIndex As Integer = lines.GetUpperBound(0) - 1
        Dim strArr(,) As String
        '2022/07/25 ekawai add E



        'テキストボックス内でコピーした時と
        'セルを選択肢してコピーした時で配列の数が変わる
        Dim iLoop As Integer
		'2022/02/14 ekawai add S
        If DGV購入品入力.GetCellCount(DataGridViewElementStates.Selected) > 1 Then
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
        '2022/07/25 ekawai add S
        ReDim strArr(rowMaxIndex, colMaxIndex)
        For rowindex As Integer = 0 To rowMaxIndex
            cols = lines(rowindex).Split(vbTab)
            For colindex As Integer = 0 To colMaxIndex
                strArr(rowindex, colindex) = cols(colindex)
            Next

        Next
        '2022/07/25 ekawai add E
        Dim i貼り付け行数 As Integer
        Dim i貼り付け列数 As Integer

        For i貼り付け行数 = 1 To iLoop
            '2022/07/25 ekawai add S
            For i貼り付け列数 = 1 To colMaxIndex + 1
                '読取専用の行には貼り付けできないようにする
                '2022/11/02 ekawai del S ReadOnlyを使わないように指示があったため
                'If DGV購入品入力(iTargetCol + i貼り付け列数 - 1, iTargetRow + i貼り付け行数 - 1).ReadOnly = True Then
                '    MsgBox("貼り付け範囲内にロックされているセルがあるため貼り付けできません")
                '    Exit Sub
                'End If
                '2022/11/02 ekawai del E
                Select Case DGV購入品入力.Columns(iTargetCol + i貼り付け列数 - 1).HeaderText
                    Case "購入ID", "承認済", "決裁", "予算額", "金額", "入荷ID", "経理検査", "設備名", "仕入先ID", "insertStamp", "作成者", "作成PC名", "作成バージョン", "updateStamp", "更新者", "更新PC名", "更新バージョン", "コピーフラグ"
                        MsgBox("貼り付け範囲内にロックされているセルがあるため貼り付けできません")
                        Exit Sub

                End Select


                If TypeOf DGV購入品入力(iTargetCol + i貼り付け列数 - 1, iTargetRow + i貼り付け行数 - 1) Is DataGridViewComboBoxCell Then
                    MsgBox("貼り付け範囲内にコンボボックスがあるため貼り付けできません")
                    Exit Sub
                End If
            Next
            '2022/07/25 ekawai add E
            If DGV行Cnt - 1 <= iTargetRow + i貼り付け行数 - 1 Then

                Dim dt As DataTable = DGV購入品入力.DataSource
                Dim row As DataRow = dt.NewRow
                row("購入ID") = getNext購入ID()
                dt.Rows.Add(row)
            End If
            ''読取専用の行には貼り付けできないようにする
            'If DGV購入品入力(iTargetCol, iTargetRow + i貼り付け行数 - 1).ReadOnly = True Then
            '    MsgBox("貼り付け範囲内にロックされているセルがあるため貼り付けできません")
            '    Exit Sub
            'End If



        Next
        Dim val As String '202/07/25 ekawai add

        For i貼り付け行数 = 1 To iLoop
            '空白になったら抜ける
            'If r = lines.GetLength(0) - 1 And "".Equals(lines(r)) Then
            '    Exit For
            'End If
            '2022/07/25 ekawai del S
            'Dim val As String = lines(i貼り付け行数 - 1).ToString
            'DGV購入品入力(iTargetCol, iTargetRow + i貼り付け行数 - 1).Value = val
            '仕入先の場合、仕入先IDと科目が自動で選ばれる必要がある
            'DGVテキストボックス変更時処理(iTargetRow + i貼り付け行数 - 1, iTargetCol)
            '2022/07/25 ekawai del E
            '2022/07/25 ekawai add S
            For i貼り付け列数 = 1 To colMaxIndex + 1
                val = strArr(i貼り付け行数 - 1, i貼り付け列数 - 1).ToString
                DGV購入品入力(iTargetCol + i貼り付け列数 - 1, iTargetRow + i貼り付け行数 - 1).Value = val
                '仕入先の場合、仕入先IDと科目が自動で選ばれる必要がある
                DGVテキストボックス変更時処理(iTargetRow + i貼り付け行数 - 1, iTargetCol + i貼り付け列数 - 1)
            Next
            '2022/07/25 ekawai add E
        Next

    End Sub

    Private Sub DGV購入品入力_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV購入品入力.CellClick
        'Dim dt As DataTable = DGV購入品入力.DataSource
        'For Each ROW As DataRow In dt.Rows
        '    Debug.Print("cellclick")
        '    Debug.Print(ROW("品名"))
        '    Debug.Print(ROW("品名", DataRowVersion.Original))
        '    Debug.Print(ROW("品名", DataRowVersion.Current))
        '    Debug.Print(ROW.RowState)
        'Next

        Active行 = e.RowIndex
        Active列 = e.ColumnIndex
    End Sub

    Private Sub btn一括稟議承認_Click(sender As Object, e As EventArgs) Handles btn一括稟議承認.Click
        Dim dt As DataTable = DGV購入品入力.DataSource
        Dim cnt As Integer = 0
        For Each row As DataRow In dt.Rows
            '未承認で絞り込むと否認と保留の場合もある。一括承認は空白行にしか行わない。
            If row("決裁") Is DBNull.Value Then
                row("決裁") = "稟議承認"
                cnt = cnt + 1
            End If
        Next
        MsgBox(cnt & "件に稟議承認を入れました。更新ボタンを押して確定してください")

    End Sub

    Private Sub btn一括承認_Click(sender As Object, e As EventArgs) Handles btn一括承認.Click

        Dim dt As DataTable = DGV購入品入力.DataSource
        Dim cnt As Integer = 0
        '未承認で絞り込むと否認と保留の場合もある。一括承認は空白行にしか行わない。
        For Each row As DataRow In dt.Rows
            If row("決裁") Is DBNull.Value Then
                row("決裁") = Select_Form._氏名
                cnt = cnt + 1
            End If
        Next
        MsgBox(cnt & "件に決裁者名を入れました。更新ボタンを押して確定してください")
    End Sub
    '指定した年月の技術在庫管理システムのTD_gijutsu_nyukaテーブルからTD_gijutsu_shukkaへ出荷データを作成する処理（旧確定コピー）
    Private Sub btn出荷データ作成_Click(sender As Object, e As EventArgs) Handles btn出荷データ作成.Click
        Dim strSql As String
        Dim mCon As New merrweth_init_DbConnection
        Dim dsShukka As DataSet

        '年月指定
        Dim val1 As String = InputBox("データ作成年を入力【西暦4桁】", "作成年は？")
        Dim val2 As String = InputBox("データ作成月を入力【1～12】", "作成月は？")
        Dim date作成年月 As DateTime

        'DateTimeに変換できるかチェック
        If DateTime.TryParse(val1 & "/" & val2 & "/1", date作成年月) = True Then
            '日付変換成功
        Else
            MsgBox("不正な年月です", MsgBoxStyle.OkOnly, "年月変換エラー")
            Exit Sub
        End If

        '実行最終確認
        If MsgBox(val1 & "年" & vbCrLf & val2 & "月" & vbCrLf & "の出荷データを作成します。よろしいですか？", vbYesNo + vbDefaultButton1) = vbYes Then
            'OKボタンで次に進む
        Else
            '処理中止
            Exit Sub
        End If

        Const INSERT_SQL As String = _
            "insert into TD_gijutsu_shukka ( NyukaID, Kubun0, ShukkaNum, DateID, ShukkaSeiriNo )" & _
            " values ("

        strSql = ""
        strSql = "alter view TQ_gijutsu_monthly_shukka"
        strSql = strSql & vbCrLf & "as"
        strSql = strSql & vbCrLf & "SELECT NyukaID, SUM(Num) AS Zaiko"
        strSql = strSql & vbCrLf & "FROM TD_gijutsu_nyuka"
        strSql = strSql & vbCrLf & "GROUP BY NyukaID, LEN(TanaNo), NyukaDate"
        strSql = strSql & vbCrLf & "HAVING (Len(TanaNo) = 0)"
        strSql = strSql & vbCrLf & "And " & MakeCriterion("NyukaDate", Year(date作成年月), Month(date作成年月))
        strSql = strSql & vbCrLf & "Union"
        strSql = strSql & vbCrLf & "SELECT NyukaID, - SUM(ShukkaNum) AS zaiko"
        strSql = strSql & vbCrLf & "FROM TD_gijutsu_shukka"
        strSql = strSql & vbCrLf & "GROUP BY NyukaID, DateID"
        strSql = strSql & vbCrLf & "HAVING " & MakeCriterion("DateID", Year(date作成年月), Month(date作成年月))

        mCon.Command()
        mCon.ExecuteSQL(strSql)

        strSql = ""
        strSql = "SELECT TQ_gijutsu_monthly_shukka.NyukaID, TD_gijutsu_nyuka.Kubun0,"
        strSql = strSql & vbCrLf & "SUM(TQ_gijutsu_monthly_shukka.Zaiko) AS sum_of_zaiko, "
        strSql = strSql & vbCrLf & "TD_gijutsu_nyuka.NyukaDate , TD_gijutsu_nyuka.SeiriNo"
        strSql = strSql & vbCrLf & "FROM TQ_gijutsu_monthly_shukka INNER JOIN TD_gijutsu_nyuka"
        strSql = strSql & vbCrLf & "ON TQ_gijutsu_monthly_shukka.NyukaID = TD_gijutsu_nyuka.NyukaID"
        strSql = strSql & vbCrLf & "GROUP BY TQ_gijutsu_monthly_shukka.NyukaID, TD_gijutsu_nyuka.Kubun0,"
        strSql = strSql & vbCrLf & "TD_gijutsu_nyuka.NyukaDate, TD_gijutsu_nyuka.SeiriNo"
        strSql = strSql & vbCrLf & "HAVING (Sum(TQ_gijutsu_monthly_shukka.Zaiko) > 0)"

        dsShukka = mCon.DataSet(strSql, "t_出荷データ作成対象")
        With dsShukka.Tables("t_出荷データ作成対象")
            For Each dr As DataRow In .Rows
                If dr("sum_of_zaiko") > 0 Then

                    '基本的な処理
                    'zaikoフィールドの数を引き（マイナスの時はプラスに）、在庫を0にする
                    strSql = ""
                    strSql = INSERT_SQL
                    strSql = strSql & vbCrLf & dr("NyukaID")
                    strSql = strSql & vbCrLf & ", '" & dr("Kubun0") & "'"
                    strSql = strSql & vbCrLf & "," & dr("sum_of_zaiko")
                    strSql = strSql & vbCrLf & ", '" & dr("NyukaDate") & "'"
                    strSql = strSql & vbCrLf & ", '" & dr("SeiriNo") & "'"
                    strSql = strSql & vbCrLf & ")"
                    '作成されたSQLに基づいて出荷データを挿入していく
                    mCon.ExecuteSQL(strSql)

                    strSql = ""
                    strSql = "update TD_gijutsu_nyuka set erase_inventory = 1"
                    strSql = strSql & vbCrLf & "where NyukaID =" & dr("NyukaID")
                    mCon.ExecuteSQL(strSql)

                End If
            Next
            '終了通知
            MsgBox(.Rows.Count & "件の出荷データを作成しました", MsgBoxStyle.OkOnly, "通知")
        End With
    End Sub
    '引数説明 FieldName：条件にするフィールド名、mintYear：出荷データ作成年、mintMonth：出荷データ作成月
    Private Function MakeCriterion(ByVal FieldName As String, ByVal mintYear As Integer, ByVal mintMonth As Integer) As String
        MakeCriterion = " " & FieldName & " >= '" & mintYear & "/" & mintMonth & "/1'"
        If mintMonth = 12 Then
            MakeCriterion = MakeCriterion & " and " & FieldName & " < '" & mintYear + 1 & "/1/1'"
        Else
            MakeCriterion = MakeCriterion & " and " & FieldName & " < '" & mintYear & "/" & mintMonth + 1 & "/1'"
        End If
    End Function
    'KeyDownイベントはどのキー（例えば［Delete］キー）を押し下げても発生する。
    'イベントの発生順序は、KeyDownイベント→KeyPressイベント→KeyUpイベントの順となる。
    'KeyDown使わずKeyPressだけにすることにした
    'Private Sub txt検索条件1_KeyDown(sender As Object, e As KeyEventArgs) Handles txt検索条件1.KeyDown
    '    If e.KeyCode = Keys.Enter Then

    '        If txt検索条件1.Text <> "" Then
    '            Disp検索結果()
    '        End If
    '    End If

    'End Sub


    'Private Sub txt検索条件2_KeyDown(sender As Object, e As KeyEventArgs) Handles txt検索条件2.KeyDown
    '    If e.KeyCode = Keys.Enter Then

    '        Disp検索結果()
    '    End If

    'End Sub

    'Private Sub txt検索条件3_KeyDown(sender As Object, e As KeyEventArgs) Handles txt検索条件3.KeyDown
    '    If e.KeyCode = Keys.Enter Then

    '        Disp検索結果()
    '    End If

    'End Sub



    'Private Sub DGV購入品入力_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DGV購入品入力.CellValueChanged
    '    Dim dt As DataTable = DGV購入品入力.DataSource
    '    For Each ROW As DataRow In dt.Rows
    '        Debug.Print("cellvaluechanged")
    '        Debug.Print(ROW("品名"))
    '        Debug.Print(ROW("品名", DataRowVersion.Original))
    '        Debug.Print(ROW("品名", DataRowVersion.Current))
    '        Debug.Print(ROW.RowState)
    '    Next

    'End Sub


    Private Sub btnManual_Click(sender As Object, e As EventArgs) Handles btnManual.Click

        Dim xlApplication As New Excel.Application()
        Dim xlBooks As Excel.Workbooks

        ' xlApplication から WorkBooks を取得する
        xlBooks = xlApplication.Workbooks

        ' 既存の Excel ブックを開く
        xlBooks.Open("\\Fs1\040システム\Manual\購入依頼.xlsx")

        ' Excel を表示する
        xlApplication.Visible = True


    End Sub


    Private Sub DGV購入品入力_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles DGV購入品入力.RowsAdded
        '新規入力で行が増える時だけでなく、DGVにDTをバインドするたびに発生する
        '再表示の時は自動計算発生させたくないからFlg再表示で制御することになった
        If Flg再表示 = False Then
            '製品からコピーした時に予算額が入らない問題の対策
            '製品からコピーした時は行が増えるから必ずRowAddedイベントを通るので、自動計算を行う
            自動計算(e.RowIndex, True)
            自動計算(e.RowIndex, False)
        End If


    End Sub
    'メルウェスの製品計画ビューアーのやり方真似したが、購入依頼のほうが圧倒的に行数も列数も多いため複数行選択や全セル選択すると固まってしまう
    'そこで、CurrentCellが指定列の時だけFor Eachを通るようにして負荷を抑えることにした
    'CurrentCellsの列は、行選択と左上で全セル選択した時は「購入ID」になり、マウスやキーボードで複数セル選択した時は、最後に選択したセルの列になることが分かった
    'つまり、この制御の仕方では最後のセルの位置しか判定できないから、金額とか数量になるように計算対象外のセルを選択することはできてしまう。例えば、仕入先ID+金額の計算は可能(仕入先IDが数値だから)
    'そこまでは制御できないと分かった上で、IF文の分岐行うことに決まった。(フォームの大きさ限られているから、マウスやキーボードで複数範囲選択出来る量って限られているから不要なセル大量に選択されても、そこまで支障ないという判断だった)
    Private Sub DGV購入品入力_SelectionChanged(sender As Object, e As EventArgs) Handles DGV購入品入力.SelectionChanged
        Dim W_total As Double = 0.0
        Dim sCurrentHeader As String
        sCurrentHeader = DGV購入品入力.Columns(DGV購入品入力.CurrentCell.ColumnIndex).HeaderText
        'Debug.Print(sCurrentHeader)
        If sCurrentHeader = "数量" Or sCurrentHeader = "予算単価" Or sCurrentHeader = "税抜単価" Or sCurrentHeader = "予算額" Or sCurrentHeader = "金額" Then
            '選択セルをループ
            For Each c As DataGridViewCell In DGV購入品入力.SelectedCells
                Dim W_Suryo As Double
                Try
                    If Double.TryParse(CStr(c.Value), W_Suryo) Then
                        W_total += W_Suryo

                    End If
                    '文字列だったらTryParseでエラーが起きて計算しないようにしている(メルウェスの生産計画ビューアー参照)
                Catch ex As Exception

                End Try
            Next

            lblセル合計.Text = W_total
        Else
            lblセル合計.Text = ""
        End If
        'Debug.Print(Now)
    End Sub

    'SelectionChangedだとイベントが発生しすぎるのではという心配からボタンを作ってみたが、自動で値がクリアされないので誤解を招くということで
    '結局SelectionChanged使うことになった。
    'Private Sub btnSUM_Click(sender As Object, e As EventArgs)
    '    Dim W_total As Double = 0.0
    '    '行選択による計算を制御するために列名指定
    '    If DGV購入品入力.Columns(DGV購入品入力.CurrentCell.ColumnIndex).HeaderText = "金額" Then
    '        '選択セルをループ
    '        For Each c As DataGridViewCell In DGV購入品入力.SelectedCells
    '            Dim W_Suryo As Double
    '            Try
    '                If Double.TryParse(CStr(c.Value), W_Suryo) Then
    '                    W_total += W_Suryo

    '                End If
    '            Catch ex As Exception

    '            End Try
    '        Next

    '        lblセル合計.Text = W_total
    '    End If
    'End Sub

    Private Sub btn検索プログラム_Click(sender As Object, e As EventArgs) Handles btn検索プログラム.Click
        Process.Start("\\Fs1\d14it\プログラム\経理\購入依頼検索プログラムⅡ\bin\購入依頼検索プログラムⅡ.exe")

    End Sub

    Private Sub btnスクロール_Click(sender As Object, e As EventArgs) Handles btnスクロール.Click
        Try
            If strScroll = "Left" Then
                DGV購入品入力.FirstDisplayedScrollingColumnIndex = DGV購入品入力.Columns("品名").Index
                strScroll = "Right"
            Else
                DGV購入品入力.FirstDisplayedScrollingColumnIndex = DGV購入品入力.Columns("購入ID").Index
                strScroll = "Left"
            End If
        Catch ex As Exception
            Exit Sub
        End Try

    End Sub

    '型式のみヘッダクリックで列幅が変わるようにする
    '半角カタカナ英数字記号が連続している場合は折返しがされない仕様なので、少しでも型式見やすくするために機能追加。(日本語やスペースが入っていたら折返しで対応できるので、この機能必要なのは技術部くらい)
    Private Sub DGV購入品入力_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DGV購入品入力.ColumnHeaderMouseClick
        If e.ColumnIndex = DGV購入品入力.Columns("型式").Index Then
            Try
                If DGV購入品入力.Columns("型式").Width = 120 Then
                    DGV購入品入力.Columns("型式").Width = 320
                Else
                    DGV購入品入力.Columns("型式").Width = 120
                End If
            Catch ex As Exception
                Exit Sub
            End Try
        End If
    End Sub

    Private Sub cmb検索条件選択_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb検索条件選択1.SelectedIndexChanged, cmb検索条件選択2.SelectedIndexChanged, cmb検索条件選択3.SelectedIndexChanged
        Dim cmb項目 As ComboBox = CType(sender, ComboBox)
        Dim i As Integer
        i = Strings.Right(cmb項目.Name, 1)

        CType(Me.Controls("txt検索条件" & i), TextBox).Text = CType(Me.Controls("cmb検索条件選択" & i), ComboBox).SelectedItem.ToString


    End Sub

    'このイベントの中で分岐しても、結局コンボのエラー以外の表示したいメッセージも出なくなってしまったので使うのやめた。

    'Private Sub DGV購入品入力_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DGV購入品入力.DataError
    '    '特定のパソコンで、テキストボックスにカーソルがある状態でコンボボックスをクリックすると
    '    'SystemFormatException:セルのフォーマットされた値に間違った型が指定されていますというエラーが出る対策

    '    If TypeOf DGV購入品入力.Columns(e.ColumnIndex) Is DataGridViewComboBoxColumn Then
    '        e.Cancel = True
    '    End If

    'End Sub
End Class
