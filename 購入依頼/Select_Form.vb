﻿Public Class Select_Form
    Private dCon As New merrweth_init_DbConnection
    Public _UserName As String = System.Environment.UserName
    Public _社員番号 As Integer = vbEmpty
    Public _PC名 As String = System.Environment.MachineName
    Public _氏名 As String
    Public s部署 As String = ""
    Public i部署ID As Integer
    Public 更新Flg As Boolean
    Public 承認Flg As Boolean
    Public 経理Flg As Boolean
    Public _MODE As String '表示モード
    Public i画面高さ As Integer
    Public i画面幅 As Integer

    Private Sub Select_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim DT部署 As DataTable
        Dim strSql As String = ""
        Dim s対象部署 As String
        Dim Dsp As System.Windows.Forms.Screen = System.Windows.Forms.Screen.FromControl(Me)
        lblcatalog.Text = "接続先:" & merrweth_init_DbConnection.strCatalog
        Me.Text = "部署選択"
        strSql &= vbCrLf & "SELECT *"
        strSql &= vbCrLf & "FROM TM_dep"
        strSql &= vbCrLf & "WHERE DepName <> '退職者' AND 廃止 <>1 AND DepName <> '役員'"
        DT部署 = dCon.DataSet(strSql, "DT").Tables(0)

        Call SetListBoxDataSource(lst部署, DT部署, "DepName", "DepID")
        s対象部署 = Get部署()
        Get氏名()
        lst部署.ClearSelected()
        If Get部署() = "未選択" Then
            '選択しない
        Else
            lst部署.Text = s対象部署
        End If
        i画面高さ = Dsp.Bounds.Height
        i画面幅 = Dsp.Bounds.Width
    End Sub
    Public Sub SetListBoxDataSource(ListBox As ListBox, DT As DataTable, DisplayName As String, ValueName As String)
        'combo,DataTable,DataTable上の表示項目名,DataTable上の値項目名からコンボボックスの準備
        Dim comboDT As New DataTable()

        Dim DTCol As DataColumn
        DTCol = DT.Columns(DisplayName)
        comboDT.Columns.Add("Display", Type.GetType(DTCol.DataType.ToString))   '表示値
        DTCol = DT.Columns(ValueName)
        comboDT.Columns.Add("Value", Type.GetType(DTCol.DataType.ToString))    '選択値

        For Each DTRow As DataRow In DT.Rows
            Dim newrow As DataRow = comboDT.NewRow()
            newrow("Display") = DTRow(DisplayName)
            newrow("Value") = DTRow(ValueName)
            comboDT.Rows.Add(newrow)
        Next
        ListBox.DataSource = comboDT
        ListBox.ValueMember = "Value"
        ListBox.DisplayMember = "Display"
        'ListBox.SelectedIndex = 0
    End Sub
    Public Function Get部署() As String

        _社員番号 = GetUserId(_UserName)
        If _社員番号 < 0 Then
            MsgBox("TM_employeeに登録されていません:" & _UserName)
            Return False
        End If


        '部署取得
        Dim strSql As String
        strSql = ""
        strSql = strSql & vbCrLf & "SELECT TM_dep.DepName"
        strSql = strSql & vbCrLf & "FROM TM_employee INNER JOIN"
        strSql = strSql & vbCrLf & "TM_group ON TM_employee.group_id = TM_group.id INNER JOIN"
        strSql = strSql & vbCrLf & "TM_dep ON TM_group.department_id = TM_dep.DepID"
        strSql = strSql & vbCrLf & "WHERE TM_employee.id = " & _社員番号

        Dim dtSet = dCon.DataSet(strSql, "DT")
        If dtSet.Tables(0).Rows.Count = 0 Then
            MsgBox(_社員番号 & "は社員マスタに登録されていません")
            Get部署 = "未選択"
        Else
            Dim dtRow = dtSet.Tables(0).Rows(0)
            If dtRow("DepName") = "役員" Then
                '役員だったら部署の自動選択はなし
                Get部署 = "未選択"
            Else
                Get部署 = dtRow("DepName")
            End If
        End If
    End Function
    Public Function GetUserId(strUserName As String) As Integer
        Dim strSql As String
        strSql = "select "
        strSql = strSql & vbCrLf & " * "
        strSql = strSql & vbCrLf & "from TM_employee"
        strSql = strSql & vbCrLf & "where user_name ='" & _UserName & "'"
        Dim dtSet = dCon.DataSet(strSql, "DT")
        If dtSet.Tables(0).Rows.Count = 0 Then
            GetUserId = -1
        Else
            Dim dtRow = dtSet.Tables(0).Rows(0)
            GetUserId = dtRow("id")
        End If

    End Function
    Private Sub btn部署選択_Click(sender As Object, e As EventArgs) Handles btn部署選択.Click
        Check_権限()
        _MODE = "全項目"
        Main_Form.ShowDialog()
        Main_Form.Close()
        Main_Form.Dispose()
    End Sub
    Sub Get氏名()
        Dim strSql As String
        strSql = ""
        strSql = "SELECT employee"
        strSql &= vbCrLf & "FROM TM_employee"
        strSql &= vbCrLf & "WHERE user_name = '" & _UserName & "'"
        Dim DT氏名 As DataTable = dCon.DataSet(strSql, "DT").Tables(0)
        _氏名 = DT氏名.Rows(0)("employee")
    End Sub
    Private Sub Check_権限()
        Dim strSql As String
        '権限リセット
        更新Flg = False
        承認Flg = False
        経理Flg = False
        s部署 = lst部署.Text '選択中の部署IDを変数に渡す
        i部署ID = lst部署.SelectedValue
        strSql = ""
        strSql = "SELECT * FROM TM_購入依頼権限"
        strSql &= vbCrLf & "WHERE 部署ID = " & i部署ID
        strSql &= vbCrLf & "AND 社員番号 = " & _社員番号
        Dim DT権限 As DataTable = dCon.DataSet(strSql, "DT").Tables(0)
        If DT権限.Rows.Count > 0 Then
            If DT権限.Rows(0)("更新") = True Then
                更新Flg = True
            End If
            If DT権限.Rows(0)("承認") = True Then
                承認Flg = True
            End If
            If DT権限.Rows(0)("経理") = True Then
                経理Flg = True
            End If
        Else
            更新Flg = False
            承認Flg = False
            経理Flg = False
        End If

    End Sub
    Private Sub btn決裁_Click(sender As Object, e As EventArgs) Handles btn決裁.Click
        Check_権限()
        _MODE = "決裁"
        Main_Form.ShowDialog()
        Main_Form.Close()
        Main_Form.Dispose()

    End Sub

End Class