Public Class Master_Form
    Private dCon As New merrweth_init_DbConnection

    Sub マスタ_read()
        Dim strSql As String
        Dim DS As DataSet
        DGV_マスタ.DataSource = Nothing
        strSql = ""
        Select Case cmbマスタ種類.Text
            Case "仕入先"
                DGV_マスタ.DataSource = Main_Form.DT部署別仕入先

            Case "科目"
                DGV_マスタ.DataSource = Main_Form.DT現科目

            Case "ワークコード"
                DGV_マスタ.DataSource = Main_Form.DT現ワークコード

        End Select

        DGV_マスタ.ReadOnly = True
    End Sub

    Private Sub Master_Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        cmbマスタ種類.Items.Add("仕入先")
        cmbマスタ種類.Items.Add("科目")
        cmbマスタ種類.Items.Add("ワークコード")
    End Sub

    Private Sub cmbマスタ種類_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbマスタ種類.SelectedIndexChanged
        マスタ_read()
    End Sub
End Class