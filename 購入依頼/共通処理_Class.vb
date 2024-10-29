Public Class 共通処理_Class

    '----------------------------------------------------------------------------------------------------------
    'ＳＱ　…'で囲むときに使う
    '----------------------------------------------------------------------------------------------------------
    '文字に''を付けたい時に使う
    Public Function SQ(str As String) As String
        Return "'" & str & "'"
    End Function

    '----------------------------------------------------------------------------------------------------------
    'ＮＺ　…数値型のセルがNullの時に0を返す
    '----------------------------------------------------------------------------------------------------------
    Public Function nz(ByVal inData)
        Dim t

        If inData Is System.DBNull.Value Then
            Return 0
        End If

        If inData Is Nothing Then
            Return 0
        End If

        If inData.Equals(vbNull) Then
            t = 0
        Else
            If IsDBNull(inData) Then
                t = 0
            Else
                If IsNumeric(inData) = False Then
                    t = 0
                Else
                    t = inData
                End If
            End If
        End If
        Return t
    End Function

    '----------------------------------------------------------------------------------------------------------
    'ＮＳ　…文字列型のセルがNullの時に空白を返す
    '----------------------------------------------------------------------------------------------------------
    Public Function ns(ByVal inData)
        Dim t As String
        'If inData Is System.DBNull.Value Then
        '    Return ""
        'End If

        If inData Is Nothing Then
            Return ""
        End If

        If inData Is System.DBNull.Value Then
            Return ""
        End If

        If inData.Equals(vbNull) Then
            t = ""
        Else
            If IsDBNull(inData) Then
                t = ""
            Else
                t = inData
            End If
        End If
        Return t
    End Function
    '----------------------------------------------------------------------------------------------------------
    'ＮＮ　…セルがNullの時にnullを返す。Nullでなければ'で囲んだ状態で文字を返す UPDATEやINSERT時、文字と数値どちらでも使える
    '----------------------------------------------------------------------------------------------------------
    '文字列用
    'Nullや空白の時は、nullを返すファンクション(必須項目以外のINSERT時に使う)
    Public Function nn(ByVal inData)
        Dim t As String

        If inData Is Nothing Then
            Return "null"
        End If

        If inData Is System.DBNull.Value Then
            Return "null"
        End If

        If inData.Equals(vbNull) Then
            t = "null"
        Else
            If IsDBNull(inData) Then
                t = "null"
            Else
                t = SQ(inData)
            End If
        End If
        Return t
    End Function
End Class
