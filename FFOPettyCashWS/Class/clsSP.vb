'Tonton 2/17/2022
'Functions to connect to database

Imports System.Data.SqlClient
Public Class clsSP
    Public Property my_connection_string As String

    Public Shared Function SP_Fetch_Query(ByRef sp As String, ByRef par_name As String(), ByRef par_val As Object()) As Object
        Dim cls As New clsConnection
        Dim sql_con As New SqlConnection(cls.getConnectionString())
        Dim sql_cmd As New SqlCommand
        Dim sql_adapt As SqlDataAdapter
        Dim ds As New DataSet
        Try
            sql_cmd.Connection = sql_con
            sql_cmd.CommandText = sp
            sql_cmd.CommandType = CommandType.StoredProcedure

            For i As Integer = 0 To par_name.Count - 1
                If Not IsNothing(par_val(i)) Then
                    sql_cmd.Parameters.AddWithValue(par_name(i), par_val(i))
                End If
            Next
            sql_adapt = New SqlDataAdapter(sql_cmd)
            sql_adapt.Fill(ds)

            sql_con.Close()
            sql_con.Dispose()
            Return ds.Tables(0)
        Catch ex As Exception
            sql_con.Close()
            sql_con.Dispose()
            Return Nothing
        End Try
    End Function
    Public Shared Function SP_Transact_Query(ByRef sp As String, ByRef par_name As String(), ByRef par_val As Object(), ByRef is_return_rec As Boolean) As Object
        Dim cls As New clsConnection
        Dim sql_con As New SqlConnection(cls.getConnectionString())
        Dim sql_cmd As New SqlCommand
        Dim str_return_val As Object = Nothing
        Try
            sql_cmd.Connection = sql_con
            sql_cmd.CommandText = sp
            sql_cmd.CommandType = CommandType.StoredProcedure

            For i As Integer = 0 To par_name.Count - 1
                If Not IsNothing(par_name(i)) Then
                    sql_cmd.Parameters.AddWithValue(par_name(i), par_val(i))
                End If
            Next

            If is_return_rec = True Then
                sql_cmd.Parameters.Add("@NewPK", SqlDbType.VarChar, 5000).Direction = ParameterDirection.Output
            End If

            sql_con.Open()
            str_return_val = sql_cmd.ExecuteNonQuery()

            If is_return_rec = True Then   'false
                str_return_val = sql_cmd.Parameters("@NewPK").Value
            End If

            sql_con.Close()
            sql_con.Dispose()
            Return str_return_val
        Catch ex As Exception
            str_return_val = "ERROR: " & ex.Message
            sql_con.Close()
            sql_con.Dispose()
            Return str_return_val
        End Try
    End Function

    Public Shared Function randomkey(Optional ByVal char_limit As Integer = 10) As String
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim str_characters() As String = "1,2,3,4,5,6,7,8,9,0,a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z".Split(",")
        Dim final_output As String = ""
        dt.Columns.Add("id", GetType(Integer))
        dt.Columns.Add("guid", GetType(String))

        For i As Integer = 0 To str_characters.Length - 1
            dr = dt.NewRow
            dr("id") = i
            dr("guid") = Guid.NewGuid().ToString
            dt.Rows.Add(dr)
        Next
        dt.DefaultView.Sort = "guid ASC"
        dt = dt.DefaultView.ToTable
        For x As Integer = 0 To dt.Rows.Count - 1
            final_output += str_characters(dt.Rows(x).Item(0))
        Next
        'Debug.Print(final_output)
        Return final_output.Substring(0, char_limit)
    End Function


    
End Class
