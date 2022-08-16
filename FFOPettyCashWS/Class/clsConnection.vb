Public Class clsConnection
#Region "Method - Connection String"
    Public Function getConnectionString() As String
        Dim connstring As String

        Dim gDataSource As String = ConfigurationManager.AppSettings("gDataSource")
        Dim gDatabasename As String = ConfigurationManager.AppSettings("gDatabasename")
        Dim gUsername As String = ConfigurationManager.AppSettings("gUsername")
        Dim gPassword As String = (New clsUtility).Decrypt(ConfigurationManager.AppSettings("gPassword"))
        Dim gAuthType As String = ConfigurationManager.AppSettings("gAuthType")

        connstring = "Data Source=" & gDataSource & ";Initial Catalog=" & gDatabasename & "; User ID=" & gUsername & "; Password=" & gPassword & ";"

        Return connstring
    End Function
#End Region
End Class
