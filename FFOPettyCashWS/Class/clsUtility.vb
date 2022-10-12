Imports System.Security.Cryptography
Imports Newtonsoft.Json
Imports System.IO

Public Class clsUtility

#Region "Method - JSON"
    Public Shared Function ConvertDatatableToJson(ByVal dt As DataTable) As String
        Dim jsonstring As String = ""
        jsonstring = JsonConvert.SerializeObject(dt)
        Return jsonstring
    End Function
    Public Shared Function ConvertJsonToDatatable(ByVal str_json As String) As DataTable
        Dim ds As New DataSet
        Dim dt As New DataTable
        Try
            ds = JsonConvert.DeserializeObject(Of DataSet)(str_json)
            dt = ds.Tables(0)
        Catch ex As Exception
            Return Nothing
        End Try
        Return dt
    End Function
    Public Shared Function ConvertJsonToDataset(ByVal str_json As String) As DataSet
        Dim ds As New DataSet
        Try
            ds = JsonConvert.DeserializeObject(Of DataSet)(str_json)
        Catch ex As Exception
            Return Nothing
        End Try
        Return ds
    End Function
    Public Shared Function ConvertStringToBytes(ByVal BytesToString As String) As Byte()
        Try
            Return Convert.FromBase64String(BytesToString)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
#End Region

#Region "Method - Token"
    Public Shared Function GenerateToken(ByVal employeeid As Integer, ByVal currentdatetime As DateTime) As String
        Dim rtoken As String = ""
        Dim toencrypt As String = ""
       
        toencrypt = employeeid.ToString.Trim & "|" & currentdatetime.ToString("ddMMyyyy") + "|" + currentdatetime.ToString("HHmm")
        rtoken = clsJCypher.StringCipher.Encrypt(toencrypt, "")

        Return rtoken
    End Function
    Public Shared Function ValidateToken(ByVal rtoken As String) As String()
        Dim isValid As Boolean = True
        Dim todescrypt As String = ""
        Dim strlist() As String = Nothing
        Dim employeeid As Integer
        Dim submitteddatetime As DateTime
        Dim searchstring As String = ""
        Dim errmsg As String = ""
        Try
            todescrypt = clsJCypher.StringCipher.Decrypt(rtoken, "")
            strlist = todescrypt.Split("|")
            If (strlist.Length > 3) Then
                searchstring = strlist(3)
            End If

            If strlist.Length > 0 Then
                Dim d As Date = Date.ParseExact(strlist(1), "ddMMyyyy", Nothing)
                Dim t As DateTime = Date.ParseExact(strlist(2), "HHmm", Nothing)
                Dim converteddate As String = ""

                employeeid = strlist(0)
                converteddate = d.ToString("dd/MM/yyyy") & " " & t.ToString("HH:mm:ss")
                submitteddatetime = Date.ParseExact(converteddate, "dd/MM/yyyy HH:mm:ss", Nothing)

                Dim tspan As Integer = DateDiff(DateInterval.Minute, submitteddatetime, clsData.getServerDate())
                If tspan > 2 Then   'if time span is greater than 2 minutes, deny request = invalid token
                    isValid = False
                    errmsg = "Error: Invalid Token " & rtoken
                    GoTo tonton
                End If
            End If

            If VerifyEmployeeRequest(employeeid) = "cancel" Then
                isValid = False     '401: Unauthorized Access.
                errmsg = "401: Unauthorized Access " & employeeid.ToString
                GoTo tonton
            End If
        Catch ex As Exception
            errmsg = ex.Message
            isValid = False
        End Try
tonton:

        Return New String() {employeeid, isValid, errmsg, searchstring}
    End Function
    Private Shared Function VerifyEmployeeRequest(ByVal employeeid As Integer) As String
        Dim dt As New DataTable
        Dim result As String = "cancel"
        dt = clsData.getGenericData(0, employeeid, "")
        If dt.Rows.Count > 0 Then
            If dt.Rows(0).Item("cnt") > 0 Then
                result = "allow"
            Else
                result = "cancel"
            End If
        End If
        Return result
    End Function
#End Region

#Region "Method - Logs"
    Public Shared Sub SaveLog(ByVal refno As String, ByVal strmodule As String, ByVal username As Integer, ByVal action As String)
        Dim dt As New DataTable
        Dim logdate As New DateTime
        logdate = clsData.getServerDate
        clsData.SaveLog(refno, strmodule, logdate, username, action)
    End Sub
#End Region

#Region "Method - Decryption"
    Public Function Decrypt(ByVal constring As String) As String
        Dim KEY_128 As Byte() = {42, 1, 52, 67, 231, 13, 94, 101, 123, 6, 0, 12, 32, 91, 4, 111, 31, 70, 21, 141, 123, 142, 234, 82, 95, 129, 187, 162, 12, 55, 98, 23}
        Dim IV_128 As Byte() = {234, 12, 52, 44, 214, 222, 200, 109, 2, 98, 45, 76, 88, 53, 23, 78}
        Dim symmetricKey As RijndaelManaged = New RijndaelManaged()
        symmetricKey.Mode = CipherMode.CBC

        Dim enc As System.Text.UTF8Encoding
        Dim decryptor As ICryptoTransform

        enc = New System.Text.UTF8Encoding
        decryptor = symmetricKey.CreateDecryptor(KEY_128, IV_128)

        Dim cypherTextBytes As Byte() = Convert.FromBase64String(constring)
        Dim memoryStream As MemoryStream = New MemoryStream(cypherTextBytes)
        Dim cryptoStream As CryptoStream = New CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read)
        Dim plainTextBytes(cypherTextBytes.Length) As Byte
        Dim decryptedByteCount As Integer = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length)
        memoryStream.Close()
        cryptoStream.Close()
        Return enc.GetString(plainTextBytes, 0, decryptedByteCount)
    End Function
#End Region

   
End Class
