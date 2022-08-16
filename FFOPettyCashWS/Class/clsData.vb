

Public Class clsData

    Public mainStrPar As New List(Of String)
    Public mainStrVal As New List(Of String)
#Region "Method - Query"

    Public Shared Function getGenericData(ByVal sop As Integer, ByVal employeeid As Integer, ByVal code As String) As DataTable
        Dim str_par() As String = {"operation", "soperation", "employeeid", "code"}
        Dim str_val() As Object = {2, sop, employeeid, code}
        Return clsSP.SP_Fetch_Query("spFFOSyncPC", str_par, str_val)
    End Function

#Region "Method - Query - Download"
    Public Function DownloadData() As DataTable
        Return clsSP.SP_Fetch_Query("spFFOSyncPC", mainStrPar.ToArray, mainStrVal.ToArray)
    End Function
#End Region

#Region "Method - Query - Upload"
    Public Shared Function UploadFundReleased(ByVal employeeid As Integer, ByVal dt As DataTable) As String
        Dim str_par() As String = {"operation", "soperation", "employeeid", "fundreleased"}
        Dim str_val() As Object = {1, 0, employeeid, dt}
        Return clsSP.SP_Transact_Query("spFFOSyncPC", str_par, str_val, True)
    End Function
    Public Shared Function UploadFundLiquidation(ByVal employeeid As Integer, ByVal dt As DataTable) As String
        Dim str_par() As String = {"operation", "soperation", "employeeid", "fundliquidation"}
        Dim str_val() As Object = {1, 1, employeeid, dt}
        Return clsSP.SP_Transact_Query("spFFOSyncPC", str_par, str_val, True)
    End Function
    Public Shared Function UploadVendor(ByVal employeeid As Integer, ByVal dt As DataTable) As String
        Dim str_par() As String = {"operation", "soperation", "employeeid", "vendor"}
        Dim str_val() As Object = {1, 2, employeeid, dt}
        Return clsSP.SP_Transact_Query("spFFOSyncPC", str_par, str_val, True)
    End Function
    Public Shared Function UploadExpenseType(ByVal employeeid As Integer, ByVal dt As DataTable) As String
        Dim str_par() As String = {"operation", "soperation", "employeeid", "expenses"}
        Dim str_val() As Object = {1, 3, employeeid, dt}
        Return clsSP.SP_Transact_Query("spFFOSyncPC", str_par, str_val, True)
    End Function
    Public Shared Function UploadExpenseReport(ByVal employeeid As Integer, ByVal dt As DataTable) As String
        Dim str_par() As String = {"operation", "soperation", "employeeid", "cashdisbursement"}
        Dim str_val() As Object = {1, 4, employeeid, dt}
        Return clsSP.SP_Transact_Query("spFFOSyncPC", str_par, str_val, True)
    End Function
    Public Shared Function UploadFundReleasedCashTransfer(ByVal employeeid As Integer, ByVal dt As DataTable) As String
        Dim str_par() As String = {"operation", "soperation", "employeeid", "cashflow"}
        Dim str_val() As Object = {1, 5, employeeid, dt}
        Return clsSP.SP_Transact_Query("spFFOSyncPC", str_par, str_val, True)
    End Function
#End Region

#Region "Method - Query - Others"
    Public Shared Sub SaveLog(ByVal refno As String, ByVal strmodule As String, ByVal logdate As DateTime, ByVal username As Integer, ByVal action As String)
        Dim str_par() As String = {"operation", "soperation", "referenceno", "module", "logdate", "username", "action"}
        Dim str_val() As Object = {1, 0, refno, strmodule, logdate, username, action}
        clsSP.SP_Transact_Query("spUserLog", str_par, str_val, False)
    End Sub

    Public Shared Function getServerDate() As DateTime
        Dim str_par() As String = {"operation"}
        Dim str_val() As Object = {1}
        Dim dt As DataTable = clsSP.SP_Fetch_Query("sp_GetServerDateTime", str_par, str_val)
        Dim vServerDate As Date = CDate(dt.Rows(0).Item("serverdatetime").ToString)
        dt = Nothing
        Return vServerDate
    End Function
#End Region

#End Region



End Class
