Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports Newtonsoft.Json
Imports System.IO

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class Service1
    Inherits System.Web.Services.WebService

    Enum myTransactCode
        CGetEmployees = 11
        CGetExpenseType = 12
        CGetFundType = 13
        CGetVendor = 14
        CGetBranch = 15
        CGetVendorCategory = 16
        CGetHBItems = 17
        CGetOHBItems = 18


        CPostFundReleased = 21
        CPostExpenseLiquidation = 22
        'CPostFundLiquidation = 22
        CPostExpenseType = 23
        CPostVendor = 24
        'CPostFundReleasedCashTransfer = 26
    End Enum
    Enum ResponseCode
        success = 200
        fail = 404
    End Enum

    <WebMethod()> _
    Public Function Generate_Token(ByVal employeeid As Long) As String
        Dim convertedvalue As String = ""
        Try
            convertedvalue = clsUtility.GenerateToken(employeeid, clsData.getServerDate())
        Catch ex As Exception
            convertedvalue = ex.Message
        End Try
        Return convertedvalue
    End Function

    <WebMethod()> _
    Public Function Download_Data(ByVal rtoken As String, ByVal code As myTransactCode, ByVal branchid As String) As String
        Dim dt As New DataTable
        Dim convertedvalue As String = ""
        Dim traceline As String = ""
        Dim moddescription As String = ""
        Dim employeeid, operation, soperation As Integer
        Dim arr As Array

        Try
            arr = clsUtility.ValidateToken(rtoken)    'validate token
            employeeid = arr(0).ToString 'employeeid
            If arr(1).ToString = False Then  'is valid
                convertedvalue = arr(2).ToString 'error message
                GoTo tonton
            End If

            traceline = "1"
            Dim dtcode As New DataTable
            dtcode = clsData.getGenericData(1, 0, code)
            If dtcode.Rows.Count > 0 Then
                moddescription = dtcode.Rows(0).Item("module")
                operation = dtcode.Rows(0).Item("operation")
                soperation = dtcode.Rows(0).Item("soperation")
            Else
                convertedvalue = "ERROR: Incorrect code."
                GoTo tonton
            End If

            Dim r As New clsData
            r.mainStrPar.Add("operation")       '0
            r.mainStrVal.Add(operation)         '0
            r.mainStrPar.Add("soperation")      '1
            r.mainStrVal.Add(soperation)        '1
            r.mainStrPar.Add("employeeid")      '2
            r.mainStrVal.Add(employeeid)        '2
            r.mainStrPar.Add("branchid")        '3
            r.mainStrVal.Add(branchid)          '3

            dt = r.DownloadData

            traceline = "0"
            convertedvalue = clsUtility.ConvertDatatableToJson(dt)

        Catch ex As Exception
            convertedvalue = "ERROR" & traceline & ": " & ex.Message
        End Try

tonton:
        clsUtility.SaveLog(code, moddescription, employeeid, convertedvalue)

        Return convertedvalue
    End Function

    <WebMethod()> _
    Public Function Upload_Data(ByVal rtoken As String, ByVal code As myTransactCode, ByVal jsondata As String) As Models.Response
        Dim dtconvert As New DataTable
        Dim dtwrap As New DataTable
        Dim convertedvalue = "", retvalarry As String = "", errval As String = "", errmsg As String = ""
        Dim traceline As String = ""
        Dim moddescription As String = ""
        Dim employeeid As Integer
        Dim arr As Array
        Try
            arr = clsUtility.ValidateToken(rtoken)    'validate token
            employeeid = arr(0).ToString 'employeeid
            If arr(1).ToString = False Then  'is valid
                errval = "404"
                errmsg = arr(2).ToString 'error message
                GoTo tonton
            End If
            traceline = "1"
            dtconvert = clsUtility.ConvertJsonToDatatable(jsondata)

            If dtconvert.Rows.Count > 0 Then
                traceline = "2"
                dtwrap = WrapDataTable(dtconvert, code)
                If dtwrap.Rows.Count > 0 Then
                    If code = myTransactCode.CPostFundReleased Then 'post_fundreleased
                        traceline = "3"
                        retvalarry = clsData.UploadFundReleased(employeeid, dtwrap)
                        moddescription = "Fund Released"

                    ElseIf code = myTransactCode.CPostExpenseLiquidation Then    'post_expenseliquidation
                        traceline = "7"
                        retvalarry = clsData.UploadExpenseLiquidation(employeeid, dtwrap)
                        moddescription = "Expense Report"

                    ElseIf code = myTransactCode.CPostExpenseType Then  'post_expensetype
                        traceline = "6"
                        retvalarry = clsData.UploadExpenseType(employeeid, dtwrap)
                        moddescription = "Expense Type"

                    ElseIf code = myTransactCode.CPostVendor Then   'post_vendor
                        traceline = "5"
                        retvalarry = clsData.UploadVendor(employeeid, dtwrap)
                        moddescription = "Vendor"

                    Else
                        errval = "404"
                        errmsg = "ERROR: Incorrect code."
                    End If
                Else
                    errval = "404"
                    errmsg = "ERROR: Wrap datatable empty."
                End If
            Else
                errval = "404"
                errmsg = "ERROR: JSON data empty."
            End If
        Catch ex As Exception
            errval = "404"
            errmsg = "ERROR" & traceline & ": " & ex.Message
        End Try

tonton:
        'dtret = WrapReturnValue(retvalarry, errmsg, errval)
        'convertedvalue = JsonConvert.SerializeObject(WrapReturnValue(retvalarry, errmsg, errval)) 'clsUtility.ConvertDatatableToJson(dtret)

        clsUtility.SaveLog(code, moddescription, employeeid, convertedvalue & IIf(errmsg.Length > 0, errval & " " & errmsg, "").ToString)

        Return WrapReturnValue(retvalarry, errmsg, errval) '(convertedvalue)
    End Function

#Region "Method - Wrap"
    Private Function WrapDataTable(ByVal dt As DataTable, ByVal transtype As myTransactCode) As DataTable
        Dim dtnew As New DataTable
        Dim rw As DataRow

        Try
            If transtype = myTransactCode.CPostFundReleased Then        'post_fundreleased (cashflow_pc)
                dtnew.Columns.Add("cashinid", GetType(Long))            '0
                dtnew.Columns.Add("transno", GetType(Long))             '1
                dtnew.Columns.Add("panelid", GetType(Long))             '2
                dtnew.Columns.Add("cashindate", GetType(DateTime))      '3
                dtnew.Columns.Add("amount", GetType(Decimal))           '4
                dtnew.Columns.Add("transtype", GetType(String))         '5
                dtnew.Columns.Add("comment", GetType(String))           '6
                dtnew.Columns.Add("status", GetType(Long))              '7
                dtnew.Columns.Add("createdbyid", GetType(Long))         '8
                dtnew.Columns.Add("datecreated", GetType(DateTime))     '9
                dtnew.Columns.Add("approvedby", GetType(Long))          '10
                dtnew.Columns.Add("dateapproved", GetType(DateTime))    '11
                dtnew.Columns.Add("updatedby", GetType(Long))           '12
                dtnew.Columns.Add("dateupdated", GetType(DateTime))     '13
                dtnew.Columns.Add("syncstatus", GetType(Boolean))       '14
                dtnew.Columns.Add("syncdate", GetType(DateTime))        '15
                dtnew.Columns.Add("dlsyncstatus", GetType(Boolean))     '16
                dtnew.Columns.Add("dlsyncdate", GetType(DateTime))      '17
                dtnew.Columns.Add("mainsyncstatus", GetType(Boolean))   '18
                dtnew.Columns.Add("mainsyncdate", GetType(DateTime))    '19
                dtnew.Columns.Add("cino", GetType(Long))                '20
                dtnew.Columns.Add("cino_ffo", GetType(Long))            '21
                dtnew.Columns.Add("doctorid", GetType(Long))            '22
                dtnew.Columns.Add("doctorname", GetType(String))        '23
                dtnew.Columns.Add("institutionid", GetType(Long))       '24
                dtnew.Columns.Add("institutionname", GetType(String))   '25
                dtnew.Columns.Add("mode", GetType(Long))                '26
                dtnew.Columns.Add("payee", GetType(String))             '27
                dtnew.Columns.Add("tuprefno", GetType(String))          '28
                dtnew.Columns.Add("fundtypeid", GetType(Long))          '29
                dtnew.Columns.Add("checkno", GetType(String))           '30
                dtnew.Columns.Add("checkdate", GetType(DateTime))       '31
                dtnew.Columns.Add("ffo_frid", GetType(String))          '32
                dtnew.Columns.Add("ffo_requestid", GetType(String))     '33
                dtnew.Columns.Add("ffo_requestamount", GetType(Decimal)) '34

                For i As Integer = 0 To dt.Rows.Count - 1
                    rw = dtnew.NewRow
                    rw(0) = 0
                    rw(1) = 0
                    rw(2) = dt.Rows(i).Item("warehouseid")
                    rw(3) = dt.Rows(i).Item("dateencoded")
                    rw(4) = dt.Rows(i).Item("amount")
                    rw(5) = ""
                    rw(6) = dt.Rows(i).Item("remarks")
                    rw(7) = 2
                    rw(8) = dt.Rows(i).Item("createdbyid")
                    rw(9) = dt.Rows(i).Item("datecreated")
                    rw(10) = dt.Rows(i).Item("approvedbyid")
                    rw(11) = dt.Rows(i).Item("dateapproved")
                    rw(12) = 0
                    rw(13) = "1/1/1900"
                    rw(14) = False
                    rw(15) = "1/1/1900"
                    rw(16) = False
                    rw(17) = "1/1/1900"
                    rw(18) = False
                    rw(19) = "1/1/1900"
                    rw(20) = 0
                    rw(21) = 0
                    rw(22) = dt.Rows(i).Item("doctorid")
                    rw(23) = dt.Rows(i).Item("doctorname")
                    rw(24) = dt.Rows(i).Item("institutionid")
                    rw(25) = dt.Rows(i).Item("institutionname")
                    rw(26) = dt.Rows(i).Item("mode")
                    rw(27) = dt.Rows(i).Item("payee")
                    rw(28) = dt.Rows(i).Item("tuprefno")
                    rw(29) = dt.Rows(i).Item("fundtypeid")
                    rw(30) = dt.Rows(i).Item("checkno")
                    rw(31) = dt.Rows(i).Item("checkdate")
                    rw(32) = dt.Rows(i).Item("ffo_frid")
                    rw(33) = dt.Rows(i).Item("ffo_requestid")
                    rw(34) = dt.Rows(i).Item("ffo_requestamount")

                    dtnew.Rows.Add(rw)
                Next

            ElseIf transtype = myTransactCode.CPostExpenseLiquidation Then    'post_expenseliquidation (cashdisbursement_pc)
                dtnew.Columns.Add("cdid", GetType(Long))                    '0
                dtnew.Columns.Add("cdno", GetType(Long))                    '1
                dtnew.Columns.Add("cddate", GetType(DateTime))              '2
                dtnew.Columns.Add("vendorid", GetType(Long))                '3
                dtnew.Columns.Add("expenseid", GetType(Long))               '4
                dtnew.Columns.Add("vat", GetType(Boolean))                  '5
                dtnew.Columns.Add("refinvoiceno", GetType(String))          '6
                dtnew.Columns.Add("amount", GetType(Decimal))               '7
                dtnew.Columns.Add("panelid", GetType(Long))                 '8
                dtnew.Columns.Add("comment", GetType(String))               '9
                dtnew.Columns.Add("actualdate", GetType(DateTime))          '10
                dtnew.Columns.Add("status", GetType(Long))                  '11
                dtnew.Columns.Add("createdbyid", GetType(Long))             '12
                dtnew.Columns.Add("datecreated", GetType(DateTime))         '13
                dtnew.Columns.Add("approvedbyid", GetType(Long))            '14
                dtnew.Columns.Add("dateapproved", GetType(DateTime))        '15
                dtnew.Columns.Add("updatedbyid", GetType(Long))             '16
                dtnew.Columns.Add("dateupdated", GetType(DateTime))         '17
                dtnew.Columns.Add("syncstatus", GetType(Boolean))           '18
                dtnew.Columns.Add("syncdate", GetType(DateTime))            '19
                dtnew.Columns.Add("dlsyncstatus", GetType(Boolean))         '20
                dtnew.Columns.Add("dlsyncdate", GetType(DateTime))          '21
                dtnew.Columns.Add("mainsyncstatus", GetType(Boolean))       '22
                dtnew.Columns.Add("mainsyncdate", GetType(DateTime))        '23
                dtnew.Columns.Add("cdno_ffo", GetType(Long))                '24
                dtnew.Columns.Add("frid", GetType(Long))                    '25
                dtnew.Columns.Add("ffo_expliqid", GetType(String))          '26
                dtnew.Columns.Add("ffo_frid", GetType(String))              '27
                dtnew.Columns.Add("fundtypeid", GetType(Long))              '28
                dtnew.Columns.Add("ffo_requestid", GetType(String))         '29
                dtnew.Columns.Add("ffo_receiptfilename", GetType(String))   '30

                For i As Integer = 0 To dt.Rows.Count - 1
                    rw = dtnew.NewRow
                    rw(0) = 0
                    rw(1) = 0
                    rw(2) = dt.Rows(i).Item("dateencoded")
                    rw(3) = dt.Rows(i).Item("vendorid")
                    rw(4) = dt.Rows(i).Item("expenseid")
                    rw(5) = dt.Rows(i).Item("vat")
                    rw(6) = dt.Rows(i).Item("refno")
                    rw(7) = dt.Rows(i).Item("amount")
                    rw(8) = dt.Rows(i).Item("warehouseid")
                    rw(9) = dt.Rows(i).Item("remarks")
                    rw(10) = dt.Rows(i).Item("refdate")
                    rw(11) = 0
                    rw(12) = dt.Rows(i).Item("createdbyid")
                    rw(13) = dt.Rows(i).Item("dateencoded")
                    rw(14) = dt.Rows(i).Item("approvedbyid")
                    rw(15) = dt.Rows(i).Item("dateapproved")
                    rw(16) = 0
                    rw(17) = "1/1/1900"
                    rw(18) = False
                    rw(19) = "1/1/1900"
                    rw(20) = False
                    rw(21) = "1/1/1900"
                    rw(22) = False
                    rw(23) = "1/1/1900"
                    rw(24) = 0
                    rw(25) = 0
                    rw(26) = dt.Rows(i).Item("ffo_expenseliquidationid")
                    rw(27) = dt.Rows(i).Item("ffo_frid")
                    rw(28) = dt.Rows(i).Item("fundtypeid")
                    rw(29) = dt.Rows(i).Item("ffo_requestid")
                    rw(30) = dt.Rows(i).Item("ffo_receiptfilename")

                    dtnew.Rows.Add(rw)
                Next

            ElseIf transtype = myTransactCode.CPostVendor Then  'post_vendor
                dtnew.Columns.Add("vendorid", GetType(Long))            '0
                dtnew.Columns.Add("vendorname", GetType(String))        '1
                dtnew.Columns.Add("address", GetType(String))           '2
                dtnew.Columns.Add("telephone", GetType(String))         '3
                dtnew.Columns.Add("vatno", GetType(String))             '4
                dtnew.Columns.Add("isdefault", GetType(Boolean))        '5
                dtnew.Columns.Add("createdbyid", GetType(Long))         '6
                dtnew.Columns.Add("datecreated", GetType(DateTime))     '7
                dtnew.Columns.Add("approvedby", GetType(Long))          '8
                dtnew.Columns.Add("dateapproved", GetType(DateTime))    '9
                dtnew.Columns.Add("updatedby", GetType(Long))           '10
                dtnew.Columns.Add("dateupdate", GetType(DateTime))      '11
                dtnew.Columns.Add("status", GetType(Long))              '12
                dtnew.Columns.Add("syncstatus", GetType(Boolean))       '13
                dtnew.Columns.Add("syncdate", GetType(DateTime))        '14
                dtnew.Columns.Add("dlsyncstatus", GetType(Boolean))     '15
                dtnew.Columns.Add("dlsyncdate", GetType(DateTime))      '16
                dtnew.Columns.Add("mainsyncstatus", GetType(Boolean))   '17
                dtnew.Columns.Add("mainsyncdate", GetType(DateTime))    '18
                dtnew.Columns.Add("vat", GetType(Boolean))              '19
                dtnew.Columns.Add("vno", GetType(Long))                 '20
                dtnew.Columns.Add("vno_ffo", GetType(Long))             '21
                dtnew.Columns.Add("vendorcategoryid", GetType(Long))    '22

                For i As Integer = 0 To dt.Rows.Count - 1
                    rw = dtnew.NewRow
                    rw(0) = dt.Rows(i).Item("branchid") 'vendorid
                    rw(1) = dt.Rows(i).Item("vendorname").ToString.Trim
                    rw(2) = dt.Rows(i).Item("address")
                    rw(3) = dt.Rows(i).Item("telephone")
                    rw(4) = dt.Rows(i).Item("vatno")
                    rw(5) = False
                    rw(6) = dt.Rows(i).Item("warehouseid")  'createdbyid
                    rw(7) = dt.Rows(i).Item("dateencoded")
                    rw(8) = 0
                    rw(9) = dt.Rows(i).Item("dateencoded")
                    rw(10) = 0
                    rw(11) = "1/1/1900"
                    rw(12) = 0
                    rw(13) = False
                    rw(14) = "1/1/1900"
                    rw(15) = False
                    rw(16) = "1/1/1900"
                    rw(17) = False
                    rw(18) = "1/1/1900"
                    rw(19) = dt.Rows(i).Item("vat")
                    rw(20) = 0
                    rw(21) = 0
                    rw(22) = dt.Rows(i).Item("vendorcategoryid")

                    dtnew.Rows.Add(rw)
                Next

            ElseIf transtype = myTransactCode.CPostExpenseType Then 'post_expensetype
                dtnew.Columns.Add("expenseid", GetType(Long))           '0
                dtnew.Columns.Add("expensetype", GetType(String))       '1
                dtnew.Columns.Add("isdefault", GetType(Boolean))        '2
                dtnew.Columns.Add("createdbyid", GetType(Long))         '3
                dtnew.Columns.Add("datecreated", GetType(DateTime))     '4
                dtnew.Columns.Add("approvedby", GetType(Long))          '5
                dtnew.Columns.Add("dateapproved", GetType(DateTime))    '6
                dtnew.Columns.Add("updatedbyid", GetType(Long))         '7
                dtnew.Columns.Add("dateupdated", GetType(DateTime))     '8
                dtnew.Columns.Add("status", GetType(Long))              '9
                dtnew.Columns.Add("syncstatus", GetType(Boolean))       '10
                dtnew.Columns.Add("syncdate", GetType(DateTime))        '11
                dtnew.Columns.Add("dlsyncstatus", GetType(Boolean))     '12
                dtnew.Columns.Add("dlsyncdate", GetType(DateTime))      '13
                dtnew.Columns.Add("mainsyncstatus", GetType(Boolean))   '14
                dtnew.Columns.Add("mainsyncdate", GetType(DateTime))    '15
                dtnew.Columns.Add("eno", GetType(Long))                 '16
                dtnew.Columns.Add("eno_ffo", GetType(Long))             '17
                dtnew.Columns.Add("isviewonsfa", GetType(Boolean))      '18

                For i As Integer = 0 To dt.Rows.Count - 1
                    rw = dtnew.NewRow
                    rw(0) = 0
                    rw(1) = dt.Rows(i).Item("expensetype").ToString.Trim
                    rw(2) = False
                    rw(3) = 0
                    rw(4) = dt.Rows(i).Item("dateencoded")
                    rw(5) = 0
                    rw(6) = dt.Rows(i).Item("dateencoded")
                    rw(7) = 0
                    rw(8) = "1/1/1900"
                    rw(9) = 0
                    rw(10) = False
                    rw(11) = "1/1/1900"
                    rw(12) = False
                    rw(13) = "1/1/1900"
                    rw(14) = False
                    rw(15) = "1/1/1900"
                    rw(16) = 0
                    rw(17) = 0
                    rw(18) = True

                    dtnew.Rows.Add(rw)
                Next


            End If
        Catch ex As Exception
            Return Nothing
        End Try

        Return dtnew
    End Function
    Private Function WrapReturnValue(ByVal retval As String, ByVal errormsg As String, ByVal errorval As String) As Models.Response
        Dim dt As New Models.Response
        Dim arrretval() As String

        arrretval = retval.Split(",")

        If arrretval(0).Length > 0 And arrretval(0) = "200" And errorval = "" Then  'Success
            dt.reponsecode = arrretval(0)
            dt.errormessage = ""
            dt.returnvalue = arrretval(1)
        Else     'Error
            dt.reponsecode = errorval
            dt.errormessage = errormsg & IIf(arrretval(0).Length > 0, " ;arrretval(0):" & arrretval(0), "")
            dt.returnvalue = ""
        End If
        Return dt
    End Function
#End Region


End Class