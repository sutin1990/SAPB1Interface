Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Configuration
Imports SAPbobsCOM

Public Class SAPB1Interface

    Dim _dtInvoiceLog As New DataTable
    Dim _dtARCreditMemoLog As New DataTable
    Dim _dtIncomingPaymentLog As New DataTable
    Dim _dtInventoryTransferLog As New DataTable
    Dim _dtInventoryTransferrequestLog As New DataTable

    Dim _dtInvHdr As New DataTable
    Dim _dtInvLine As New DataTable
    Dim _dtInvSerial As New DataTable
    Dim _dtInvDeposit As New DataTable

    Dim _dtCreditMemo As New DataTable
    Dim _dtCreditMemoLine As New DataTable
    Dim _dtCreditMemoSerial As New DataTable

    Dim _dtStfHdr As New DataTable
    Dim _dtStfLine As New DataTable
    Dim _dtStfSerial As New DataTable

    Dim _dtStfrequestHdr As New DataTable
    Dim _dtStfrequestLine As New DataTable
    Dim _dtStfrequestSerial As New DataTable

    Dim _TransDate As DateTime = Nothing
    Dim cmd As SqlCommand
    Dim _fs As System.IO.StreamWriter

    Dim da_CheckDocnum As New SqlDataAdapter
    Dim ds_CheckDocnum As New DataSet
    Dim dt_CheckDocnum As New DataTable

    Dim header_Payment As New DataTable
    Dim detail_Payment As New DataTable
    Dim mean_Payment As New DataTable

    Public getdate_file As String = ""

    Sub ConnectSAP()
        getdate_file = "AISLog" & Date.Now().ToString("yyMMdd") & ".txt"
        Try
            'If File.Exists(Application.StartupPath & getdate_file) Then
            '    File.Delete(Application.StartupPath & getdate_file)
            'End If

            oCompany = New SAPbobsCOM.Company
            oCompany.Server = ConfigurationManager.AppSettings.Item("SAPServerName")
            oCompany.CompanyDB = ConfigurationManager.AppSettings.Item("SAPDbName")
            oCompany.UserName = ConfigurationManager.AppSettings.Item("SAPUserName")
            oCompany.Password = ConfigurationManager.AppSettings.Item("SAPUserPassword")
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016
            oCompany.DbUserName = ConfigurationManager.AppSettings.Item("SAPDbUserName")
            oCompany.DbPassword = ConfigurationManager.AppSettings.Item("SAPDbPassword")
            oCompany.LicenseServer = ConfigurationManager.AppSettings.Item("SAPLicenseServer")
            lRetCode = oCompany.Connect

            oCompany.GetLastError(lErrorCode, sErrMsg)

            If lRetCode = 0 Then
                _fs = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & getdate_file, True)
                _fs.WriteLine(Now.ToString)
                _fs.WriteLine("Connect SAP B1[" & oCompany.CompanyDB & "] Success.")
                _fs.WriteLine("-----------------------------------------------------")
                _fs.Close()
            Else
                _fs = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & getdate_file, True)
                _fs.WriteLine(Now.ToString)
                _fs.WriteLine("Connect SAP B1[" & oCompany.CompanyDB & "] Fail.")
                _fs.WriteLine("Error Code : " & lErrorCode.ToString)
                _fs.WriteLine("Error Message : " & sErrMsg)
                _fs.WriteLine("-----------------------------------------------------")
                _fs.Close()
            End If

        Catch ex As Exception
            _fs = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & getdate_file, True)
            _fs.WriteLine(Now.ToString)
            _fs.WriteLine("Connect BOne Error.")
            _fs.WriteLine("Error Message : " & ex.Message)
            _fs.WriteLine("-----------------------------------------------------")
            _fs.Close()
        End Try
    End Sub

    Sub Initial()
        Try
            'A/R Credit Memo
            _dtARCreditMemoLog = New DataTable
            _dtARCreditMemoLog.Columns.Add("Title", GetType(System.String))
            _dtARCreditMemoLog.Columns.Add("TransDate", GetType(System.DateTime))
            _dtARCreditMemoLog.Columns.Add("Status", GetType(System.String))
            _dtARCreditMemoLog.Columns.Add("POSDocNum", GetType(System.String)) 'from POS(Interface)
            _dtARCreditMemoLog.Columns.Add("docEntry", GetType(System.String)) 'SAP Running Doc Number 
            _dtARCreditMemoLog.Columns.Add("ErrCode", GetType(System.Int64))
            _dtARCreditMemoLog.Columns.Add("ErrMsg", GetType(System.String))
            _dtARCreditMemoLog.Columns.Add("StatusPast", GetType(System.Boolean))




            'Invoice
            _dtInvoiceLog = New DataTable
            _dtInvoiceLog.Columns.Add("Title", GetType(System.String))
            _dtInvoiceLog.Columns.Add("TransDate", GetType(System.DateTime))
            _dtInvoiceLog.Columns.Add("Status", GetType(System.String))
            _dtInvoiceLog.Columns.Add("POSDocNum", GetType(System.String)) 'from POS(Interface)
            _dtInvoiceLog.Columns.Add("docEntry", GetType(System.String)) 'SAP Running Doc Number 
            _dtInvoiceLog.Columns.Add("ErrCode", GetType(System.Int64))
            _dtInvoiceLog.Columns.Add("ErrMsg", GetType(System.String))
            _dtInvoiceLog.Columns.Add("StatusPast", GetType(System.Boolean))
            _dtInvoiceLog.CaseSensitive = False

            'Payment
            _dtIncomingPaymentLog = New DataTable
            _dtIncomingPaymentLog.Columns.Add("Title", GetType(System.String))
            _dtIncomingPaymentLog.Columns.Add("TransDate", GetType(System.DateTime))
            _dtIncomingPaymentLog.Columns.Add("Status", GetType(System.String))
            _dtIncomingPaymentLog.Columns.Add("POSDocNum", GetType(System.String))
            _dtIncomingPaymentLog.Columns.Add("docEntry", GetType(System.String))
            _dtIncomingPaymentLog.Columns.Add("ErrCode", GetType(System.Int64))
            _dtIncomingPaymentLog.Columns.Add("ErrMsg", GetType(System.String))
            _dtIncomingPaymentLog.Columns.Add("StatusPast", GetType(System.Boolean))
            _dtIncomingPaymentLog.CaseSensitive = False

            'InventoryTransfer
            _dtInventoryTransferLog = New DataTable
            _dtInventoryTransferLog.Columns.Add("Title", GetType(System.String))
            _dtInventoryTransferLog.Columns.Add("TransDate", GetType(System.DateTime))
            _dtInventoryTransferLog.Columns.Add("Status", GetType(System.String))
            _dtInventoryTransferLog.Columns.Add("POSDocNum", GetType(System.String))
            _dtInventoryTransferLog.Columns.Add("docEntry", GetType(System.String))
            _dtInventoryTransferLog.Columns.Add("ErrCode", GetType(System.Int64))
            _dtInventoryTransferLog.Columns.Add("ErrMsg", GetType(System.String))
            _dtInventoryTransferLog.Columns.Add("StatusPast", GetType(System.Boolean))
            _dtInventoryTransferLog.CaseSensitive = False

            'InventoryTransferrequest
            _dtInventoryTransferrequestLog = New DataTable
            _dtInventoryTransferrequestLog.Columns.Add("Title", GetType(System.String))
            _dtInventoryTransferrequestLog.Columns.Add("TransDate", GetType(System.DateTime))
            _dtInventoryTransferrequestLog.Columns.Add("Status", GetType(System.String))
            _dtInventoryTransferrequestLog.Columns.Add("POSDocNum", GetType(System.String))
            _dtInventoryTransferrequestLog.Columns.Add("docEntry", GetType(System.String))
            _dtInventoryTransferrequestLog.Columns.Add("ErrCode", GetType(System.Int64))
            _dtInventoryTransferrequestLog.Columns.Add("ErrMsg", GetType(System.String))
            _dtInventoryTransferrequestLog.Columns.Add("StatusPast", GetType(System.Boolean))
            _dtInventoryTransferrequestLog.CaseSensitive = False
        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

    Sub Execute()
        Try
            ConnectSAP()

            Dim da As New SqlDataAdapter
            Dim dsInvoice As New DataSet
            Dim dsCreditMemo As New DataSet
            Dim dsIncoming As New DataSet
            Dim dsStock As New DataSet
            Dim dsStockRequest As New DataSet

            _TransDate = Now

            'CheckDocnum()

            'Dim dr_CheckDocnum As DataRow() = dt_CheckDocnum.Select("DocNum = '000000005'")
            'MessageBox.Show(dr_CheckDocnum.Count)

            '1 prepare datasource
            Initial()



            '2 read config =====================================================




            'Read Invoice ======================================================
            ConnectInterfaceDB()
            Dim returnValue As String()

            da = New SqlDataAdapter



            ''Start Incoming Payment ==============================================================================
            'da = New SqlDataAdapter
            'cmd = New SqlCommand
            'With cmd
            '    .Connection = oConn
            '    .CommandType = CommandType.StoredProcedure
            '    .CommandTimeout = 0
            '    .CommandText = "m_sp_ais_Payment"
            '    '.CommandText = ""
            'End With
            'da.SelectCommand = cmd
            'da.Fill(dsIncoming)

            'If dsIncoming.Tables.Count > 0 Then

            '    header_Payment = dsIncoming.Tables(0).Copy 'Invoice Header

            '    If dsIncoming.Tables.Count > 1 Then detail_Payment = dsIncoming.Tables(1).Copy 'Invoice Line
            '    If dsIncoming.Tables.Count > 2 Then mean_Payment = dsIncoming.Tables(2).Copy 'Invoice Serial

            '    'CreateIncomingPayment(header_Payment, detail_Payment, mean_Payment, "")
            '    CreateIncomingPayment(header_Payment, detail_Payment, mean_Payment, "", "O", 0, "IN")
            'End If
            ''End Incoming Payment ==============================================================================




            Try
                'For Each pp In My.Application.CommandLineArgs
                '    MessageBox.Show(pp)
                'Next
                returnValue = Environment.GetCommandLineArgs()

                If returnValue.Length > 1 Then
                    'MessageBox.Show(returnValue(1).ToString())
                    If returnValue(1).ToString() = "Invoice/Payment" Then
                        'Start Invoice ========================================================================================
                        cmd = New SqlCommand
                        With cmd
                            .Connection = oConn
                            .CommandType = CommandType.StoredProcedure
                            .CommandTimeout = 0
                            .CommandText = "m_sp_ais_Invoice"
                        End With
                        da.SelectCommand = cmd
                        da.Fill(dsInvoice)

                        If dsInvoice.Tables.Count > 0 Then

                            _dtInvHdr = dsInvoice.Tables(0).Copy 'Invoice Header

                            If dsInvoice.Tables.Count > 1 Then _dtInvLine = dsInvoice.Tables(1).Copy 'Invoice Line
                            If dsInvoice.Tables.Count > 2 Then _dtInvSerial = dsInvoice.Tables(2).Copy 'Invoice Serial
                            If dsInvoice.Tables.Count > 3 Then _dtInvDeposit = dsInvoice.Tables(3).Copy

                            If dsInvoice.Tables.Count > 4 Then header_Payment = dsInvoice.Tables(4).Copy
                            If dsInvoice.Tables.Count > 5 Then detail_Payment = dsInvoice.Tables(5).Copy
                            If dsInvoice.Tables.Count > 6 Then mean_Payment = dsInvoice.Tables(6).Copy

                            CreateInvoice(_dtInvHdr, _dtInvLine, _dtInvSerial)
                        End If
                        'End Invoice ========================================================================================
                    End If

                    If returnValue(1).ToString() = "MovementRequest" Then
                        '=========== START GET Data From Stored Inventory Transfer================
                        cmd = New SqlCommand
                        With cmd
                            .Connection = oConn
                            .CommandType = CommandType.StoredProcedure
                            .CommandTimeout = 0
                            .CommandText = "m_sp_ais_InvtTRReqSAP"
                        End With
                        da.SelectCommand = cmd
                        da.Fill(dsStockRequest)

                        If dsStockRequest.Tables.Count > 0 Then

                            _dtStfrequestHdr = dsStockRequest.Tables(0).Copy 'InventoryRequest Header

                            If dsStockRequest.Tables.Count > 1 Then _dtStfrequestLine = dsStockRequest.Tables(1).Copy 'InventoryRequest Line
                            'If dsStockRequest.Tables.Count > 2 Then _dtStfrequestSerial = dsStockRequest.Tables(2).Copy 'InventoryRequest Serial

                            CreateInventoryTransferrequest(_dtStfrequestHdr, _dtStfrequestLine)
                        End If
                        '=========== END GET Data From Stored Inventory TransferRequest================
                    End If

                    If returnValue(1).ToString() = "Movement" Then
                        '=========== START GET Data From Stored Inventory Transfer================
                        cmd = New SqlCommand
                        With cmd
                            .Connection = oConn
                            .CommandType = CommandType.StoredProcedure
                            .CommandTimeout = 0
                            .CommandText = "m_sp_ais_InvtTRSAP"
                        End With
                        da.SelectCommand = cmd
                        da.Fill(dsStock)

                        If dsStock.Tables.Count > 0 Then

                            _dtStfHdr = dsStock.Tables(0).Copy 'Inventory Header

                            If dsStock.Tables.Count > 1 Then _dtStfLine = dsStock.Tables(1).Copy 'Inventory Line
                            If dsStock.Tables.Count > 2 Then _dtStfSerial = dsStock.Tables(2).Copy 'Inventory Serial
                            'For i As Integer = _dtStfLine.Rows.Count - 1 To 0 Step -1
                            '    Dim CheckDocnum As DataRow() = _dtStfSerial.Select("DocNum='" & _dtStfLine.Rows(i)("DocNum") & "' and LineNum = " & _dtStfLine.Rows(i)("LineNum"))
                            '    If CheckDocnum.Length = 0 Then
                            '        _dtStfLine.Rows.Remove(_dtStfLine.Rows(i))
                            '    End If
                            'Next


                            CreateInventoryTransfer(_dtStfHdr, _dtStfLine, _dtStfSerial)
                        End If
                        '=========== END GET Data From Stored Inventory Transfer================
                    End If

                    If returnValue(1).ToString() = "CN/DN" Then
                        'Read A/R Credit Memo ==============================================
                        cmd = New SqlCommand
                        da = New SqlDataAdapter
                        With cmd
                            .Connection = oConn
                            .CommandType = CommandType.StoredProcedure
                            .CommandTimeout = 0
                            .CommandText = ""
                        End With
                        'da.SelectCommand = cmd
                        'da.Fill(dsCreditMemo)
                        '==================================================================

                        'A/R Credit Memo =================================================================
                        If dsCreditMemo.Tables.Count > 0 Then

                            '_dtCreditMemo = dsCreditMemo.Tables(0).Copy 'A/R Credit Memo Header

                            'If dsCreditMemo.Tables.Count > 1 Then _dtCreditMemoLine = dsCreditMemo.Tables(1).Copy 'A/R Credit Memo Line
                            'If dsCreditMemo.Tables.Count > 2 Then _dtCreditMemoSerial = dsCreditMemo.Tables(2).Copy 'A/R Credit Memo Serial

                            'CreateARCreditMemo(_dtCreditMemo, _dtCreditMemoLine, _dtCreditMemoSerial)
                        End If
                        '===================================================================================
                    End If
                Else
                    MessageBox.Show("Nothing Parameter Process.")
                End If
            Catch sqlex As SqlException
                MessageBox.Show(sqlex.Message)
            Catch ex As Exception
                'MessageBox.Show(ex.Message)
                Throw ex
            End Try

            CloseConnect()

            '=======================================================test hardcode Inventory====================================
            Dim _dtInventoryH As New DataTable
            _dtInventoryH.Columns.Add("DocNum", GetType(System.Int32))
            _dtInventoryH.Columns.Add("DocDate", GetType(System.DateTime))
            _dtInventoryH.Columns.Add("TaxDate", GetType(System.DateTime))
            _dtInventoryH.Columns.Add("CardCode", GetType(System.String))
            _dtInventoryH.Columns.Add("FromWhs", GetType(System.String))
            _dtInventoryH.Columns.Add("ToWhs", GetType(System.String))
            _dtInventoryH.Columns.Add("Comments", GetType(System.String))
            _dtInventoryH.Rows.Add({310100001, Now.Date, Now.Date, "BB01", "WH05", "WH07", "310100001"})
            _dtInventoryH.Rows.Add({310100002, Now.Date, Now.Date, "BB02", "WH05", "WH08", "310100002"})

            Dim _dtInventoryL As New DataTable
            _dtInventoryL.Columns.Add("DocNum", GetType(System.Int32))
            _dtInventoryL.Columns.Add("ItemCode", GetType(System.String))
            _dtInventoryL.Columns.Add("Quantity", GetType(System.Double))
            _dtInventoryL.Columns.Add("LineNum", GetType(System.Int32))
            _dtInventoryL.Rows.Add({310100001, "MZB6718EU", 1, 1})
            _dtInventoryL.Rows.Add({310100002, "MZB6718EU", 1, 1})
            '_dtInventoryL.Rows.Add({310100001, "UYG4021RT", 1, 23})
            '_dtInventoryL.Rows.Add({310100001, "UYG4021RT", 1, 24})
            '_dtInventoryL.Rows.Add({310100001, "UYG4021RT", 1, 25})

            Dim _dtInventoryS As New DataTable
            _dtInventoryS.Columns.Add("DocNum", GetType(System.Int32))
            _dtInventoryS.Columns.Add("Serialnum", GetType(System.String))
            _dtInventoryS.Columns.Add("Quantity", GetType(System.Double))
            _dtInventoryS.Columns.Add("LineNum", GetType(System.Int32))
            _dtInventoryS.Rows.Add({310100001, "MZ00001", 1, 1})
            _dtInventoryS.Rows.Add({310100002, "MZ00002", 1, 1})

            'CreateInventoryTransfer(_dtInventoryH, _dtInventoryL, _dtInventoryS)
            '=========================================== end test hardcode Inventory====================================================


            'last writelog
            'WriteLog()

            _fs = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & getdate_file, True)
            _fs.WriteLine("End")
            _fs.WriteLine("-----------------------------------------------------")
            _fs.Close()
            'If File.Exists(Application.StartupPath & getdate_file) Then Process.Start(Application.StartupPath & getdate_file)


        Catch ex As Exception
            _fs = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & getdate_file, True)
            _fs.WriteLine(Now.ToString)
            _fs.WriteLine("Execute")
            _fs.WriteLine("Error Message : " & ex.Message)
            _fs.WriteLine("-----------------------------------------------------")
            _fs.Close()
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

    Function m_sp_ais_WriteLog(ByVal type As String, ByVal skey As Integer, ByVal LoadStatus As String, ByVal SAPRecID As String, ByVal Title As String, ByVal POSDocNum As String, ByVal ErrCode As String, ByVal ErrMsg As String)
        Dim statuscode
        Dim statusmessage
        Try
            Dim dbreader As SqlDataReader

            ConnectInterfaceDB()
            cmd = New SqlCommand
            With cmd
                .Connection = oConn
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .CommandText = "m_sp_ais_WriteLog"
                .Parameters.Add("@type", SqlDbType.NVarChar).Value = type
                .Parameters.Add("@skey", SqlDbType.Int).Value = skey
                .Parameters.Add("@LoadtoSAPDate", SqlDbType.DateTime).Value = Now
                .Parameters.Add("@LoadStatus", SqlDbType.NVarChar, 1).Value = IIf(LoadStatus = "S", "Y", "E")
                .Parameters.Add("@SAPRecID", SqlDbType.Int).Value = If(SAPRecID = "", 0, Convert.ToInt32(SAPRecID))
                .Parameters.Add("@Title", SqlDbType.NVarChar).Value = Title
                .Parameters.Add("@POSDocNum", SqlDbType.NVarChar).Value = POSDocNum
                .Parameters.Add("@ErrCode", SqlDbType.NVarChar).Value = ErrCode
                .Parameters.Add("@ErrMsg", SqlDbType.NVarChar).Value = ErrMsg
            End With
            dbreader = cmd.ExecuteReader()


            If dbreader.HasRows() Then
                While dbreader.Read()
                    statuscode = dbreader.Item("statuscode")
                    statusmessage = dbreader.Item("statusmessage")
                    If statuscode <> 0 Then
                        _fs = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & getdate_file, True)
                        _fs.WriteLine(Now.ToString)
                        _fs.WriteLine("AddPayment" & statusmessage & " skey=>" & skey & ", docnum=>" & POSDocNum)
                        _fs.WriteLine("-----------------------------------------------------")
                        _fs.Close()
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                End While
            End If

            CloseConnect()
            Return statuscode

        Catch ex As Exception
            If ex.Message = "Could not find stored procedure 'm_sp_ais_WriteLog'" Then
                _fs = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & getdate_file, True)
                _fs.WriteLine(Now.ToString)
                _fs.WriteLine("UpdateInvoice")
                _fs.WriteLine("Error Message : " & ex.Message & "code=>" & ex.Data.Keys.ToString())
                _fs.WriteLine("-----------------------------------------------------")
                _fs.Close()

                CloseConnect()
                Return statuscode
            Else
                _fs = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & getdate_file, True)
                _fs.WriteLine(Now.ToString)
                _fs.WriteLine("UpdateInvoice")
                _fs.WriteLine("Error Message : " & ex.Message & "code=>" & ex.Data.Keys.ToString())
                _fs.WriteLine("-----------------------------------------------------")
                _fs.Close()

                CloseConnect()
                Return statuscode
                Throw ex
            End If

        End Try

    End Function

#Region "Call Function SAP"

    'invoice
    Public Sub CreateInvoice(header As DataTable, detail As DataTable, serial As DataTable)
        Try
            UpdateProcess("Invoice", "notIdle")
            Dim sqlTrans As SqlTransaction = Nothing
            Dim sqlTrans1 As SqlTransaction = Nothing
            Dim dbreader As SqlDataReader
            Dim statusWriteLog
            Dim msgError = ""
            For Each row As DataRow In header.Select()

                Dim oDoc As SAPbobsCOM.Documents = Nothing
                Dim oDoc_cancel As SAPbobsCOM.Documents = Nothing
                Dim getkey
                If row("DocType").ToString().ToUpper() = "IN" Then
                    oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                    Dim docEntry As String = ""
                    If row("DocStatus").ToString().ToUpper() = "V" Then
                        'CancelRecord(row("SAPRecID"), row("DocNum"), row("DocType"))
                        'CreateARCreditMemo(_dtInvHdr, _dtInvLine, _dtInvSerial, row("DocType"), row("SAPRecID"))
                        If row("SAPRecID_Invoice") IsNot DBNull.Value Then
                            row("SAPRecID_Invoice") = NullZero(row("SAPRecID_Invoice"))

                            Dim rs As SAPbobsCOM.Recordset
                            Dim strsql As String
                            Dim statuscancel = ""
                            strsql = "select DocEntry from OINV where DocEntry = " & row("SAPRecID_Invoice") & ""
                            rs = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

                            'lRetCode = oDocP.GetByKey(1906000005)
                            rs.DoQuery(strsql)
                            'rs.MoveFirst()
                            If Not rs.EoF Then

                                statuscancel = CancelRecord(row("SAPRecID_Invoice"), row("DocNum"), row("DocType"), row("skey"))

                                If statuscancel.ToString().ToUpper() = "S" Then '' ถ้า canceled ผ่าน หรือ ถูก canceled ไปก่อนหน้านี้แล้ว
                                    msgError = CreateARCreditMemo(row("DocType"), row("SAPRecID_Invoice"), row("skey"))
                                    If msgError = "S" Then
                                        AddErrorLog("Cenceled Invoice", row("DocNum"), 0, row("SAPRecID_Invoice"), _dtInvoiceLog, True, row("skey"))
                                        'add WriteLog and update Invoice Voice Success
                                        statusWriteLog = m_sp_ais_WriteLog("InvHeader", row("skey"), "S", row("SAPRecID_Invoice"), "Canceled Invoice", row("DocNum"), lRetCode, sErrMsg)
                                    Else
                                        AddErrorLog("Cenceled Invoice", row("DocNum"), -10, row("SAPRecID_Invoice"), _dtInvoiceLog, False, row("skey"))
                                        statusWriteLog = m_sp_ais_WriteLog("InvHeader", row("skey"), "E", row("SAPRecID_Invoice"), "Canceled Invoice", row("DocNum"), "-10", msgError)
                                    End If

                                Else
                                    AddErrorLog("Cenceled Invoice", row("DocNum"), -10, row("SAPRecID_Invoice"), _dtInvoiceLog, False, row("skey"))
                                    statusWriteLog = m_sp_ais_WriteLog("InvHeader", row("skey"), "E", row("SAPRecID_Invoice"), "Canceled Invoice", row("DocNum"), "-10", sErrMsg)
                                End If
                            Else
                                docEntry = "This DocEntry [" & row("SAPRecID_Invoice") & "] not found in Payment "
                                'docEntry = String.Format("{0}-{1}", oCompany.GetLastErrorCode, oCompany.GetLastErrorDescription)
                                AddErrorLog("Cenceled Invoice", row("DocNum"), -10, docEntry, _dtInvoiceLog, False, row("skey"))
                                statusWriteLog = m_sp_ais_WriteLog("InvHeader", row("skey"), "E", row("SAPRecID_Invoice"), "Canceled Invoice", row("DocNum"), "-10", docEntry)
                            End If
                            'CreateIncomingPayment(header_Payment, detail_Payment, mean_Payment, row("DocNum"), "V", row("SAPRecID"), Convert.ToInt32(docEntry), row("DocType"))
                        Else
                            docEntry = "This DocEntry ['NULL'] not found in Invoice "
                            AddErrorLog("Cenceled Invoice", row("DocNum"), -10, docEntry, _dtInvoiceLog, False, row("skey"))
                            statusWriteLog = m_sp_ais_WriteLog("InvHeader", row("skey"), "E", DBNull.Value.ToString(), "Canceled Invoice", row("DocNum"), "-10", docEntry)
                        End If

                    Else
                        'add header
                        oDoc.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES
                        'oDoc.ManualNumber = "Yes"
                        '===================cancel=====================
                        'oDoc.GetByKey(row("DocNum"))
                        'Dim oCancelDoc As SAPbobsCOM.Documents = oDoc.CreateCancellationDocument()
                        'oCancelDoc.Add()
                        oDoc.DocNum = row("DocNum")
                        oDoc.CardCode = row("CardCode")
                        oDoc.DocDate = row("DocDate")
                        oDoc.DocDueDate = row("DocDueDate")
                        oDoc.NumAtCard = row("NumAtCard") 'Customer Ref No.
                        oDoc.PayToCode = NullString(row("PayToCode"))
                        oDoc.DocCurrency = NullString(row("DocCur"))
                        oDoc.DocRate = NullZero(row("DocRate"))
                        oDoc.DocTotal = NullZero(row("DocTotal"))
                        oDoc.TaxDate = IIf(row("TaxDate") Is DBNull.Value, Nothing, row("TaxDate"))
                        oDoc.Comments = NullString(row("Comments"))
                        oDoc.AddressExtension.BillToStreet = NullString(row("StreetB"))
                        oDoc.AddressExtension.BillToBlock = NullString(row("BlockB"))
                        oDoc.AddressExtension.BillToBuilding = NullString(row("BuildingB"))
                        oDoc.AddressExtension.BillToCity = NullString(row("CityB")) & " " & NullString(row("StateB"))
                        oDoc.AddressExtension.BillToZipCode = NullString(row("ZipCodeB"))
                        'oDoc.AddressExtension.BillToState = NullString(row("StateB"))
                        oDoc.AddressExtension.BillToCountry = NullString(row("CountryB"))
                        oDoc.AddressExtension.BillToStreetNo = NullString(row("StreetNoB"))
                        oDoc.AddressExtension.BillToGlobalLocationNumber = row("GlbLocNumB")
                        oDoc.UserFields.Fields.Item("U_M_Branch_No").Value = NullString(row("U_M_Branch_No"))
                        oDoc.UserFields.Fields.Item("U_F_POS_Ref").Value = NullString(row("POSReference"))
                        oDoc.UserFields.Fields.Item("U_M_Original_Inv").Value = NullString(row("OriginalInvNo"))
                        oDoc.SalesPersonCode = NullZero(row("Slpcode"))

                        For Each dr As DataRow In detail.Select("DocNum = " & row("DocNum"))

                            Dim oLines As SAPbobsCOM.Document_Lines = oDoc.Lines
                            Dim oSerial As SAPbobsCOM.SerialNumbers = oLines.SerialNumbers

                            'add line ==================================
                            With oLines
                                .ItemCode = dr("ItemCode")
                                .Quantity = NullZero(dr("Quantity"))
                                .DiscountPercent = NullZero(dr("DiscPrcnt"))
                                .WarehouseCode = NullString(dr("WhsCode"))
                                .VatGroup = NullString(dr("VatGroup"))
                                .PriceAfterVAT = NullZero(dr("PriceAfVAT"))
                                .TaxTotal = NullZero(dr("LineVat"))
                                .GrossTotal = NullZero(dr("GTotal"))
                                .GrossPrice = NullZero(dr("GPBefDisc"))
                                .COGSCostingCode = NullString(dr("CogsOcrCod"))
                                .COGSCostingCode2 = NullString(dr("CogsOcrCo2"))
                                .COGSCostingCode3 = NullString(dr("CogsOcrCo3"))
                            End With

                            'add serial number ======================================================================================
                            For Each r As DataRow In serial.Select("DocNum = " & row("DocNum") & " and LineNum = " & dr("LineNum"))
                                If NullString(r("Serialnum")) <> "" Then
                                    oSerial.InternalSerialNumber = r("Serialnum")
                                    oSerial.Add()
                                End If
                            Next

                            oLines.Add()
                        Next
                        If Not oCompany.InTransaction Then
                            oCompany.StartTransaction()
                        End If

                        lRetCode = oDoc.Add()
                        oCompany.GetLastError(lErrorCode, sErrMsg)
                        oCompany.GetNewObjectCode(docEntry)
                        'DIErrorHandler(String.Format("{0} Create Document", docEntry))
                        AddErrorLog("AddInvoice", row("DocNum"), lRetCode, docEntry, _dtInvoiceLog, True, row("skey"))
                        If lRetCode = 0 Then
                            Dim count_table = header_Payment.Select("DocNum=" & row("DocNum")).Count
                            If count_table > 0 Then
                                'add WriteLog and update Invoice Success
                                statusWriteLog = m_sp_ais_WriteLog("InvHeader", row("skey"), "S", docEntry, "AddInvoice", row("DocNum"), lRetCode, sErrMsg)

                                If statusWriteLog = 0 Then
                                    CreateIncomingPayment(header_Payment, detail_Payment, mean_Payment, row("DocNum"), "O", Convert.ToInt32(docEntry), row("DocType"))
                                End If

                            Else
                                AddErrorLog("AddPayment", row("DocNum"), lRetCode, docEntry, _dtIncomingPaymentLog, True, row("skey"))
                            End If
                        Else
                            'add WriteLog and update Invoice Error
                            statusWriteLog = m_sp_ais_WriteLog("InvHeader", row("skey"), "E", docEntry, "AddInvoice", row("DocNum"), lRetCode, sErrMsg)

                            If oCompany.InTransaction Then
                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                        End If

                    End If
                End If

                If row("DocType").ToString().ToUpper() = "DT" Then

                    Dim docEntry As String = ""
                    If row("DocStatus").ToString().ToUpper() = "V" Then
                        Dim rs As SAPbobsCOM.Recordset
                        Dim strsql As String
                        Dim statuscancel = ""
                        If row("SAPRecID_Invoice") IsNot DBNull.Value Then
                            row("SAPRecID_Invoice") = NullZero(row("SAPRecID_Invoice"))
                            strsql = "select * from ODPI where DocEntry = " & row("SAPRecID_Invoice") & ""
                            rs = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

                            'lRetCode = oDocP.GetByKey(1906000005)
                            rs.DoQuery(strsql)
                            rs.MoveFirst()
                            If Not rs.EoF Then
                                statuscancel = CancelRecord(row("SAPRecID_Invoice"), row("DocNum"), row("DocType"), row("skey"))
                                If statuscancel.ToString().ToUpper() = "S" Then '' ถ้า canceled ผ่าน หรือ ถูก canceled ไปก่อนหน้านี้แล้ว
                                    msgError = CreateARCreditMemo(row("DocType"), row("SAPRecID_Invoice"), row("skey"))
                                    If msgError = "S" Then
                                        AddErrorLog("Cenceled Downpayment", row("DocNum"), 0, row("SAPRecID_Invoice"), _dtInvoiceLog, True, row("skey"))
                                        'add WriteLog and update Invoice Voice Success
                                        statusWriteLog = m_sp_ais_WriteLog("InvHeader", row("skey"), "S", row("SAPRecID_Invoice"), "Canceled Downpayment", row("DocNum"), lRetCode, sErrMsg)
                                    Else
                                        AddErrorLog("Cenceled Downpayment", row("DocNum"), -10, row("SAPRecID_Invoice"), _dtInvoiceLog, False, row("skey"))
                                        statusWriteLog = m_sp_ais_WriteLog("InvHeader", row("skey"), "E", row("SAPRecID_Invoice"), "Canceled Downpayment", row("DocNum"), "-10", msgError)
                                    End If
                                Else
                                    AddErrorLog("Cenceled Downpayment", row("DocNum"), -10, row("SAPRecID_Invoice"), _dtInvoiceLog, False, row("skey"))
                                    statusWriteLog = m_sp_ais_WriteLog("InvHeader", row("skey"), "E", row("SAPRecID_Invoice"), "Cenceled Downpayment", row("DocNum"), "-10", sErrMsg)
                                End If

                            Else
                                docEntry = "This DocEntry [" & row("SAPRecID_Invoice") & "] not found in Payment "
                                AddErrorLog("Cenceled Downpayment", row("DocNum"), lRetCode, docEntry, _dtInvoiceLog, False, row("skey"))
                                statusWriteLog = m_sp_ais_WriteLog("InvHeader", row("skey"), "E", row("SAPRecID_Invoice"), "Cenceled Downpayment", row("DocNum"), "-10", docEntry)
                            End If
                        Else

                            docEntry = "This DocEntry [ 'NULL' ] not found in Invoice "
                            AddErrorLog("Cenceled Downpayment", row("DocNum"), -10, docEntry, _dtInvoiceLog, False, row("skey"))
                            statusWriteLog = m_sp_ais_WriteLog("InvHeader", row("skey"), "E", DBNull.Value.ToString(), "Canceled Downpayment", row("DocNum"), "-10", sErrMsg)
                        End If


                    Else
                        oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
                        oDoc.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES
                        oDoc.DocNum = row("DocNum")
                        oDoc.DownPaymentType = SAPbobsCOM.DownPaymentTypeEnum.dptInvoice
                        oDoc.CardCode = row("CardCode")
                        oDoc.DocDate = row("DocDate")
                        oDoc.DocDueDate = row("DocDueDate")
                        oDoc.NumAtCard = row("NumAtCard") 'Customer Ref No.
                        oDoc.PayToCode = NullString(row("PayToCode"))
                        oDoc.DocCurrency = NullString(row("DocCur"))
                        oDoc.DocRate = NullZero(row("DocRate"))
                        oDoc.DocTotal = NullZero(row("DocTotal"))
                        oDoc.TaxDate = IIf(row("TaxDate") Is DBNull.Value, Nothing, row("TaxDate"))
                        oDoc.Comments = NullString(row("Comments"))
                        oDoc.AddressExtension.BillToStreet = NullString(row("StreetB"))
                        oDoc.AddressExtension.BillToBlock = NullString(row("BlockB"))
                        oDoc.AddressExtension.BillToBuilding = NullString(row("BuildingB"))
                        oDoc.AddressExtension.BillToCity = NullString(row("CityB")) & " " & NullString(row("StateB"))
                        oDoc.AddressExtension.BillToZipCode = NullString(row("ZipCodeB"))
                        'oDoc.AddressExtension.BillToState = NullString(row("StateB"))
                        oDoc.AddressExtension.BillToCountry = NullString(row("CountryB"))
                        oDoc.AddressExtension.BillToStreetNo = NullString(row("StreetNoB"))
                        oDoc.AddressExtension.BillToGlobalLocationNumber = row("GlbLocNumB")
                        oDoc.UserFields.Fields.Item("U_M_Branch_No").Value = NullString(row("U_M_Branch_No"))
                        oDoc.UserFields.Fields.Item("U_F_POS_Ref").Value = NullString(row("POSReference"))
                        oDoc.UserFields.Fields.Item("U_M_Original_Inv").Value = NullString(row("OriginalInvNo"))
                        oDoc.SalesPersonCode = NullZero(row("Slpcode"))

                        For Each dr As DataRow In detail.Select("DocNum = " & row("DocNum"))

                            Dim oLines As SAPbobsCOM.Document_Lines = oDoc.Lines
                            Dim oSerial As SAPbobsCOM.SerialNumbers = oLines.SerialNumbers

                            'add line ==================================
                            With oLines
                                .ItemCode = dr("ItemCode")
                                .Quantity = NullZero(dr("Quantity"))
                                .DiscountPercent = NullZero(dr("DiscPrcnt"))
                                .WarehouseCode = NullString(dr("WhsCode"))
                                .VatGroup = NullString(dr("VatGroup"))
                                .PriceAfterVAT = NullZero(dr("PriceAfVAT"))
                                .TaxTotal = NullZero(dr("LineVat"))
                                .GrossTotal = NullZero(dr("GTotal"))
                                .GrossPrice = NullZero(dr("GPBefDisc"))
                                .CogsOcrCod = NullString(dr("CogsOcrCod"))
                                .CogsOcrCo2 = NullString(dr("CogsOcrCo2"))
                                .CogsOcrCo3 = NullString(dr("CogsOcrCo3"))
                            End With

                            'add serial number ======================================================================================
                            For Each r As DataRow In serial.Select("DocNum = " & row("DocNum") & " and LineNum = " & dr("LineNum"))
                                If NullString(r("Serialnum")) <> "" Then
                                    oSerial.InternalSerialNumber = r("Serialnum")
                                    oSerial.Add()
                                End If
                            Next

                            oLines.Add()
                        Next

                        If Not oCompany.InTransaction Then
                            oCompany.StartTransaction()
                        End If

                        lRetCode = oDoc.Add()
                        oCompany.GetNewObjectCode(docEntry)
                        'DIErrorHandler(String.Format("{0} Create Document", docEntry))
                        AddErrorLog("AddDownPayment", row("DocNum"), lRetCode, docEntry, _dtInvoiceLog, True, row("skey"))
                        Dim count_table = header_Payment.Select("DocNum=" & row("DocNum")).Count
                        If lRetCode = 0 Then
                            If count_table > 0 Then
                                'add WriteLog and update Invoice downpayment success
                                statusWriteLog = m_sp_ais_WriteLog("InvHeader", row("skey"), "S", docEntry, "AddDownPayment", row("DocNum"), lRetCode, sErrMsg)
                                If statusWriteLog = 0 Then
                                    CreateIncomingPayment(header_Payment, detail_Payment, mean_Payment, row("DocNum"), "O", Convert.ToInt32(docEntry), row("DocType"))
                                End If

                            Else
                                AddErrorLog("AddPayment", row("DocNum"), lRetCode, docEntry, _dtIncomingPaymentLog, True, row("skey"))
                            End If
                        Else 'add WriteLog and update Invoice downpayment error
                            m_sp_ais_WriteLog("InvHeader", row("skey"), "E", docEntry, "AddDownPayment", row("DocNum"), lRetCode, sErrMsg)

                            If oCompany.InTransaction Then
                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                        End If
                    End If

                End If

                If row("DocType") = "CN" Then
                    oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)

                    '//add header
                    oDoc.CardCode = row("CardCode") '//Customer
                    oDoc.DocDate = row("DocDate")
                    oDoc.DocDueDate = row("DocDueDate")
                    oDoc.NumAtCard = row("NumAtCard") '//Customer Ref. No.
                    oDoc.Comments = row("Comments")
                    oDoc.UserFields.Fields.Item("U_M_Branch_No").Value = NullString(row("U_M_Branch_No"))
                    oDoc.UserFields.Fields.Item("U_F_POS_Ref").Value = NullString(row("POSReference"))
                    oDoc.UserFields.Fields.Item("U_M_Original_Inv").Value = NullString(row("OriginalInvNo"))

                    For Each dr As DataRow In detail.Select("DocNum = " & row("DocNum"))

                        Dim creditMemoLines As SAPbobsCOM.Document_Lines = oDoc.Lines
                        Dim oSerial As SAPbobsCOM.SerialNumbers = creditMemoLines.SerialNumbers

                        'add line ===========================
                        With creditMemoLines
                            .ItemCode = dr("ItemCode")
                            .Quantity = NullZero(dr("Quantity"))
                            .DiscountPercent = NullZero(dr("DiscPrcnt"))
                            .WarehouseCode = NullString(dr("WhsCode"))
                            .VatGroup = NullString(dr("VatGroup"))
                            .PriceAfterVAT = NullZero(dr("PriceAfVAT"))
                            .TaxTotal = NullZero(dr("LineVat"))
                            .GrossTotal = NullZero(dr("GTotal"))
                            .GrossPrice = NullZero(dr("GPBefDisc"))
                            .CogsOcrCod = NullString(dr("CogsOcrCod"))
                            .CogsOcrCo2 = NullString(dr("CogsOcrCo2"))
                            .CogsOcrCo3 = NullString(dr("CogsOcrCo3"))

                            '.BaseEntry = 5
                        End With

                        'add serial number =======================================================================================
                        For Each r As DataRow In serial.Select("DocNum = " & row("DocNum") & " and LineNum = " & dr("LineNum"))
                            oSerial.InternalSerialNumber = r("Serialnum")
                            oSerial.Add()
                        Next

                        creditMemoLines.Add()

                    Next

                    'execute =======================
                    lRetCode = oDoc.Add()

                    '//Get Running Number of SAP
                    Dim docEntry As String = ""
                    oCompany.GetNewObjectCode(docEntry)

                    AddErrorLog("Add Credit Memo", row("DocNum"), lRetCode, docEntry, _dtARCreditMemoLog, True, row("skey"))
                    'add WriteLog and update Invoice creditnote
                    If lRetCode = 0 Then
                        m_sp_ais_WriteLog("InvHeader", row("skey"), "S", docEntry, "Add Credit Memo", row("DocNum"), lRetCode, sErrMsg)
                    Else
                        m_sp_ais_WriteLog("InvHeader", row("skey"), "E", docEntry, "Add Credit Memo", row("DocNum"), lRetCode, sErrMsg)
                    End If

                End If

            Next

            UpdateProcess("Invoice", "Idle")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function CancelRecord(ByVal SAPRecID As String, ByVal DocNum As String, ByVal typeinvoice As String, ByVal skey As String)
        Try
            Dim docEntry As String = ""
            Dim ErrCode As Long
            Dim ErrMsg As String
            Dim msgtype As String
            Dim InvType As String = ""
            If typeinvoice = "DT" Then
                InvType = "203"
                msgtype = "Downpayment"
            End If
            If typeinvoice = "IN" Then
                InvType = "13"
                msgtype = "Invoice"
            End If

            Dim oDocP As SAPbobsCOM.Payments = Nothing
            Dim rs As SAPbobsCOM.Recordset
            Dim STR1 As String
            Dim getkey
            Dim status_canceled
            Dim statuspast = "S"
            STR1 = "SELECT T0.DocEntry,T0.Canceled FROM ORCT T0 INNER JOIN RCT2 T1 ON T0.DocEntry  = T1.DocNum Where T1.DocEntry = '" & SAPRecID & "' and T1.InvType=" & InvType & ""
            rs = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            'lRetCode = oDocP.GetByKey(1906000005)
            rs.DoQuery(STR1)
            rs.MoveFirst()
            If rs.RecordCount = 0 Then
                statuspast = "E"
                Return statuspast

            End If
            Do While Not rs.EoF
                oDocP = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                getkey = rs.Fields.Item(0).Value.ToString()
                status_canceled = rs.Fields.Item(1).Value.ToString()
                'lRetCode = oDocP.GetByKey(getkey)
                If oDocP.GetByKey(getkey) Then
                    If status_canceled = "N" Then
                        lRetCode = oDocP.Cancel
                        If lRetCode <> 0 Then
                            docEntry = String.Format("{0}-{1}", oCompany.GetLastErrorCode, oCompany.GetLastErrorDescription)
                            AddErrorLog("CancelPayment[" & msgtype & "]", DocNum, lRetCode, docEntry, _dtInvoiceLog, False, skey)
                            statuspast = docEntry
                            Return statuspast
                            Exit Function
                        End If
                    Else '' มีการ canceled ไปแล้ว
                        docEntry = "This DocEntry [" & getkey & "] status is Canceled in Payment "
                        AddErrorLog("CancelPayment" & msgtype & "]", DocNum, lRetCode, getkey, _dtInvoiceLog, True, skey)
                        Return statuspast
                    End If
                    '' add canceled แล้ว success
                    AddErrorLog("CancelPayment" & msgtype & "]", DocNum, lRetCode, getkey, _dtInvoiceLog, True, skey)
                    Return statuspast
                Else
                    docEntry = "This DocEntry [" & getkey & "] not found in Payment "
                    AddErrorLog("CancelPayment" & msgtype & "]", DocNum, lRetCode, docEntry, _dtInvoiceLog, False, skey)
                    statuspast = docEntry
                    Return statuspast
                    Exit Function
                End If
                rs.MoveNext()
            Loop


            'MessageBox.Show(lRetCode)
            'If lRetCode <> 0 Then
            '    oDocP.GetLastError(ErrCode, ErrMsg)
            '    MessageBox.Show("Failed to Retrieve the record " & ErrCode & " " & ErrMsg)
            '    AddErrorLog(type, 1906000005, lRetCode, docEntry, _dtIncomingPaymentLog)
            '    Exit Sub
            'End If

            'lRetCode = oDocP.Cancel
            'If lRetCode <> 0 Then
            '    'oDocP.GetLastError(ErrCode, ErrMsg)
            '    'MessageBox.Show("Failed to Cancel the record " & ErrCode & " " & ErrMsg)
            '    MsgBox(String.Format("{0}-{1}", oCompany.GetLastErrorCode, oCompany.GetLastErrorDescription))
            'End If
            'oCompany.GetNewObjectCode(docEntry)


        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            GC.Collect()
        End Try
        ' Return statuspast
    End Function

    'credit memo
    Public Function CreateARCreditMemo(ByVal obj_type As String, ByVal base_entry As String, ByVal skey As String)
        Dim errormsge = "S"
        Try
            UpdateProcess("CNDN", "notIdle")
            Dim STR_OINV As String
            Dim basetype_serial
            Dim DistNumber

            '######## insert ข้อมูลแบบเช็คในตารางของ sapb1 และดึงข้อมูลมาจากตาราง sapb1 นำมา insert ########

            'Dim rs_OINV = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim rs_OINV As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            STR_OINV = "select CardCode,DocDate,DocDueDate,NumAtCard,Comments,U_M_Branch_No,U_F_POS_Ref,U_M_Original_Inv,DocNum from OINV where DocEntry = " & base_entry & ""
            'rs_OINV = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            rs_OINV.DoQuery(STR_OINV)
            rs_OINV.MoveFirst()
            Do While Not rs_OINV.EoF
                Dim ocreditMemo As SAPbobsCOM.Documents = Nothing
                ocreditMemo = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                ocreditMemo.CardCode = rs_OINV.Fields.Item(0).Value.ToString() '//Customer
                ocreditMemo.DocDate = rs_OINV.Fields.Item(1).Value
                ocreditMemo.DocDueDate = rs_OINV.Fields.Item(2).Value
                ocreditMemo.NumAtCard = rs_OINV.Fields.Item(3).Value '//Customer Ref. No.
                ocreditMemo.Comments = rs_OINV.Fields.Item(4).Value
                ocreditMemo.UserFields.Fields.Item("U_M_Branch_No").Value = NullString(rs_OINV.Fields.Item(5).Value)
                ocreditMemo.UserFields.Fields.Item("U_F_POS_Ref").Value = NullString(rs_OINV.Fields.Item(6).Value)
                ocreditMemo.UserFields.Fields.Item("U_M_Original_Inv").Value = NullString(rs_OINV.Fields.Item(7).Value)

                Dim STR_INV1 As String
                Dim rs_INV1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                STR_INV1 = "select ItemCode,Quantity,PriceAfVAT,Price from INV1 where DocEntry = " & base_entry & ""
                rs_INV1.DoQuery(STR_INV1)

                Dim baseline As Integer = 0
                rs_INV1.MoveFirst()
                Do While Not rs_INV1.EoF
                    Dim creditMemoLines As SAPbobsCOM.Document_Lines = ocreditMemo.Lines
                    Dim oSerial As SAPbobsCOM.SerialNumbers = creditMemoLines.SerialNumbers
                    DistNumber = rs_INV1.Fields.Item(0).Value
                    'add line ===========================
                    With creditMemoLines
                        .ItemCode = rs_INV1.Fields.Item(0).Value
                        .Quantity = rs_INV1.Fields.Item(1).Value
                        .UnitPrice = rs_INV1.Fields.Item(2).Value
                        .Price = rs_INV1.Fields.Item(3).Value

                        If obj_type = "DT" Then
                            basetype_serial = SAPbobsCOM.BoObjectTypes.oDownPayments
                            .BaseType = SAPbobsCOM.BoObjectTypes.oDownPayments
                            .BaseEntry = base_entry
                            .BaseLine = baseline
                        End If
                        If obj_type = "IN" Then
                            basetype_serial = SAPbobsCOM.BoObjectTypes.oInvoices
                            .BaseType = SAPbobsCOM.BoObjectTypes.oInvoices
                            .BaseEntry = base_entry
                            .BaseLine = baseline

                        End If
                    End With
                    Dim STR_OSRN As String
                    Dim sr_OSRN As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    STR_OSRN = "select b.DistNumber from SRI1 a " &
                                    "inner join OSRN b on a.SysSerial = b.SysNumber and a.ItemCode = b.ItemCode " &
                                    "where b.DistNumber = '" & DistNumber & "' AND a.BaseEntry = " & base_entry & " AND a.BaseType = " & basetype_serial & ""
                    sr_OSRN.DoQuery(STR_OSRN)
                    sr_OSRN.MoveFirst()
                    Do While Not sr_OSRN.EoF
                        oSerial.InternalSerialNumber = rs_OINV.Fields.Item(0).Value.ToString()
                        oSerial.Add()

                        sr_OSRN.MoveNext()
                    Loop
                    creditMemoLines.Add()
                    baseline += 1
                    rs_INV1.MoveNext()
                Loop

                lRetCode = ocreditMemo.Add()

                Dim docEntry As String = ""
                oCompany.GetNewObjectCode(docEntry)
                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrorCode, sErrMsg)
                    errormsge = sErrMsg
                End If

                'DIErrorHandler(String.Format("{0} Create Document", docEntry))
                AddErrorLog("A/R Credit Memo", rs_OINV.Fields.Item(8).Value.ToString(), lRetCode, docEntry, _dtARCreditMemoLog, True, skey)
                rs_OINV.MoveNext()
            Loop

            UpdateProcess("CNDN", "Idle")
        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try

        Return errormsge

    End Function
    'incoming payments
    Public Sub CreateIncomingPayment(ByVal h As DataTable, ByVal d As DataTable, ByVal t As DataTable, ByVal docnum As String, ByVal Docstatus As String, ByVal invoicedocEntry As Integer, ByVal typeinvoice As String)
        Try
            Dim sqlTrans As SqlTransaction = Nothing
            Dim sqlTrans1 As SqlTransaction = Nothing
            Dim cmd As New SqlCommand
            Dim dbreader As SqlDataReader
            UpdateProcess("Payment", "notIdle")
            For Each row_h As DataRow In h.Select("DocNum=" & docnum) 'loop header
                'For Each row_h As DataRow In h.Select() 'loop header

                If d.Select("docnum=" & row_h("DocNum")).Count > 0 Then

                    Dim oDoc As SAPbobsCOM.Payments = Nothing
                    Dim docEntry As String = ""

                    oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                    oDoc.CardCode = row_h("CardCode").ToString()
                    oDoc.DocDate = row_h("docdate")
                    oDoc.TaxDate = row_h("docdate")
                    oDoc.DocNum = row_h("DocNum")
                    'oDoc.CounterReference = ""

                    For Each row_d As DataRow In d.Select("docnum=" & row_h("DocNum")) 'add invoice
                        oDoc.Invoices.SetCurrentLine(0)
                        oDoc.Invoices.SumApplied = NullZero(row_d("PaymentValue"))
                        oDoc.Invoices.AppliedFC = NullZero(row_d("PaymentValue"))
                        If typeinvoice.ToUpper() = "DT" Then
                            oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_DownPayment
                        End If
                        If typeinvoice.ToUpper() = "IN" Then
                            oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice
                        End If
                        'oDoc.Invoices.DocEntry = row_d("invoiceID") 'invoicedocEntry
                        oDoc.Invoices.DocEntry = invoicedocEntry
                        oDoc.Invoices.Add()
                    Next

                    For Each row_t As DataRow In t.Select("docnum=" & row_h("DocNum")) 'add เงิน และแยกประเภทการจ่าย
                        If row_t("pmttype") = 1 Then 'cash
                            oDoc.CashSum = NullZero(row_t("CashSum"))
                            oDoc.CashAccount = NullString(row_t("CashAcct"))
                        Else 'credit card
                            oDoc.CreditCards.CreditCard = NullZero(row_t("CreditCard"))
                            oDoc.CreditCards.CardValidUntil = row_t("CardValid")
                            oDoc.CreditCards.CreditCardNumber = NullString(row_t("CrCardNum"))
                            oDoc.CreditCards.CreditSum = NullZero(row_t("CreditSum"))
                            oDoc.CreditCards.VoucherNum = NullString(row_t("VoucherNum"))
                            oDoc.CreditCards.PaymentMethodCode = 1
                            oDoc.CreditCards.Add()
                        End If
                    Next


                    lRetCode = oDoc.Add()

                    oCompany.GetNewObjectCode(docEntry)
                    oCompany.GetLastError(lErrorCode, sErrMsg)
                    'DIErrorHandler(String.Format("{0} Create Document", docEntry))
                    AddErrorLog("AddPayment", row_h("DocNum"), lRetCode, docEntry, _dtIncomingPaymentLog, True, row_h("skey"))

                    If lRetCode <> 0 Then
                        Dim WriteLog = m_sp_ais_WriteLog("Payment", row_h("skey"), "E", docEntry, "Add Payment", row_h("DocNum"), lRetCode, sErrMsg)
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

                    Else ' add WriteLog and update payment
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        m_sp_ais_WriteLog("Payment", row_h("skey"), "S", docEntry, "Add Payment", row_h("DocNum"), lRetCode, sErrMsg)

                    End If

                Else
                    ManualAddErrorLog("Payment", row_h("DocNum"), -10, "", _dtIncomingPaymentLog)
                End If

            Next
            UpdateProcess("Payment", "Idle")
        Catch ex As Exception

            ' MessageBox.Show(ex.Message)
        End Try
    End Sub

    'inventorytransferrequest
    Public Sub CreateInventoryTransferrequest(ByVal h As DataTable, ByVal l As DataTable)
        Try
            UpdateProcess("MovementRequest", "notIdle")
            Dim serial As String = ""
            For Each row_h As DataRow In h.Select() 'loop header

                If l.Select("DocNum='" & row_h("DocNum") & "'").Count > 0 Then
                    Dim oDoc As SAPbobsCOM.StockTransfer = Nothing
                    oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest)
                    oDoc.DocDate = row_h("DocDate")
                    oDoc.DueDate = row_h("DocDueDate")
                    oDoc.TaxDate = row_h("TaxDate")
                    oDoc.CardCode = NullString(row_h("CardCode"))
                    oDoc.FromWarehouse = NullString(row_h("FromWhs"))
                    oDoc.ToWarehouse = NullString(row_h("ToWhs"))
                    oDoc.CardName = row_h("CardName")
                    'oDoc.Comments = NullString(row_h("Comments"))
                    oDoc.UserFields.Fields.Item("U_F_POS_Ref").Value = row_h("DocNum")
                    oDoc.UserFields.Fields.Item("U_BP_Name").Value = row_h("CardName")
                    Dim docnum As String
                    For Each row_d As DataRow In l.Select("DocNum='" & row_h("DocNum") & "'") 'add Lines
                        'Dim inventoryline As SAPbobsCOM.Document_Lines = oDoc.Lines
                        'Dim inventorySerial As SAPbobsCOM.SerialNumbers = inventoryline.SerialNumbers

                        docnum = row_d("DocNum")

                        oDoc.Lines.ItemCode = row_d("ItemCode")
                        oDoc.Lines.Quantity = row_d("Quantity")

                        'oDoc.Lines.WarehouseCode = NullString(row_h("ToWhs"))
                        'oDoc.Lines.FromWarehouseCode = NullString(row_h("FromWhs"))


                        'For Each row_s As DataRow In s.Select("DocNum='" & row_d("DocNum") & "' and LineNum = " & row_d("LineNum")) 'add serial

                        '    serial = row_s("Serialnum")
                        '    If serial.Length > 0 Then
                        '        oDoc.Lines.SerialNumbers.SetCurrentLine(0)
                        '        oDoc.Lines.SerialNumbers.InternalSerialNumber = row_s("Serialnum")
                        '        oDoc.Lines.SerialNumbers.Quantity = row_s("Quantity")
                        '        oDoc.Lines.SerialNumbers.Add()
                        '    End If

                        'Next

                        oDoc.Lines.Add()
                    Next

                    Dim docEntry As String = ""

                    lRetCode = oDoc.Add()
                    oCompany.GetNewObjectCode(docEntry)
                    oCompany.GetLastError(lErrorCode, sErrMsg)
                    'DIErrorHandler(String.Format("{0} Create Document", docEntry))
                    If lRetCode <> 0 Then
                        AddErrorLog("InventoryTransferrequest", row_h("DocNum"), lRetCode, docEntry, _dtInventoryTransferrequestLog, True, 0)
                        m_sp_ais_WriteLog("InvtTRReqPOS", 0, "E", docEntry, "Add InventoryTransferrequest", row_h("DocNum"), lRetCode, sErrMsg)
                    Else
                        AddErrorLog("InventoryTransferrequest", row_h("DocNum"), lRetCode, docEntry, _dtInventoryTransferrequestLog, True, 0)
                        m_sp_ais_WriteLog("InvtTRReqPOS", 0, "S", docEntry, "Add InventoryTransferrequest", row_h("DocNum"), lRetCode, sErrMsg)
                        ' m_sp_ais_WriteLog(ByVal type As String, ByVal skey As Integer, ByVal LoadStatus As String, ByVal SAPRecID As String, ByVal Title As String, ByVal POSDocNum As String, ByVal ErrCode As String, ByVal ErrMsg As String)
                    End If

                Else
                    ManualAddErrorLog("InventoryTransferrequest", row_h("DocNum"), -10, "", _dtInventoryTransferrequestLog)
                End If

            Next

            UpdateProcess("MovementRequest", "Idle")
        Catch ex As Exception
            MsgBoxWrapper(String.Format("{0} Exception", ex))
            Throw ex
        End Try
    End Sub
    'inventorytransfer
    Public Sub CreateInventoryTransfer(ByVal h As DataTable, ByVal l As DataTable, ByVal s As DataTable)
        Try
            UpdateProcess("Movement", "notIdle")
            Dim serial As String = ""
            Dim CheckDocnum As DataRow()
            For Each row_h As DataRow In h.Select() 'loop header
                'Dim v_InvTransferRequestEntry As SAPbobsCOM.StockTransfer = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest)

                'oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                ' MessageBox.Show(l.Select("DocNum="&"RT0000000004").Count)
                If l.Select("DocNum='" & row_h("DocNum") & "'").Count > 0 Then
                    Dim oDoc As SAPbobsCOM.StockTransfer '= Nothing
                    oDoc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)

                    ' oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer

                    oDoc.DocDate = row_h("DocDate")
                    oDoc.TaxDate = row_h("TaxDate")
                    oDoc.CardCode = NullString(row_h("CardCode"))
                    oDoc.FromWarehouse = NullString(row_h("FromWhs"))
                    oDoc.ToWarehouse = NullString(row_h("ToWhs"))
                    oDoc.Comments = NullString(row_h("Comments"))
                    oDoc.UserFields.Fields.Item("U_F_POS_Ref").Value = row_h("DocNum")

                    For Each row_d As DataRow In l.Select("DocNum='" & row_h("DocNum") & "'") 'add Lines
                        'Dim inventoryline As SAPbobsCOM.Document_Lines = oDoc.Lines
                        'Dim inventorySerial As SAPbobsCOM.SerialNumbers = inventoryline.SerialNumbers
                        'CheckDocnum = s.Select("DocNum='" & row_d("DocNum") & "' and LineNum = " & row_d("LineNum"))
                        'If (CheckDocnum.Count > 0) Then
                        oDoc.Lines.ItemCode = row_d("ItemCode")
                        oDoc.Lines.Quantity = row_d("Quantity")
                        oDoc.Lines.WarehouseCode = NullString(row_h("ToWhs"))
                        oDoc.Lines.FromWarehouseCode = NullString(row_h("FromWhs"))
                        Dim i As Integer = 0
                        For Each row_s As DataRow In s.Select("DocNum='" & row_d("DocNum") & "' and LineNum = " & row_d("LineNum")) 'add serial

                            serial = row_s("Serialnum")
                            If serial.Length > 0 Then
                                'oDoc.Lines.SerialNumbers.SetCurrentLine(0)
                                oDoc.Lines.SerialNumbers.InternalSerialNumber = row_s("Serialnum")
                                oDoc.Lines.SerialNumbers.BaseLineNumber = row_s("baseline")
                                'oDoc.Lines.SerialNumbers.Quantity = row_s("Quantity")
                                oDoc.Lines.SerialNumbers.Add()
                            End If

                        Next
                        oDoc.Lines.Add()
                        'End If


                    Next

                    Dim docEntry As String = ""

                    lRetCode = oDoc.Add()
                    oCompany.GetNewObjectCode(docEntry)
                    oCompany.GetLastError(lErrorCode, sErrMsg)
                    'DIErrorHandler(String.Format("{0} Create Document", docEntry))
                    'AddErrorLog("InventoryTransfer", row_h("DocNum"), lRetCode, docEntry, _dtInventoryTransferLog, True, row_h("DocNum"))
                    If lRetCode <> 0 Then
                        AddErrorLog("InventoryTransfer", row_h("DocNum"), lRetCode, docEntry, _dtInventoryTransferLog, True, 0)
                        m_sp_ais_WriteLog("InvtTRSAP", 0, "E", docEntry, "Add InventoryTransfer", row_h("DocNum"), lRetCode, sErrMsg)
                    Else
                        AddErrorLog("InventoryTransfer", row_h("DocNum"), lRetCode, docEntry, _dtInventoryTransferLog, True, 0)
                        m_sp_ais_WriteLog("InvtTRSAP", 0, "S", docEntry, "Add InventoryTransfer", row_h("DocNum"), lRetCode, sErrMsg)
                    End If
                Else
                    ManualAddErrorLog("InventoryTransfer", row_h("DocNum"), -10, "", _dtInventoryTransferLog)
                End If

            Next

            UpdateProcess("Movement", "Idle")
        Catch ex As Exception
            MsgBoxWrapper(String.Format("{0} Exception", ex))
            Throw ex
        End Try
    End Sub

#End Region

#Region "Write Log"

    Sub AddErrorLog(Title As String, ByVal POSDucNum As String, ByVal lRetCode As Int64, ByVal docEntry As String, ByRef LogTable As DataTable, statuspast As Boolean, ByVal skey As String)
        Try
            Select Case lRetCode
                Case 0 'success
                    LogTable.Rows.Add({Title, DBNull.Value, "S", POSDucNum, docEntry, DBNull.Value, DBNull.Value, statuspast})

                    _fs = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & getdate_file, True)
                    _fs.WriteLine(Now.ToString)
                    _fs.WriteLine(Title & " success" & " skey=>" & skey & ", docnum=>" & POSDucNum)
                    _fs.WriteLine("-----------------------------------------------------")
                    _fs.Close()

                Case 888 'no record
                    LogTable.Rows.Add({Title, DBNull.Value, "F", POSDucNum, DBNull.Value, "888", "No record.", statuspast})

                    _fs = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & getdate_file, True)
                    _fs.WriteLine(Now.ToString)
                    _fs.WriteLine(Title & " Error [No record]" & " skey=>" & skey & ", docnum=>" & POSDucNum)
                    _fs.WriteLine("-----------------------------------------------------")
                    _fs.Close()
                Case Else 'error
                    oCompany.GetLastError(lErrorCode, sErrMsg)
                    LogTable.Rows.Add({Title, DBNull.Value, "F", POSDucNum, DBNull.Value, lErrorCode, sErrMsg, statuspast})

                    _fs = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & getdate_file, True)
                    _fs.WriteLine(Now.ToString)
                    _fs.WriteLine(Title & " Error [" & sErrMsg & "]" & " skey=>" & skey & ", docnum=>" & POSDucNum)
                    _fs.WriteLine("-----------------------------------------------------")
                    _fs.Close()
            End Select

        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

    Sub ManualAddErrorLog(Title As String, ByVal POSDucNum As String, ByVal lRetCode As Int64, ByVal docEntry As String, ByRef LogTable As DataTable)
        Try
            If lRetCode = -10 Then 'Error
                LogTable.Rows.Add({Title, DBNull.Value, "F", POSDucNum, DBNull.Value, -10, "Not found invoice."})
            End If
        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

    Sub WriteLog()
        Try
            Dim dtInterfaceLog As New DataTable
            Dim da As New SqlDataAdapter

            ConnectInterfaceDB()

            '============================= get schema =============================
            cmd = New SqlCommand
            With cmd
                .Connection = oConn
                .CommandTimeout = 0
                .CommandType = CommandType.Text
                .CommandText = "select * from Interface_Log where 1=2"
            End With

            da = New SqlDataAdapter
            da.SelectCommand = cmd
            da.Fill(dtInterfaceLog)
            '======================================================================


            If _dtInvoiceLog.Rows.Count > 0 Then
                For Each dr As DataRow In _dtInvoiceLog.Select
                    dr("TransDate") = _TransDate
                    dtInterfaceLog.ImportRow(dr)
                Next
            End If

            If _dtARCreditMemoLog.Rows.Count > 0 Then
                For Each dr As DataRow In _dtARCreditMemoLog.Select
                    dr("TransDate") = _TransDate
                    dtInterfaceLog.ImportRow(dr)
                Next
            End If

            If _dtIncomingPaymentLog.Rows.Count > 0 Then
                For Each dr As DataRow In _dtIncomingPaymentLog.Select
                    dr("TransDate") = _TransDate
                    dtInterfaceLog.ImportRow(dr)
                Next
            End If

            If _dtInventoryTransferLog.Rows.Count > 0 Then
                For Each dr As DataRow In _dtInventoryTransferLog.Select
                    dr("TransDate") = _TransDate
                    dtInterfaceLog.ImportRow(dr)
                Next
            End If

            If _dtInventoryTransferrequestLog.Rows.Count > 0 Then
                For Each dr As DataRow In _dtInventoryTransferrequestLog.Select
                    dr("TransDate") = _TransDate
                    dtInterfaceLog.ImportRow(dr)
                Next
            End If

            ''insert log table
            Dim sqlBulkCopy As SqlBulkCopy = New SqlBulkCopy(oConn)
            sqlBulkCopy.DestinationTableName = "Interface_Log"
            sqlBulkCopy.BulkCopyTimeout = 0
            sqlBulkCopy.BatchSize = 1000
            sqlBulkCopy.ColumnMappings.Add("Title","Title")
            sqlBulkCopy.ColumnMappings.Add("TransDate", "TransDate")
            sqlBulkCopy.ColumnMappings.Add("Status", "Status")
            sqlBulkCopy.ColumnMappings.Add("POSDocNum", "POSDocNum")
            sqlBulkCopy.ColumnMappings.Add("docEntry", "docEntry")
            sqlBulkCopy.ColumnMappings.Add("ErrCode", "ErrCode")
            sqlBulkCopy.ColumnMappings.Add("ErrMsg", "ErrMsg")
            sqlBulkCopy.WriteToServer(dtInterfaceLog)

            sqlBulkCopy.Close()

            CloseConnect()

            'UpdateInvoice()
            'UpdateIncomingPayment()
            'UpdateInventoryTransfer()
            'UpdateInventoryTransferrequest()

        Catch ex As Exception

            Dim fs As System.IO.StreamWriter
            fs = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\WriteErrorLog.txt", True)
            fs.WriteLine(Now.ToString)
            fs.WriteLine("Write log error, Please contact administration.")
            fs.WriteLine("Error Message : " & ex.Message)
            fs.WriteLine("-----------------------------------------------------")
            fs.Close()

            CloseConnect() 'sql close connection

        End Try
    End Sub

    Sub UpdateInvoice()
        Dim sqlTrans As SqlTransaction = Nothing
        Try
            ConnectInterfaceDB()

            sqlTrans = oConn.BeginTransaction("UpdateInvoice")

            If _dtInvoiceLog.Rows.Count > 0 Then

                For Each dr As DataRow In _dtInvoiceLog.Select
                    If dr("statuspast") = True Then
                        cmd = New SqlCommand
                        With cmd
                            .Parameters.Clear()
                            .Connection = oConn
                            .Transaction = sqlTrans
                            .CommandTimeout = 0
                            .CommandType = CommandType.Text
                            .CommandText = "update InvHeader set LoadtoSAPDate = @LoadtoSAPDate , LoadStatus = @LoadStatus, SAPRecID = @SAPRecID where DocNum = @DocNum"
                            .Parameters.Add("@DocNum", SqlDbType.Int).Value = dr("POSDocNum")
                            .Parameters.Add("@LoadtoSAPDate", SqlDbType.DateTime).Value = Now
                            .Parameters.Add("@LoadStatus", SqlDbType.NVarChar, 1).Value = IIf(dr("Status") = "S", "Y", "E")
                            .Parameters.Add("@SAPRecID", SqlDbType.Int).Value = dr("docEntry")
                        End With

                        cmd.ExecuteNonQuery()
                    End If

                Next

            End If

            sqlTrans.Commit()

            CloseConnect()

        Catch ex As Exception

            _fs = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & getdate_file, True)
            _fs.WriteLine(Now.ToString)
            _fs.WriteLine("UpdateInvoice")
            _fs.WriteLine("Error Message : " & ex.Message)
            _fs.WriteLine("-----------------------------------------------------")
            _fs.Close()

            Try
                sqlTrans.Rollback()
            Catch ex2 As Exception

            End Try
            CloseConnect()
        End Try
    End Sub

    Sub UpdateIncomingPayment()
        Dim sqlTrans As SqlTransaction = Nothing
        Dim cmd As New SqlCommand
        Try
            ConnectInterfaceDB()

            sqlTrans = oConn.BeginTransaction("UpdateIncomingPayment")

            If _dtIncomingPaymentLog.Rows.Count > 0 Then

                For Each dr As DataRow In _dtIncomingPaymentLog.Select
                    If dr("statuspast") = True Then
                        cmd = New SqlCommand
                        With cmd
                            .Parameters.Clear()
                            .Connection = oConn
                            .Transaction = sqlTrans
                            .CommandTimeout = 0
                            .CommandType = CommandType.Text
                            .CommandText = "update Payment set LoadtoSAPDate = @LoadtoSAPDate , LoadStatus = @LoadStatus, SAPRecID = @SAPRecID where DocNum = @DocNum"
                            .Parameters.Add("@DocNum", SqlDbType.Int).Value = dr("POSDocNum")
                            .Parameters.Add("@LoadtoSAPDate", SqlDbType.DateTime).Value = Now
                            .Parameters.Add("@LoadStatus", SqlDbType.NVarChar, 1).Value = IIf(dr("Status") = "S", "Y", "E")
                            .Parameters.Add("@SAPRecID", SqlDbType.Int).Value = dr("docEntry")
                        End With

                        cmd.ExecuteNonQuery()
                    End If


                Next

            End If

            sqlTrans.Commit()

            CloseConnect()

        Catch ex As Exception
            Try
                sqlTrans.Rollback()
            Catch ex2 As Exception
            End Try
            CloseConnect()
        End Try
    End Sub


    Sub UpdateInventoryTransfer()
        Dim sqlTrans As SqlTransaction = Nothing
        Dim cmd As New SqlCommand
        Try
            ConnectInterfaceDB()

            sqlTrans = oConn.BeginTransaction("UpdateInventoryTransfer")

            If _dtInventoryTransferLog.Rows.Count > 0 Then

                For Each dr As DataRow In _dtInventoryTransferLog.Select

                    cmd = New SqlCommand
                    With cmd
                        .Parameters.Clear()
                        .Connection = oConn
                        .Transaction = sqlTrans
                        .CommandTimeout = 0
                        .CommandType = CommandType.Text
                        .CommandText = "update InvtTRSAP set LoadtoSAPDate = @LoadtoSAPDate , LoadStatus = @LoadStatus, SAPRecID = @SAPRecID where DocNum = @DocNum"
                        .Parameters.Add("@DocNum", SqlDbType.NVarChar).Value = dr("POSDocNum")
                        .Parameters.Add("@LoadtoSAPDate", SqlDbType.DateTime).Value = Now
                        .Parameters.Add("@LoadStatus", SqlDbType.NVarChar, 1).Value = IIf(dr("Status") = "S", "Y", "E")
                        .Parameters.Add("@SAPRecID", SqlDbType.Int).Value = dr("docEntry")
                    End With

                    cmd.ExecuteNonQuery()

                Next

            End If

            sqlTrans.Commit()

            CloseConnect()

        Catch ex As Exception
            Try
                sqlTrans.Rollback()
            Catch ex2 As Exception
            End Try
            CloseConnect()
        End Try
    End Sub

    Sub UpdateInventoryTransferrequest()
        Dim sqlTrans As SqlTransaction = Nothing
        Dim cmd As New SqlCommand
        Try
            ConnectInterfaceDB()

            sqlTrans = oConn.BeginTransaction("UpdateInventoryTransferrequest")

            If _dtInventoryTransferrequestLog.Rows.Count > 0 Then

                For Each dr As DataRow In _dtInventoryTransferrequestLog.Select

                    cmd = New SqlCommand
                    With cmd
                        .Parameters.Clear()
                        .Connection = oConn
                        .Transaction = sqlTrans
                        .CommandTimeout = 0
                        .CommandType = CommandType.Text
                        .CommandText = "update InvtTRReqPOS set LoadtoSAPDate = @LoadtoSAPDate , LoadStatus = @LoadStatus, SAPRecID = @SAPRecID where DocNum = @DocNum"
                        .Parameters.Add("@DocNum", SqlDbType.NVarChar).Value = dr("POSDocNum")
                        .Parameters.Add("@LoadtoSAPDate", SqlDbType.DateTime).Value = Now
                        .Parameters.Add("@LoadStatus", SqlDbType.NVarChar, 1).Value = IIf(dr("Status") = "S", "Y", "E")
                        .Parameters.Add("@SAPRecID", SqlDbType.Int).Value = dr("docEntry")
                    End With

                    cmd.ExecuteNonQuery()

                Next

            End If

            sqlTrans.Commit()

            CloseConnect()

        Catch ex As Exception
            Try
                sqlTrans.Rollback()
            Catch ex2 As Exception
            End Try
            CloseConnect()
        End Try
    End Sub

    Sub UpdateProcess(ByVal operation As String, ByVal status As String)
        Dim sqlTrans As SqlTransaction = Nothing
        Dim cmd As New SqlCommand
        Try
            ConnectInterfaceDB()

            sqlTrans = oConn.BeginTransaction("UpdateProcess")

            cmd = New SqlCommand
            With cmd
                .Parameters.Clear()
                .Connection = oConn
                .Transaction = sqlTrans
                .CommandTimeout = 0
                .CommandType = CommandType.Text
                .CommandText = "update AIS_Monitoring set AMT_LogMsg = @AMT_LogMsg , AMT_LogDate = @AMT_LogDate where AMT_Setting = @AMT_Setting"
                .Parameters.Add("@AMT_LogMsg", SqlDbType.NVarChar).Value = status
                .Parameters.Add("@AMT_LogDate", SqlDbType.DateTime).Value = Now
                .Parameters.Add("@AMT_Setting", SqlDbType.NVarChar).Value = operation

            End With
            'If cmd.Connection.State <> ConnectionState.Open Then
            '    cmd.Connection.Open()

            'End If

            cmd.ExecuteNonQuery()
            sqlTrans.Commit()

            CloseConnect()

        Catch ex As Exception
            Try
                sqlTrans.Rollback()
            Catch ex2 As Exception
            End Try
            CloseConnect()
        End Try
    End Sub

    Sub CheckDocnum()
        Dim sqlTrans As SqlTransaction = Nothing
        Dim cmd As New SqlCommand

        Try
            ConnectInterfaceDB()

            sqlTrans = oConn.BeginTransaction("CheckDocnum")

            cmd = New SqlCommand
            With cmd
                .Parameters.Clear()
                .Connection = oConn
                .Transaction = sqlTrans
                .CommandTimeout = 0
                .CommandType = CommandType.Text
                .CommandText = "select * from InvtTRReqPOS where LoadStatus = 'Y' "
            End With
            da_CheckDocnum.SelectCommand = cmd
            da_CheckDocnum.Fill(ds_CheckDocnum)

            If ds_CheckDocnum.Tables.Count > 0 Then
                dt_CheckDocnum = ds_CheckDocnum.Tables(0).Copy
            End If
            'If cmd.Connection.State <> ConnectionState.Open Then
            '    cmd.Connection.Open()

            'End If

            cmd.ExecuteNonQuery()
            sqlTrans.Commit()

            CloseConnect()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Try
                sqlTrans.Rollback()
            Catch ex2 As Exception
                MessageBox.Show(ex2.Message)
            End Try
            CloseConnect()
        End Try
    End Sub


#End Region


End Class
