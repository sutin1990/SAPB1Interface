Imports SAPbouiCOM
Imports SAPbobsCOM
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration


Module AddOn
    Public oCompany As SAPbobsCOM.Company = Nothing
    Public lRetCode As Integer
    Public lErrCode As Integer = 0
    Public lErrorCode As String
    Public sErrMsg As String = String.Empty

    Public WithEvents oApp As SAPbouiCOM.Application

    Public SboGuiApi As SAPbouiCOM.SboGuiApi
    Public sConnectionString As String

    Public Event AppEvent(ByVal EventType As BoAppEventTypes)
    Public WithEvents oApplication As SAPbouiCOM.Application = Nothing
    Public oConn As SqlConnection

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles oApp.AppEvent

        Select Case EventType

            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                oApp.MessageBox("A Shut Down Event has been caught")
                System.Windows.Forms.Application.Exit()
        End Select
    End Sub

    Sub Main()
        '_SAPServerName = System.Configuration.ConfigurationSettings.AppSettings.Get(0)
        'Connect()


        'RegisterInterfaceTTMS()
        'CreateMenus()

        'Dim frm As frmSAPInterfaceTTMS = New frmSAPInterfaceTTMS
        'frm.ShowDialog()

        'System.Windows.Forms.Application.Run()



        Dim c As New SAPB1Interface
        c.Execute()

    End Sub

    Sub Connect()
        Try
            'GetConnection()

            'Dim uiAPI As SAPbouiCOM.SboGuiApi = New SAPbouiCOM.SboGuiApi

            ' 2. Add On Identifier / License Check
            ' ConnStr in Project Properties > Debug > CommandLineArg
            'Dim connStr As String = Environment.GetCommandLineArgs.GetValue(1)



            '' 3. Connect to UI API
            'uiAPI.Connect(connStr)

            '' 4. Get the active B1 Application Object
            'oApp = uiAPI.GetApplication

            'oApplication = uiAPI.GetApplication
            'oApp4MenuEvent = oApplication
            'oApp4ItemEvent = oApplication

            oCompany = New SAPbobsCOM.Company
            oCompany.Server = ConfigurationManager.AppSettings.Item("SAPServerName") '"TECHNB15"
            'oCompany.Server = oApp.Company.ServerName
            oCompany.CompanyDB = ConfigurationManager.AppSettings.Item("SAPDbName") '"IMED_GB"
            'oCompany.CompanyDB = oApp.Company.DatabaseName
            oCompany.UserName = ConfigurationManager.AppSettings.Item("SAPUserName") '"manager"
            'oCompany.UserName = oApp.Company.UserName
            oCompany.Password = ConfigurationManager.AppSettings.Item("SAPUserPassword") '"1234"
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016
            oCompany.DbUserName = ConfigurationManager.AppSettings.Item("SAPDbUserName") '"sa"
            oCompany.DbPassword = ConfigurationManager.AppSettings.Item("SAPDbPassword") '"P@ssw0rd"
            'oCompany.LicenseServer = oApp.Company.ServerName
            oCompany.LicenseServer = ConfigurationManager.AppSettings.Item("SAPLicenseServer") '"TECHNB15"
            lRetCode = oCompany.Connect
            MsgBoxWrapper(lRetCode)

            If lRetCode = 0 Then

            Else

            End If


            ''oCompany = New SAPbobsCOM.Company
            'Dim sCookie As String = oCompany.GetContextCookie()
            '' 3. Get the Conn Info for DI Company from UI App
            'Dim connContext As String =
            '    oApp.Company.GetConnectionContext(sCookie)
            '' 4. Set Conn to DI Company
            'oCompany.SetSboLoginContext(connContext)
            'lRetCode = oCompany.Connect

            'SboGuiApi = New SAPbouiCOM.SboGuiApi

            ''sConnectionString = Environment.GetCommandLineArgs.GetValue(1)

            'If Environment.GetCommandLineArgs.Length > 1 Then
            '    sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
            'Else
            '    sConnectionString = Environment.GetCommandLineArgs.GetValue(0)
            'End If

            'SboGuiApi.Connect(sConnectionString)
            'oApp = SboGuiApi.GetApplication()

            'oCompany = New SAPbobsCOM.Company
            'oCompany = oApp.Company.GetDICompany

            'Dim retVal As Integer

            'oCompany = New SAPbobsCOM.Company
            ''oCompany.Server = "SSPNB22"
            'oCompany.Server = oApp.Company.ServerName
            ''oCompany.CompanyDB = "intercompany1"
            'oCompany.CompanyDB = oApp.Company.DatabaseName
            ''oCompany.UserName = "manager"
            'oCompany.UserName = oApp.Company.UserName
            'oCompany.Password = "1234"
            'oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
            'oCompany.DbUserName = "sa"
            'oCompany.DbPassword = "P@ssw0rd"
            'oCompany.LicenseServer = oApp.Company.ServerName
            ''oCompany.LicenseServer = "SSPNB22"

            'retVal = oCompany.Connect

            'oApp.MessageBox("สวัสดีชาว Thai Localization", 1, "Continue", "Cancel")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub ConnectInterfaceDB()
        Try
            Dim ConStr As String = ""
            ConStr = "Server=" & ConfigurationManager.AppSettings.Item("DbServerName") & ";"
            ConStr &= "Database=" & ConfigurationManager.AppSettings.Item("DbName") & ";"
            ConStr &= "User Id=" & ConfigurationManager.AppSettings.Item("DbUserName") & ";"
            ConStr &= "Password=" & ConfigurationManager.AppSettings.Item("DbUserPassword") & ";"


            'MessageBox.Show(ConStr)
            oConn = New SqlConnection(ConStr)
            If oConn.State <> ConnectionState.Closed Then oConn.Close()
            oConn.Open()

        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

    Public Sub CloseConnect()
        Try
            If oConn.State <> ConnectionState.Closed Then oConn.Close()
        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

    Sub MsgBoxWrapper(ByVal msg As String,
                      Optional ByVal msgBoxtype As MsgBoxType = MsgBoxType.WinMsgBox)
        If oApplication IsNot Nothing Then
            'UIAPI Initialized
            If msgBoxtype = AddOn.MsgBoxType.B1MsgBox Then
                'B1 Msg Box
                oApplication.MessageBox(msg)
            Else
                'B1 Status Bar
                oApplication.SetStatusBarMessage(msg,
                                                   SAPbouiCOM.BoMessageTime.bmt_Medium,
                                                   False)
            End If
        Else
            'UIAPI not Initialized
            MsgBox(msg)
        End If
    End Sub

    Enum MsgBoxType
        WinMsgBox = 0
        B1MsgBox = 1
        B1StatusBar = 2
    End Enum

    Sub DIErrorHandler(ByVal operation As String)
        Dim msg As String = String.Format("{0} succeeded", operation)

        If lRetCode <> 0 Then
            'Error
            oCompany.GetLastError(lErrorCode, sErrMsg)
            msg = String.Format("{0} operation failed. Error Code {1}, Error Msg: {2}", operation, lErrorCode, sErrMsg)
        Else
            'Success
        End If
        MsgBoxWrapper(msg)
    End Sub

    Public Function NullString(ByVal s As Object)
        Try
            NullString = IIf(s Is DBNull.Value, "", s)
        Catch ex As Exception
            NullString = ""
        End Try
    End Function

    Public Function NullZero(ByVal o As Object)
        Try
            NullZero = IIf(o Is DBNull.Value, 0, o)
        Catch ex As Exception
            NullZero = 0
        End Try
    End Function

End Module
