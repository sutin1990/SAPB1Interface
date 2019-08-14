Module EventHandlerMenu
    Public WithEvents oApp4MenuEvent As SAPbouiCOM.Application

    Sub MenuEventHandler(ByRef pVal As SAPbouiCOM.MenuEvent,
                         ByRef BubbleEvent As Boolean) Handles oApp4MenuEvent.MenuEvent
        Try
            'If pVal.MenuUID.Equals("D_TTMS") Then
            '    If pVal.BeforeAction Then
            '        ''Create PQC Form
            '        'CreateInterfaceTTMS()
            '        Dim frm As frmSAPInterfaceTTMS = New frmSAPInterfaceTTMS
            '        frm.ShowDialog()
            '        BubbleEvent = False
            '    End If

            'Else

            'End If
        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

    Function CreateInterfaceTTMS()
        Dim oForm As SAPbouiCOM.Form = Nothing
        Try
            'oForm = CreateFormViaXmlFile("InterfaceTTMS.xml")
            '' Form | Series ComboBox UID | Header Table | DocNum UID
            'ManageSeries(oForm, "Item_3", "@TH_OPQC", "Item_4")

            'Try
            '    'oEditPostingDate = oForm.Items.Item("22_U_E").Specific
            'Catch ex As Exception

            'End Try
            '' Alternative Approach - Localization / Language specific
            ''1. Detect Language
            ''2. Load Language specific form

            ''Dim oForm As SAPbouiCOM.Form = Nothing
            ''If oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Chinese Then
            ''    oForm = CreateFormViaXmlFile("UDO_UDO_TH_OPQC.xml")
            ''else
            '' 'ToDo
            ''End If
            'Return oForm
        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try

        Return oForm
    End Function
End Module
