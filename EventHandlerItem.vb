Module EventHandlerItem
    Public WithEvents oApp4ItemEvent As SAPbouiCOM.Application

    Sub ItemEventHandler(ByVal FormUID As String,
                        ByRef pVal As SAPbouiCOM.ItemEvent,
                        ByRef BubbleEvent As Boolean) Handles oApp4ItemEvent.ItemEvent
        Try
            'Define FormUID > FormTypeEx of your Form
            'ToDo: Manage other forms

            If pVal.FormTypeEx.Equals("UDO_FT_UDO_TH_OPQC") Then
                'EventHandlerForm_OPQC.UDO_TH_OPQC_ItemEventHandler(FormUID, pVal, BubbleEvent)
            Else
            End If
        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try

    End Sub

End Module
