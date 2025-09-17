Option Explicit On
Option Strict On

Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.UI.Hub
' 필요하면 네임스페이스 맞춰 사용:
'Imports KKY_Tool_Revit.UI.Hub
' 또는 프로젝트에서 UiBridgeExternalEvent가 같은 루트 네임스페이스면 Imports 없이 사용 가능

<Transaction(TransactionMode.Manual)>
<Regeneration(RegenerationOption.Manual)>
Public Class DuplicateExport
    Implements IExternalCommand

    Public Function Execute(commandData As ExternalCommandData,
                            ByRef message As String,
                            elements As ElementSet) As Result Implements IExternalCommand.Execute
        Try
            ' 허브 창을 띄운다(단일 인스턴스 강제까지는 하지 않고, 안정적으로 실행되도록 단순화)
            Dim wnd As New HubHostWindow()

            ' 브릿지 초기화(웹 ↔ Revit 메시지 라우팅)
            ' 네임스페이스가 KKY_Tool_Revit.UI.Hub 인 경우: KKY_Tool_Revit.UI.Hub.UiBridgeExternalEvent.Initialize(wnd)
            UiBridgeExternalEvent.Initialize(wnd)

            wnd.Show()
            wnd.Activate()
            Return Result.Succeeded

        Catch ex As Exception
            message = ex.Message
            Return Result.Failed
        End Try
    End Function

End Class
