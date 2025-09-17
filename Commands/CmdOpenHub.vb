Option Explicit On
Option Strict On

Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.UI.Hub
' �ʿ��ϸ� ���ӽ����̽� ���� ���:
'Imports KKY_Tool_Revit.UI.Hub
' �Ǵ� ������Ʈ���� UiBridgeExternalEvent�� ���� ��Ʈ ���ӽ����̽��� Imports ���� ��� ����

<Transaction(TransactionMode.Manual)>
<Regeneration(RegenerationOption.Manual)>
Public Class DuplicateExport
    Implements IExternalCommand

    Public Function Execute(commandData As ExternalCommandData,
                            ByRef message As String,
                            elements As ElementSet) As Result Implements IExternalCommand.Execute
        Try
            ' ��� â�� ����(���� �ν��Ͻ� ���������� ���� �ʰ�, ���������� ����ǵ��� �ܼ�ȭ)
            Dim wnd As New HubHostWindow()

            ' �긴�� �ʱ�ȭ(�� �� Revit �޽��� �����)
            ' ���ӽ����̽��� KKY_Tool_Revit.UI.Hub �� ���: KKY_Tool_Revit.UI.Hub.UiBridgeExternalEvent.Initialize(wnd)
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
