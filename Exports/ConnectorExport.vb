Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic

Namespace Exports
    Public Module ConnectorExport
        ' rows: ȣ���ڰ� ��Ű���� ���� �ѱ�
        ' ���� ��Ű��: ElementId | ī�װ� | tol | unit | param | ��
        Public Function Save(rows As IEnumerable(Of IEnumerable(Of Object))) As String
            Dim headers = New String() {"ElementId", "ī�װ�", "tol", "unit", "param", "��"}
            Return Infrastructure.ExcelCore.PickAndSaveXlsx(
              $"Connector_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
              "Connector",
              headers,
              rows
            )
        End Function
    End Module
End Namespace
