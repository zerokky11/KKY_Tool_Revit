Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic

Namespace Exports
    Public Module ConnectorExport
        ' rows: 호출자가 스키마에 맞춰 넘김
        ' 예시 스키마: ElementId | 카테고리 | tol | unit | param | 값
        Public Function Save(rows As IEnumerable(Of IEnumerable(Of Object))) As String
            Dim headers = New String() {"ElementId", "카테고리", "tol", "unit", "param", "값"}
            Return Infrastructure.ExcelCore.PickAndSaveXlsx(
              $"Connector_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
              "Connector",
              headers,
              rows
            )
        End Function
    End Module
End Namespace
