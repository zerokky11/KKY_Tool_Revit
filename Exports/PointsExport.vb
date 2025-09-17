Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic

Namespace Exports
    Public Module PointsExport
        Public Function Save(headers As IEnumerable(Of String),
                             rows As IEnumerable(Of IEnumerable(Of Object))) As String
            Return Infrastructure.ExcelCore.PickAndSaveXlsx(
              $"Points_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
              "Points",
              headers,
              rows
            )
        End Function
    End Module
End Namespace
