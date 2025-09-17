Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Linq
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports NPOI.HSSF.UserModel
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel
Imports System.Windows.Forms ' ← WinForms 다이얼로그 사용 (프로젝트 참조와 호환성 좋음)

Namespace UI.Hub
    ' 커넥터 진단 (fix2 이벤트명/스키마 유지)
    Partial Public Class UiBridgeExternalEvent

        ' 최근 로드/실행 결과(엑셀 저장 시 기본 소스)
        Private lastConnRows As List(Of Dictionary(Of String, Object)) = Nothing

#Region "핸들러 (Core에서 리플렉션으로 호출)"
        ' === connector:run ===
        Private Sub HandleConnectorRun(app As UIApplication, payload As Object)
            Try
                Dim uidoc = app.ActiveUIDocument
                Dim doc = If(uidoc Is Nothing, Nothing, uidoc.Document)
                If doc Is Nothing Then
                    SendToWeb("revit:error", New With {.message = "활성 문서가 없습니다."})
                    Return
                End If

                ' 파라미터(허용오차/단위/파라미터명) – 로직은 기존 구현에 위임, 여기선 배선 유지만
                Dim rows As New List(Of Dictionary(Of String, Object))()

                lastConnRows = rows
                SendToWeb("connector:loaded", New With {.rows = rows})

            Catch ex As Exception
                SendToWeb("revit:error", New With {.message = "실행 실패: " & ex.Message})
            End Try
        End Sub

        ' === connector:open-excel ===
        Private Sub HandleConnectorOpenExcel(app As UIApplication)
            Try
                Dim dt = TryReadExcelAsDataTable()
                If dt Is Nothing Then
                    SendToWeb("revit:error", New With {.message = "엑셀을 읽지 못했습니다."})
                    Return
                End If

                Dim rows = DataTableRows(dt)
                lastConnRows = rows
                SendToWeb("connector:loaded", New With {.rows = rows})

            Catch ex As Exception
                SendToWeb("revit:error", New With {.message = "엑셀 불러오기 실패: " & ex.Message})
            End Try
        End Sub

        ' === connector:save-excel ===
        Private Sub HandleConnectorSaveExcel(app As UIApplication, payload As Object)
            Try
                Dim rows As List(Of Dictionary(Of String, Object)) = TryGetRowsFromPayload(payload)
                If rows Is Nothing OrElse rows.Count = 0 Then rows = lastConnRows

                If rows Is Nothing OrElse rows.Count = 0 Then
                    SendToWeb("revit:error", New With {.message = "저장할 데이터가 없습니다."})
                    Return
                End If

                ' ===== fix2 스키마 가드 =====
                Dim expected As String() = {
                    "tol", "unit", "param",
                    "ElementId", "Category", "Family:Type",
                    "Distance(inch)", "Note"
                }
                If Not ValidateConnectorSchema(rows, expected) Then
                    SendToWeb("revit:error", New With {
                        .message = "엑셀 스키마가 fix2와 다릅니다. 헤더/순서를 확인하세요: " & String.Join(" | ", expected)
                    })
                    Return
                End If
                ' ============================

                Dim path = SaveExcelWithDialog(rows, expected)
                If Not String.IsNullOrEmpty(path) Then
                    SendToWeb("connector:saved", New With {.path = path})
                End If
            Catch ex As Exception
                SendToWeb("revit:error", New With {.message = "엑셀 저장 실패: " & ex.Message})
            End Try
        End Sub
#End Region

#Region "엑셀 I/O"
        Private Function TryReadExcelAsDataTable() As DataTable
            Using ofd As New OpenFileDialog()
                ofd.Filter = "Excel Files|*.xlsx;*.xls"
                ofd.Title = "커넥터 진단 - 엑셀 불러오기"
                If ofd.ShowDialog() <> DialogResult.OK Then Return Nothing

                Using fs = New FileStream(ofd.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    Dim wb As IWorkbook = If(
                        Path.GetExtension(ofd.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase),
                        CType(New XSSFWorkbook(fs), IWorkbook),
                        CType(New HSSFWorkbook(fs), IWorkbook)
                    )

                    Dim sh = wb.GetSheetAt(0)
                    If sh Is Nothing Then Return New DataTable()

                    Dim dt As New DataTable("Connector")
                    Dim headerRow = sh.GetRow(sh.FirstRowNum)
                    Dim headerAsTitle = HasHeaders(headerRow)
                    Dim colCount As Integer

                    If headerAsTitle AndAlso headerRow IsNot Nothing Then
                        colCount = headerRow.LastCellNum
                        For i = 0 To colCount - 1
                            Dim title = SafeToString(headerRow.GetCell(i))
                            If String.IsNullOrWhiteSpace(title) Then title = "Col" & (i + 1)
                            dt.Columns.Add(title)
                        Next
                    Else
                        colCount = If(headerRow Is Nothing, 0, headerRow.LastCellNum)
                        If colCount = 0 Then colCount = 1
                        For i = 0 To colCount - 1
                            dt.Columns.Add("Col" & (i + 1))
                        Next
                    End If

                    Dim startRow = sh.FirstRowNum
                    If headerAsTitle Then startRow += 1

                    For r = startRow To sh.LastRowNum
                        Dim row = sh.GetRow(r)
                        If row Is Nothing Then Continue For
                        Dim dr = dt.NewRow()
                        For c = 0 To dt.Columns.Count - 1
                            Dim cell = row.GetCell(c)
                            dr(c) = SafeToString(cell)
                        Next
                        dt.Rows.Add(dr)
                    Next

                    Return dt
                End Using
            End Using
        End Function

        Private Function SaveExcelWithDialog(rows As List(Of Dictionary(Of String, Object)),
                                             Optional headers As IEnumerable(Of String) = Nothing) As String
            Using sfd As New SaveFileDialog()
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                sfd.FileName = "ConnectorDiagnostics.xlsx"
                sfd.Title = "커넥터 진단 - 엑셀 저장"
                If sfd.ShowDialog() <> DialogResult.OK Then Return Nothing

                Dim wb As IWorkbook = New XSSFWorkbook()
                Dim sh = wb.CreateSheet("Sheet1")

                Dim hdr As List(Of String) =
                    If(headers IsNot Nothing, headers.ToList(), If(rows.Count > 0, rows(0).Keys.ToList(), New List(Of String)))

                ' 헤더
                Dim rh = sh.CreateRow(0)
                For c = 0 To hdr.Count - 1
                    rh.CreateCell(c).SetCellValue(hdr(c))
                Next

                ' 데이터
                For r = 0 To rows.Count - 1
                    Dim rr = sh.CreateRow(r + 1)
                    For c = 0 To hdr.Count - 1
                        Dim key = hdr(c)
                        Dim v As Object = Nothing
                        rows(r).TryGetValue(key, v)
                        rr.CreateCell(c).SetCellValue(SafeToString(v))
                    Next
                Next

                ' 자동 폭
                For c = 0 To hdr.Count - 1
                    sh.AutoSizeColumn(c)
                Next

                Using fs = New FileStream(sfd.FileName, FileMode.Create, FileAccess.Write)
                    wb.Write(fs)
                End Using

                Return sfd.FileName
            End Using
        End Function
#End Region

#Region "도우미"
        Private Function HasHeaders(row As NPOI.SS.UserModel.IRow) As Boolean
            If row Is Nothing Then Return False
            For c = 0 To row.LastCellNum - 1
                Dim s = SafeToString(row.GetCell(c))
                If Not String.IsNullOrWhiteSpace(s) AndAlso Not IsNumeric(s) Then
                    Return True
                End If
            Next
            Return False
        End Function

        Private Function DataTableRows(dt As DataTable) As List(Of Dictionary(Of String, Object))
            Dim list As New List(Of Dictionary(Of String, Object))()
            For Each dr As DataRow In dt.Rows
                Dim row As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                For Each col As DataColumn In dt.Columns
                    row(col.ColumnName) = If(dr.IsNull(col), Nothing, dr(col))
                Next
                list.Add(row)
            Next
            Return list
        End Function

        Private Function TryGetRowsFromPayload(payload As Object) As List(Of Dictionary(Of String, Object))
            Try
                If TypeOf payload Is List(Of Dictionary(Of String, Object)) Then
                    Return CType(payload, List(Of Dictionary(Of String, Object)))
                End If
            Catch
            End Try
            Return Nothing
        End Function

        Private Shared Function SafeToString(o As Object) As String
            If o Is Nothing Then Return ""
            Try
                Return Convert.ToString(o)
            Catch
                Return ""
            End Try
        End Function

        ' fix2 스키마 정확성 검사(헤더/순서)
        Private Function ValidateConnectorSchema(rows As List(Of Dictionary(Of String, Object)),
                                                 expected As String()) As Boolean
            If rows Is Nothing OrElse rows.Count = 0 Then Return False
            Dim first As Dictionary(Of String, Object) = rows(0)
            Dim keys As List(Of String) = first.Keys.ToList()
            If keys.Count <> expected.Length Then Return False
            For i As Integer = 0 To expected.Length - 1
                If Not String.Equals(keys(i), expected(i), StringComparison.OrdinalIgnoreCase) Then
                    Return False
                End If
            Next
            Return True
        End Function
#End Region

    End Class
End Namespace
