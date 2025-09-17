Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports Autodesk.Revit.UI
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Namespace UI.Hub
    Partial Public Class UiBridgeExternalEvent

        ' ========== Export: 폴더 선택 ==========
        Private Sub HandleExportBrowse()
            Using dlg As New System.Windows.Forms.FolderBrowserDialog()
                Dim r = dlg.ShowDialog()
                If r = System.Windows.Forms.DialogResult.OK Then
                    Dim files As String() = Directory.GetFiles(dlg.SelectedPath, "*.rvt", SearchOption.TopDirectoryOnly)
                    _host?.SendToWeb("export:files", New With {.files = files})
                End If
            End Using
        End Sub

        ' ========== Export: 미리보기 ==========
        Private Sub HandleExportPreview(app As UIApplication, payload As Dictionary(Of String, Object))
            Try
                Dim files = ExtractStringList(payload, "files")
                Dim rows = TryCallExportPointsService(app, files)
                If rows Is Nothing Then
                    _host?.SendToWeb("revit:error", New With {.message = "Export Points 서비스가 준비되지 않았습니다."})
                    _host?.SendToWeb("export:previewed", New With {.rows = New List(Of Dictionary(Of String, Object))()})
                    Return
                End If
                rows = rows.Select(Function(r) AdaptExportRow(r)).ToList()
                Export_LastExportRows = rows
                _host?.SendToWeb("export:previewed", New With {.rows = rows})
            Catch ex As Exception
                _host?.SendToWeb("revit:error", New With {.message = "미리보기 실패: " & ex.Message})
                _host?.SendToWeb("export:previewed", New With {.rows = New List(Of Dictionary(Of String, Object))()})
            End Try
        End Sub

        ' ========== Export: 엑셀 저장 ==========
        Private Sub HandleExportSaveExcel(payload As Dictionary(Of String, Object))
            Try
                Dim rows = TryGetRowsFromPayload(payload)
                If rows Is Nothing OrElse rows.Count = 0 Then rows = Export_LastExportRows
                If rows Is Nothing Then rows = New List(Of Dictionary(Of String, Object))()
                Dim dt = BuildExportDataTableFromRows(rows)
                Dim savePath As String = SaveExcelWithDialog(dt)
                If Not String.IsNullOrEmpty(savePath) Then _host?.SendToWeb("export:saved", New With {.path = savePath})
            Catch ex As Exception
                _host?.SendToWeb("revit:error", New With {.message = "엑셀 저장 실패: " & ex.Message})
            End Try
        End Sub

        ' -------- 서비스 호출/어댑터/테이블 --------
        Private Function TryCallExportPointsService(app As UIApplication, files As List(Of String)) As List(Of Dictionary(Of String, Object))
            Dim names = {"KKY_Tool_Revit.Services.ExportPointsService", "Services.ExportPointsService"}
            For Each n In names
                Dim t = FindType(n, "ExportPointsService")
                If t Is Nothing Then Continue For
                Dim m = t.GetMethod("Run", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance)
                If m Is Nothing Then Continue For
                Dim inst As Object = If(m.IsStatic, Nothing, Activator.CreateInstance(t))
                Dim result = m.Invoke(inst, New Object() {app, files})
                Return AnyToRows(result)
            Next
            Return Nothing
        End Function

        Private Function AdaptExportRow(r As Dictionary(Of String, Object)) As Dictionary(Of String, Object)
            If r Is Nothing Then Return New Dictionary(Of String, Object)(StringComparer.Ordinal)
            ' 컬럼 명세를 고정 (home.js/export.js 스키마와 일치)
            Dim d As New Dictionary(Of String, Object)(StringComparer.Ordinal)
            d("File") = If(r.ContainsKey("File"), r("File"), If(r.ContainsKey("file"), r("file"), ""))
            d("ProjectPoint_E(mm)") = SafeToString(r, "ProjectPoint_E(mm)")
            d("ProjectPoint_N(mm)") = SafeToString(r, "ProjectPoint_N(mm)")
            d("ProjectPoint_Z(mm)") = SafeToString(r, "ProjectPoint_Z(mm)")
            d("SurveyPoint_E(mm)") = SafeToString(r, "SurveyPoint_E(mm)")
            d("SurveyPoint_N(mm)") = SafeToString(r, "SurveyPoint_N(mm)")
            d("SurveyPoint_Z(mm)") = SafeToString(r, "SurveyPoint_Z(mm)")
            d("TrueNorthAngle(deg)") = SafeToString(r, "TrueNorthAngle(deg)")
            Return d
        End Function

        Private Function BuildExportDataTableFromRows(rows As List(Of Dictionary(Of String, Object))) As DataTable
            Dim dt As New DataTable("Export")
            Dim headers = {
                "File",
                "ProjectPoint_E(mm)", "ProjectPoint_N(mm)", "ProjectPoint_Z(mm)",
                "SurveyPoint_E(mm)", "SurveyPoint_N(mm)", "SurveyPoint_Z(mm)",
                "TrueNorthAngle(deg)"
            }
            For Each h In headers : dt.Columns.Add(h) : Next
            For Each r In rows
                Dim dr = dt.NewRow()
                dr(0) = SafeToString(r, "File")
                dr(1) = SafeToString(r, "ProjectPoint_E(mm)")
                dr(2) = SafeToString(r, "ProjectPoint_N(mm)")
                dr(3) = SafeToString(r, "ProjectPoint_Z(mm)")
                dr(4) = SafeToString(r, "SurveyPoint_E(mm)")
                dr(5) = SafeToString(r, "SurveyPoint_N(mm)")
                dr(6) = SafeToString(r, "SurveyPoint_Z(mm)")
                dr(7) = SafeToString(r, "TrueNorthAngle(deg)")
                dt.Rows.Add(dr)
            Next
            Return dt
        End Function

        ' ==================================================================
        ' Export local helpers (self-contained; no cross-module dependency)
        ' ==================================================================

        ' 마지막 미리보기 결과(엑셀 저장 시 payload 없을 때 사용)
        Private Shared Export_LastExportRows As List(Of Dictionary(Of String, Object)) _
            = New List(Of Dictionary(Of String, Object))()

        ' payload에서 string 리스트 추출(e.g., files[])
        Private Shared Function ExtractStringList(payload As Dictionary(Of String, Object), key As String) As List(Of String)
            Dim res As New List(Of String)()
            If payload Is Nothing OrElse Not payload.ContainsKey(key) OrElse payload(key) Is Nothing Then Return res
            Dim v = payload(key)
            Dim arr = TryCast(v, System.Collections.IEnumerable)
            If arr Is Nothing Then
                Dim s As String = TryCast(v, String)
                If Not String.IsNullOrEmpty(s) Then res.Add(s)
                Return res
            End If
            For Each o In arr
                If o Is Nothing Then Continue For
                Dim s = o.ToString()
                If Not String.IsNullOrWhiteSpace(s) Then res.Add(s)
            Next
            Return res
        End Function

        ' 다양한 반환값 → 표준 rows
        Private Shared Function AnyToRows(any As Object) As List(Of Dictionary(Of String, Object))
            Dim result As New List(Of Dictionary(Of String, Object))()
            If any Is Nothing Then Return result

            If TypeOf any Is List(Of Dictionary(Of String, Object)) Then
                Return DirectCast(any, List(Of Dictionary(Of String, Object)))
            End If

            Dim dt As DataTable = TryCast(any, DataTable)
            If dt IsNot Nothing Then
                Return DataTableToRows(dt)
            End If

            Dim ie = TryCast(any, System.Collections.IEnumerable)
            If ie IsNot Nothing Then
                For Each item In ie
                    Dim d As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                    Dim dict = TryCast(item, System.Collections.IDictionary)
                    If dict IsNot Nothing Then
                        For Each k In dict.Keys
                            d(k.ToString()) = dict(k)
                        Next
                    Else
                        Dim t = item.GetType()
                        For Each p In t.GetProperties()
                            d(p.Name) = p.GetValue(item, Nothing)
                        Next
                    End If
                    result.Add(d)
                Next
            End If
            Return result
        End Function

        ' 행 딕셔너리에서 컬럼값 안전 추출
        Private Shared Function SafeToString(row As Dictionary(Of String, Object), col As String) As String
            If row Is Nothing Then Return String.Empty
            Dim v As Object = Nothing
            If row.TryGetValue(col, v) AndAlso v IsNot Nothing Then
                Return Convert.ToString(v, Globalization.CultureInfo.InvariantCulture)
            End If
            Return String.Empty
        End Function

        ' DataTable → rows
        Private Shared Function DataTableToRows(dt As DataTable) As List(Of Dictionary(Of String, Object))
            Dim list As New List(Of Dictionary(Of String, Object))()
            If dt Is Nothing Then Return list
            For Each r As DataRow In dt.Rows
                Dim d As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                For Each c As DataColumn In dt.Columns
                    d(c.ColumnName) = If(r.IsNull(c), Nothing, r(c))
                Next
                list.Add(d)
            Next
            Return list
        End Function

        ' 로딩된 어셈블리들에서 타입 찾기(정식명/간단명)
        Private Shared Function FindType(fullOrSimple As String, Optional simpleMatch As String = Nothing) As Type
            ' 직접 시도
            Dim t = Type.GetType(fullOrSimple, False)
            If t IsNot Nothing Then Return t
            ' 로드된 어셈블리 순회
            For Each asm In AppDomain.CurrentDomain.GetAssemblies()
                Try
                    t = asm.GetType(fullOrSimple, False)
                    If t IsNot Nothing Then Return t
                    If Not String.IsNullOrEmpty(simpleMatch) Then
                        For Each ti In asm.GetTypes()
                            If String.Equals(ti.Name, simpleMatch, StringComparison.OrdinalIgnoreCase) Then
                                Return ti
                            End If
                        Next
                    End If
                Catch
                End Try
            Next
            Return Nothing
        End Function

        ' payload에서 rows 추출(있으면) — 없으면 빈 리스트
        Private Shared Function TryGetRowsFromPayload(payload As Dictionary(Of String, Object)) As List(Of Dictionary(Of String, Object))
            If payload Is Nothing Then Return New List(Of Dictionary(Of String, Object))()
            If payload.ContainsKey("rows") AndAlso payload("rows") IsNot Nothing Then
                Return AnyToRows(payload("rows"))
            End If
            If payload.ContainsKey("data") AndAlso payload("data") IsNot Nothing Then
                Return AnyToRows(payload("data"))
            End If
            Dim ie = TryCast(payload, System.Collections.IEnumerable)
            If ie IsNot Nothing Then
                Return AnyToRows(ie)
            End If
            Return New List(Of Dictionary(Of String, Object))()
        End Function

        ' DataTable을 저장 대화상자로 엑셀로 저장하고 경로 반환(취소 시 "")
        Private Shared Function SaveExcelWithDialog(dt As DataTable) As String
            If dt Is Nothing OrElse dt.Columns.Count = 0 Then Return String.Empty
            Dim dlg As New Microsoft.Win32.SaveFileDialog() With {
                .Filter = "Excel (*.xlsx)|*.xlsx",
                .FileName = "export.xlsx"
            }
            Dim ok = dlg.ShowDialog()
            If ok <> True Then Return String.Empty
            Dim path = dlg.FileName
            Try
                Dim wb As IWorkbook = New XSSFWorkbook()
                Dim sh = wb.CreateSheet("Export")
                ' 헤더
                Dim hr = sh.CreateRow(0)
                For c = 0 To dt.Columns.Count - 1
                    hr.CreateCell(c).SetCellValue(dt.Columns(c).ColumnName)
                Next
                ' 데이터
                Dim rIndex = 1
                For Each dr As DataRow In dt.Rows
                    Dim rr = sh.CreateRow(rIndex) : rIndex += 1
                    For c = 0 To dt.Columns.Count - 1
                        Dim v = If(dr.IsNull(c), "", Convert.ToString(dr(c), Globalization.CultureInfo.InvariantCulture))
                        rr.CreateCell(c).SetCellValue(v)
                    Next
                Next
                ' 자동 너비
                For c = 0 To dt.Columns.Count - 1 : sh.AutoSizeColumn(c) : Next
                Using fs As New FileStream(path, FileMode.Create, FileAccess.Write)
                    wb.Write(fs)
                End Using
                Return path
            Catch ex As Exception
                _host?.SendToWeb("host:error", New With {.message = "엑셀 저장 실패: " & ex.Message})
                Return String.Empty
            End Try
        End Function

    End Class
End Namespace
