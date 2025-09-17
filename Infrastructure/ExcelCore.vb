Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports Microsoft.Win32
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Namespace Infrastructure

    Public Module ExcelCore

        ' ── 모듈 초기화: 코드페이지 등록 + 애드인 폴더 DLL 우선 로드 ──
        Sub New()
            Try : Encoding.RegisterProvider(CodePagesEncodingProvider.Instance) : Catch : End Try
            Dim baseDir As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            AddHandler AppDomain.CurrentDomain.AssemblyResolve,
              Function(_, args)
                  Try
                      Dim dllName As String = (New AssemblyName(args.Name)).Name & ".dll"
                      Dim candidate As String = Path.Combine(baseDir, dllName)
                      If File.Exists(candidate) Then Return Assembly.LoadFrom(candidate)
                  Catch
                  End Try
                  Return Nothing
              End Function
        End Sub

        ' ── 경로 선택 포함 저장기 ──────────────────────────────────────────────
        Public Function PickAndSaveXlsx(defaultFileName As String,
                                        sheet As String,
                                        headers As IEnumerable(Of String),
                                        rows As IEnumerable(Of IEnumerable(Of Object))) As String
            Dim dlg As New SaveFileDialog() With {
              .Filter = "Excel Workbook (*.xlsx)|*.xlsx",
              .FileName = defaultFileName,
              .AddExtension = True,
              .OverwritePrompt = True,
              .ValidateNames = True
            }
            Dim ok = dlg.ShowDialog()
            If Not ok.HasValue OrElse Not ok.Value Then Return Nothing

            SaveXlsx(dlg.FileName, sheet, headers, rows)
            Return dlg.FileName
        End Function

        ' ── 실제 XLSX 생성기(헤더 Bold, 전체 테두리, AutoSize + 폰트 폴백) ───
        Public Sub SaveXlsx(path As String,
                            sheet As String,
                            headers As IEnumerable(Of String),
                            rows As IEnumerable(Of IEnumerable(Of Object)))

            Using wb As IWorkbook = New XSSFWorkbook()
                Dim sh = wb.CreateSheet(If(String.IsNullOrWhiteSpace(sheet), "Sheet1", sheet))

                Dim bold = wb.CreateFont() : bold.IsBold = True
                Dim h = wb.CreateCellStyle() : h.SetFont(bold) : Border(h)
                Dim c = wb.CreateCellStyle() : Border(c)

                ' 헤더
                Dim r As Integer = 0
                Dim hr = sh.CreateRow(r) : r += 1
                Dim col As Integer = 0
                For Each head In headers
                    Dim cell = hr.CreateCell(col) : col += 1
                    cell.SetCellValue(If(head, "")) : cell.CellStyle = h
                Next

                ' 데이터
                For Each row In rows
                    Dim dr = sh.CreateRow(r) : r += 1
                    col = 0
                    If row IsNot Nothing Then
                        For Each v In row
                            Dim cell = dr.CreateCell(col) : col += 1
                            SetCell(cell, v) : cell.CellStyle = c
                        Next
                    End If
                Next

                ' AutoSize (폰트 미설치 환경 폴백)
                Try
                    For i = 0 To CInt(hr.LastCellNum) - 1 : sh.AutoSizeColumn(i) : Next
                Catch
                    For i = 0 To CInt(hr.LastCellNum) - 1 : sh.SetColumnWidth(i, 22 * 256) : Next
                End Try

                Dim dir As String = path.GetDirectoryName(path)
                If Not String.IsNullOrEmpty(dir) AndAlso Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
                Using fs As New FileStream(path, FileMode.Create, FileAccess.Write) : wb.Write(fs) : End Using
            End Using
        End Sub

        ' ── 내부 유틸 ─────────────────────────────────────────────────────────
        Private Sub Border(st As ICellStyle)
            st.BorderBottom = BorderStyle.Thin : st.BorderTop = BorderStyle.Thin
            st.BorderLeft = BorderStyle.Thin : st.BorderRight = BorderStyle.Thin
        End Sub

        Private Sub SetCell(cell As ICell, v As Object)
            If v Is Nothing Then cell.SetCellValue("") : Return
            If TypeOf v Is Integer OrElse TypeOf v Is Long OrElse TypeOf v Is Double Then
                cell.SetCellValue(Convert.ToDouble(v))
            ElseIf TypeOf v Is DateTime Then
                cell.SetCellValue(CType(v, DateTime))
            Else
                cell.SetCellValue(v.ToString())
            End If
        End Sub

    End Module

End Namespace
