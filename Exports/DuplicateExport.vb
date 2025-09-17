Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports Autodesk.Revit.UI
Imports Infrastructure   ' ExcelCore 호출용(완전한 네임스페이스)

Namespace Exports

    Public Module DuplicateExport

        ' 스키마: ElementId | 카테고리 | Family | Type | 패밀리:타입 | 연결수/연결객체 | 상태
        Public Function Save(rows As IEnumerable(Of Object)) As String
            If rows Is Nothing Then Return Nothing

            Dim headers = New String() {"ElementId", "카테고리", "Family", "Type", "패밀리:타입", "연결수/연결객체", "상태"}
            Dim recs As New List(Of Object())()
            Dim keys As New List(Of String)()

            ' 머티리얼라이즈 + 그룹키(같은 연결세트 + Family/Type/Category)
            For Each row In rows
                Dim id As Integer = ToInt(GetVal(row, "ElementId"))
                Dim cat As String = ToStr(GetVal(row, "Category"))
                Dim fam As String = ToStr(GetVal(row, "Family"))
                Dim typ As String = ToStr(GetVal(row, "Type"))
                Dim combined As String = If(String.IsNullOrEmpty(fam) AndAlso String.IsNullOrEmpty(typ), "— : —", fam & " : " & typ)
                Dim cc As Integer = ToInt(GetVal(row, "ConnectedCount"))
                Dim cids As String = ToStr(GetVal(row, "ConnectedIds"))
                Dim del As Boolean = ToBool(GetVal(row, "Deleted"))
                Dim cand As Boolean = ToBool(GetVal(row, "Candidate"))
                Dim status As String = If(del, "삭제됨", If(cand, "삭제후보", ""))

                Dim setIds = New List(Of String)()
                If Not String.IsNullOrWhiteSpace(cids) Then setIds.AddRange(cids.Split(","c).Select(Function(s) s.Trim()))
                setIds.Add(id.ToString())
                setIds = setIds.Distinct().OrderBy(Function(s) s).ToList()

                Dim gKey As String = $"{cat}|{fam}|{typ}|{String.Join(",", setIds)}"
                keys.Add(gKey)

                recs.Add(New Object() {id, cat, fam, typ, combined, $"{cc} / {cids}", status})
            Next

            ' 같은 그룹끼리 인접하도록 정렬
            Dim ordered = recs.
              Select(Function(r, i) New With {.Row = r, .K = keys(i)}).
              OrderBy(Function(x) x.K).
              ThenBy(Function(x) Convert.ToInt32(x.Row(0))).
              Select(Function(x) x.Row).
              ToList()

            ' 공통 코어 호출(경로 선택 포함)
            Dim path As String = Infrastructure.ExcelCore.PickAndSaveXlsx(
              $"Duplicate_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
              "Duplicate",
              headers,
              ordered
            )

            ' 저장 후 열기(사용자 확인)
            If Not String.IsNullOrEmpty(path) Then
                Dim td As New TaskDialog("엑셀 내보내기")
                td.MainInstruction = "추출 완료"
                td.MainContent = "파일을 지금 확인하시겠습니까?"
                td.CommonButtons = TaskDialogCommonButtons.Yes Or TaskDialogCommonButtons.No
                If td.Show() = TaskDialogResult.Yes Then
                    Try : System.Diagnostics.Process.Start(path) : Catch : End Try
                End If
            End If

            Return path
        End Function

        ' ── 값 추출 유틸 ──────────────────────────────────────────────────────
        Private Function GetVal(o As Object, name As String) As Object
            If o Is Nothing Then Return Nothing
            Try
                Dim t = o.GetType()
                Dim p = t.GetProperty(name) : If p IsNot Nothing Then Return p.GetValue(o, Nothing)
                Dim f = t.GetField(name) : If f IsNot Nothing Then Return f.GetValue(o)
            Catch
            End Try
            Return Nothing
        End Function

        Private Function ToStr(o As Object) As String
            If o Is Nothing Then Return "" Else Return o.ToString()
        End Function

        Private Function ToInt(o As Object) As Integer
            If o Is Nothing Then Return 0
            Try : Return Convert.ToInt32(o) : Catch : Return 0 : End Try
        End Function

        Private Function ToBool(o As Object) As Boolean
            If o Is Nothing Then Return False
            Try : Return Convert.ToBoolean(o) : Catch : Return False : End Try
        End Function

    End Module

End Namespace
