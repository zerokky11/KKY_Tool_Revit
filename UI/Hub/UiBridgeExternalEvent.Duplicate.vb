Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports Exports

Namespace UI.Hub

    ' Web(JS) ↔ Revit(VB) 브릿지: Duplicate Inspector
    ' - 이벤트/페이로드/응답 이름: 기존과 100% 동일 유지
    ' - 삭제/복구: A안(세션 TransactionGroup + Replay)
    ' - 엑셀 내보내기: FIX2 스키마 고정

    Partial Public Class UiBridgeExternalEvent

#Region "상태 (세션/결과 보관)"
        ' 문서당 1개의 삭제 세션을 유지
        Private _deleteSession As TransactionGroup = Nothing
        Private _sessionDoc As Document = Nothing

        ' 실제 삭제 적용된 ElementId 집합 (정수로 보관)
        Private ReadOnly _deletedSet As New HashSet(Of Integer)()

        ' 마지막 스캔 결과(엑셀 내보내기용) — FIX2 스키마로 만들기 쉽게 DTO로 보관
        Private _lastRows As New List(Of DupRowDto)

        Private Class DupRowDto
            Public Property ElementId As Integer
            Public Property Category As String
            Public Property Family As String
            Public Property [Type] As String
            Public Property ConnectedCount As Integer
            Public Property ConnectedIds As String
            Public Property Candidate As Boolean
            Public Property Deleted As Boolean
        End Class
#End Region

#Region "핸들러 (Core에서 리플렉션 호출)"

        ' ====== 중복 스캔 ======
        Private Sub HandleDupRun(app As UIApplication, payload As Object)
            Dim uiDoc As UIDocument = app.ActiveUIDocument
            If uiDoc Is Nothing OrElse uiDoc.Document Is Nothing Then
                SendToWeb("host:error", New With {.message = "활성 문서가 없습니다."})
                Return
            End If

            Dim doc As Document = uiDoc.Document

            ' 이전 삭제 세션이 남아있다면 보수적으로 확정
            ConfirmPendingDeletesIfAny()

            ' 스캔 옵션 (payload에서 tolFeet 받을 수 있음)
            Dim tolFeet As Double = 1.0 / 64.0 ' ≈ 5mm
            Dim useActiveViewOnly As Boolean = False
            Try
                Dim tolObj = GetProp(payload, "tolFeet")
                If tolObj IsNot Nothing Then tolFeet = Math.Max(0.000001, Convert.ToDouble(tolObj))
            Catch
            End Try

            ' 1) 후보 수집
            Dim rows As New List(Of DupRowDto)()
            Dim total As Integer = 0
            Dim groupsWithDup As Integer = 0
            Dim candidates As Integer = 0

            Dim collector As New FilteredElementCollector(doc)
            collector.WhereElementIsNotElementType()
            If useActiveViewOnly AndAlso uiDoc.ActiveView IsNot Nothing Then
                collector = New FilteredElementCollector(doc, uiDoc.ActiveView.Id).WhereElementIsNotElementType()
            End If

            ' 좌표 양자화(버킷팅) 함수
            Dim q = Function(x As Double) As Long
                        Return CLng(Math.Round(x / tolFeet))
                    End Function

            ' Key: (Category, Family, Type, Level, QuantizedXYZ)
            Dim buckets As New Dictionary(Of String, List(Of ElementId))(StringComparer.Ordinal)
            ' 캐시(문자열 생성 비용 절감)
            Dim catCache As New Dictionary(Of Integer, String)
            Dim famCache As New Dictionary(Of Integer, String)
            Dim typCache As New Dictionary(Of Integer, String)

            For Each e As Element In collector
                total += 1
                If ShouldSkipForQuantity(e) Then Continue For
                If e Is Nothing OrElse e.Category Is Nothing Then Continue For

                Dim center As XYZ = TryGetCenter(e)
                If center Is Nothing Then Continue For

                Dim catName As String = SafeCategoryName(e, catCache)
                Dim famName As String = SafeFamilyName(e, famCache)
                Dim typName As String = SafeTypeName(e, typCache)
                Dim lvl As Integer = TryGetLevelId(e)

                Dim key As String =
                  String.Concat(catName, "|", famName, "|", typName, "|L", lvl.ToString(),
                                "|Q(", q(center.X).ToString(), ",", q(center.Y).ToString(), ",", q(center.Z).ToString(), ")")

                Dim list As List(Of ElementId) = Nothing
                If Not buckets.TryGetValue(key, list) Then
                    list = New List(Of ElementId)()
                    buckets.Add(key, list)
                End If
                list.Add(e.Id)
            Next

            ' 2) 그룹 → 행 변환
            For Each kv In buckets
                Dim ids As List(Of ElementId) = kv.Value
                If ids.Count <= 1 Then Continue For

                groupsWithDup += 1

                For Each id As ElementId In ids
                    Dim e As Element = doc.GetElement(id)
                    If e Is Nothing Then Continue For

                    Dim catName As String = SafeCategoryName(e, catCache)
                    Dim famName As String = SafeFamilyName(e, famCache)
                    Dim typName As String = SafeTypeName(e, typCache)

                    Dim connIds = ids.Where(Function(x) x.IntegerValue <> id.IntegerValue).
                                      Select(Function(x) x.IntegerValue.ToString()).
                                      ToArray()

                    rows.Add(New DupRowDto With {
                      .ElementId = id.IntegerValue,
                      .Category = catName,
                      .Family = famName,
                      .Type = typName,
                      .ConnectedCount = connIds.Length,
                      .ConnectedIds = String.Join(", ", connIds),
                      .Candidate = True,
                      .Deleted = False
                    })
                    candidates += 1
                Next
            Next

            ' 서버 상태에 보관(엑셀/상태 계산용)
            _lastRows = rows

            ' 3) 웹으로 전송 (원본 스키마 그대로)
            Dim wireRows = rows.Select(Function(r) New With {
                .elementId = r.ElementId,
                .category = r.Category,
                .family = r.Family,
                .type = r.Type,
                .connectedCount = r.ConnectedCount,
                .connectedIds = r.ConnectedIds,
                .candidate = r.Candidate,
                .deleted = r.Deleted
            }).ToList()

            SendToWeb("dup:list", wireRows)
            SendToWeb("dup:result", New With {
              .scan = total,
              .groups = groupsWithDup,
              .candidates = candidates
            })
        End Sub

        ' ====== 선택/줌 ======
        Private Sub HandleDuplicateSelect(app As UIApplication, payload As Object)
            Dim uiDoc As UIDocument = app.ActiveUIDocument
            If uiDoc Is Nothing Then Return

            Dim idVal As Integer = SafeInt(GetProp(payload, "id"))
            If idVal <= 0 Then Return

            Dim elId As New ElementId(idVal)
            Dim el As Element = uiDoc.Document.GetElement(elId)
            If el Is Nothing Then
                SendToWeb("host:warn", New With {.message = $"요소 {idVal} 을(를) 찾을 수 없습니다."})
                Return
            End If

            Try
                uiDoc.Selection.SetElementIds(New List(Of ElementId) From {elId})
            Catch
            End Try

            Dim bb As BoundingBoxXYZ = GetBoundingBox(el)
            Try
                If bb IsNot Nothing Then
                    Dim views = uiDoc.GetOpenUIViews()
                    Dim target = views.FirstOrDefault(Function(v) v.ViewId.IntegerValue = uiDoc.ActiveView.Id.IntegerValue)
                    If target IsNot Nothing Then
                        target.ZoomAndCenterRectangle(bb.Min, bb.Max)
                    Else
                        uiDoc.ShowElements(elId)
                    End If
                Else
                    uiDoc.ShowElements(elId)
                End If
            Catch
            End Try
        End Sub

        ' ====== 삭제: duplicate:delete { id } | { ids: [...] } ======
        Private Sub HandleDuplicateDelete(app As UIApplication, payload As Object)
            Dim uiDoc As UIDocument = app.ActiveUIDocument
            If uiDoc Is Nothing OrElse uiDoc.Document Is Nothing Then
                SendToWeb("revit:error", New With {.message = "활성 문서를 찾을 수 없습니다."})
                Return
            End If

            Dim doc As Document = uiDoc.Document
            Dim ids As List(Of Integer) = ExtractIds(payload)
            If ids Is Nothing OrElse ids.Count = 0 Then
                SendToWeb("revit:error", New With {.message = "잘못된 요청입니다(id 누락/형식 오류)."})
                Return
            End If

            EnsureDeleteSession(doc)

            ' 유효한 ElementId만 모아 배치 삭제
            Dim eidList As New List(Of ElementId)
            For Each i In ids
                If i > 0 Then
                    Dim eid As New ElementId(i)
                    If doc.GetElement(eid) IsNot Nothing Then eidList.Add(eid)
                End If
            Next
            If eidList.Count = 0 Then
                SendToWeb("host:warn", New With {.message = "삭제할 유효한 요소가 없습니다."})
                Return
            End If

            Using t As New Transaction(doc, $"KKY Dup Delete ({eidList.Count})")
                t.Start()
                Try
                    doc.Delete(eidList)
                    t.Commit()
                Catch ex As Exception
                    t.RollBack()
                    SendToWeb("revit:error", New With {.message = $"삭제 실패({eidList.Count}개): {ex.Message}"})
                    Return
                End Try
            End Using

            ' 세션 상태/서버 상태 반영 + UI 통지
            For Each eid In eidList
                _deletedSet.Add(eid.IntegerValue)
                Dim row = _lastRows.FirstOrDefault(Function(r) r.ElementId = eid.IntegerValue)
                If row IsNot Nothing Then row.Deleted = True
                SendToWeb("dup:deleted", New With {.id = eid.IntegerValue})
            Next
        End Sub

        ' ====== 복구: duplicate:restore { id } | { ids: [...] } ======
        Private Sub HandleDuplicateRestore(app As UIApplication, payload As Object)
            Dim uiDoc As UIDocument = app.ActiveUIDocument
            If uiDoc Is Nothing OrElse uiDoc.Document Is Nothing Then
                SendToWeb("revit:error", New With {.message = "활성 문서를 찾을 수 없습니다."})
                Return
            End If

            Dim doc As Document = uiDoc.Document
            If _deleteSession Is Nothing OrElse _sessionDoc Is Nothing Then
                SendToWeb("host:warn", New With {.message = "되돌릴 삭제 세션이 없습니다."})
                Return
            End If

            Dim ids As List(Of Integer) = ExtractIds(payload)
            If ids Is Nothing OrElse ids.Count = 0 Then
                SendToWeb("revit:error", New With {.message = "잘못된 요청입니다(id 누락/형식 오류)."})
                Return
            End If

            ' 1) 세션 전체 복귀
            Try
                _deleteSession.RollBack()
            Catch ex As Exception
                SendToWeb("revit:error", New With {.message = $"세션 복귀 실패: {ex.Message}"})
                _deleteSession = Nothing
                _sessionDoc = Nothing
                _deletedSet.Clear()
                ' 서버 상태도 복구
                For Each i In ids
                    Dim r = _lastRows.FirstOrDefault(Function(x) x.ElementId = i)
                    If r IsNot Nothing Then r.Deleted = False
                Next
                Return
            End Try

            ' 2) 되돌릴 대상들을 삭제 집합에서 제거 + 서버 상태 복구
            For Each i In ids
                _deletedSet.Remove(i)
                Dim r = _lastRows.FirstOrDefault(Function(x) x.ElementId = i)
                If r IsNot Nothing Then r.Deleted = False
            Next

            ' 3) 새 세션 시작 후 나머지 재삭제(빠르게 일괄)
            _deleteSession = New TransactionGroup(doc, "KKY Duplicate Delete Session")
            _deleteSession.Start()
            _sessionDoc = doc

            ReplayDeletes(doc)

            ' 4) UI 반영
            For Each i In ids
                SendToWeb("dup:restored", New With {.id = i})
            Next
        End Sub

        ' ====== 엑셀 내보내기 ======
        Private Sub HandleDuplicateExport(app As UIApplication)
            If _lastRows Is Nothing OrElse _lastRows.Count = 0 Then
                SendToWeb("host:warn", New With {.message = "내보낼 데이터가 없습니다."})
                Return
            End If

            Try
                Dim path As String = DuplicateExport.Save(_lastRows)
                If Not String.IsNullOrEmpty(path) Then
                    ' 웹측 토스트/버튼용(기존 이벤트명 유지)
                    SendToWeb("dup:exported", New With {.path = path})
                End If
            Catch ex As Exception
                SendToWeb("revit:error", New With {.message = $"엑셀 저장에 실패했습니다: {ex.Message}"})
            End Try
        End Sub

#End Region

#Region "세션 보장/리플레이"

        ' 현재 문서에 대한 세션 보장
        Private Sub EnsureDeleteSession(doc As Document)
            If _deleteSession IsNot Nothing AndAlso _sessionDoc IsNot Nothing AndAlso _sessionDoc.Equals(doc) Then
                Exit Sub
            End If

            ' 기존 세션이 남아 있다면 보수적으로 확정
            If _deleteSession IsNot Nothing Then
                SafeEndDeleteSession(assimilate:=True)
            End If

            _deleteSession = New TransactionGroup(doc, "KKY Duplicate Delete Session")
            _deleteSession.Start()
            _sessionDoc = doc
            _deletedSet.Clear()
        End Sub

        Private Sub SafeEndDeleteSession(assimilate As Boolean)
            Try
                If _deleteSession Is Nothing Then Return
                If assimilate Then
                    _deleteSession.Assimilate()
                Else
                    _deleteSession.RollBack()
                End If
            Catch
            Finally
                _deleteSession = Nothing
                _sessionDoc = Nothing
                _deletedSet.Clear()
            End Try
        End Sub

        ' 이전 세션 확정(dup:run 직전 호출)
        Private Sub ConfirmPendingDeletesIfAny()
            If _deleteSession IsNot Nothing Then
                SafeEndDeleteSession(assimilate:=True)
                SendToWeb("host:warn", New With {.message = "이전 삭제 세션을 확정했습니다."})
            End If
        End Sub

        ' 남아있는 _deletedSet을 한 번에 재삭제
        Private Sub ReplayDeletes(doc As Document)
            If _deletedSet.Count = 0 Then Return

            Dim ids As New List(Of ElementId)
            For Each i In _deletedSet
                Dim id As New ElementId(i)
                Dim el = doc.GetElement(id)
                If el IsNot Nothing Then ids.Add(id)
            Next
            If ids.Count = 0 Then Return

            Using t As New Transaction(doc, "KKY Dup Replay Deletes")
                t.Start()
                Try
                    doc.Delete(ids)
                    t.Commit()
                Catch ex As Exception
                    t.RollBack()
                    SendToWeb("host:warn", New With {.message = $"일부 요소 재삭제에 실패했습니다. {ex.Message}"})
                End Try
            End Using
        End Sub

#End Region

#Region "필터/유틸 (원본 유지)"

        ' 제외 대상:
        ' - ImportInstance (DWG 등)
        ' - Center line/Centerline/중심선
        ' - Analytical *
        ' - <Area Boundary>, <Sketch> (스케치/영역경계 선)
        ' - 중첩 패밀리(상위 컴포넌트 존재)
        Private Shared Function ShouldSkipForQuantity(e As Element) As Boolean
            If e Is Nothing Then Return True
            If TypeOf e Is ImportInstance Then Return True

            Try
                If e.Category IsNot Nothing Then
                    Dim name As String = If(e.Category.Name, "")
                    Dim n As String = name.ToLowerInvariant()

                    If n.Contains("center line") OrElse n.Contains("centerline") OrElse n.Contains("중심선") Then Return True
                    If n.StartsWith("analytical") Then Return True
                    If n.Contains("<area boundary>") OrElse n.Contains("area boundary") Then Return True
                    If n.Contains("<sketch>") OrElse n = "sketch" Then Return True
                End If
            Catch
            End Try

            Dim fi = TryCast(e, FamilyInstance)
            If fi IsNot Nothing Then
                Try
                    If fi.SuperComponent IsNot Nothing Then Return True
                Catch
                End Try
            End If

            Return False
        End Function

        Private Shared Function SafeCategoryName(e As Element, cache As Dictionary(Of Integer, String)) As String
            If e Is Nothing OrElse e.Category Is Nothing Then Return ""
            Dim id As Integer = e.Category.Id.IntegerValue
            Dim s As String = Nothing
            If cache.TryGetValue(id, s) Then Return s
            s = e.Category.Name
            cache(id) = s
            Return s
        End Function

        Private Shared Function SafeFamilyName(e As Element, cache As Dictionary(Of Integer, String)) As String
            Dim fi = TryCast(e, FamilyInstance)
            If fi Is Nothing OrElse fi.Symbol Is Nothing OrElse fi.Symbol.Family Is Nothing Then Return ""
            Dim id As Integer = fi.Symbol.Family.Id.IntegerValue
            Dim s As String = Nothing
            If cache.TryGetValue(id, s) Then Return s
            s = fi.Symbol.Family.Name
            cache(id) = s
            Return s
        End Function

        Private Shared Function SafeTypeName(e As Element, cache As Dictionary(Of Integer, String)) As String
            Dim fi = TryCast(e, FamilyInstance)
            If fi IsNot Nothing AndAlso fi.Symbol IsNot Nothing Then
                Dim id As Integer = fi.Symbol.Id.IntegerValue
                Dim s As String = Nothing
                If cache.TryGetValue(id, s) Then Return s
                s = fi.Symbol.Name
                cache(id) = s
                Return s
            End If
            Return e.Name
        End Function

        Private Shared Function TryGetLevelId(e As Element) As Integer
            Try
                Dim p As Parameter = e.Parameter(BuiltInParameter.LEVEL_PARAM)
                If p IsNot Nothing Then
                    Dim lvid As ElementId = p.AsElementId()
                    If lvid IsNot Nothing AndAlso lvid <> ElementId.InvalidElementId Then
                        Return lvid.IntegerValue
                    End If
                End If
            Catch
            End Try
            Try
                Dim pi = e.GetType().GetProperty("LevelId")
                If pi IsNot Nothing Then
                    Dim id = TryCast(pi.GetValue(e, Nothing), ElementId)
                    If id IsNot Nothing AndAlso id <> ElementId.InvalidElementId Then
                        Return id.IntegerValue
                    End If
                End If
            Catch
            End Try
            Return -1
        End Function

        Private Shared Function TryGetCenter(e As Element) As XYZ
            If e Is Nothing Then Return Nothing
            Try
                Dim loc As Location = e.Location
                If TypeOf loc Is LocationPoint Then
                    Return CType(loc, LocationPoint).Point
                ElseIf TypeOf loc Is LocationCurve Then
                    Dim crv = CType(loc, LocationCurve).Curve
                    If crv IsNot Nothing Then
                        Return crv.Evaluate(0.5, True)
                    End If
                End If
            Catch
            End Try

            Dim bb = GetBoundingBox(e)
            If bb IsNot Nothing Then
                Return (bb.Min + bb.Max) * 0.5
            End If
            Return Nothing
        End Function

        Private Shared Function GetBoundingBox(e As Element) As BoundingBoxXYZ
            Try
                Dim bb As BoundingBoxXYZ = e.BoundingBox(Nothing)
                If bb IsNot Nothing Then Return bb
            Catch
            End Try
            Return Nothing
        End Function

        Private Shared Function SafeInt(o As Object) As Integer
            If o Is Nothing Then Return 0
            Try
                Return Convert.ToInt32(o)
            Catch
                Return 0
            End Try
        End Function

        ' ====== 단건/다건 id 파싱 유틸 ======
        Private Shared Function ExtractIds(payload As Object) As List(Of Integer)
            Dim result As New List(Of Integer)

            ' 1) 단건: { id: 123 }
            Dim singleObj = GetProp(payload, "id")
            Dim v As Integer = SafeToInt(singleObj)
            If v > 0 Then
                result.Add(v)
                Return result
            End If

            ' 2) 다건: { ids: [...] }
            Dim arr = GetProp(payload, "ids")
            If arr Is Nothing Then Return result

            Dim enumerable = TryCast(arr, System.Collections.IEnumerable)
            If enumerable IsNot Nothing Then
                For Each o In enumerable
                    Dim iv = SafeToInt(o)
                    If iv > 0 Then result.Add(iv)
                Next
            End If

            Return result
        End Function

        Private Shared Function SafeToInt(o As Object) As Integer
            If o Is Nothing Then Return 0
            Try
                If TypeOf o Is Integer Then Return CInt(o)
                If TypeOf o Is Long Then Return CInt(CLng(o))
                If TypeOf o Is Double Then Return CInt(CDbl(o))
                If TypeOf o Is String Then
                    Dim s As String = CStr(o)
                    Dim iv As Integer
                    If Integer.TryParse(s, iv) Then Return iv
                End If
            Catch
            End Try
            Return 0
        End Function

#End Region

    End Class
End Namespace
