﻿Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Reflection
Imports Autodesk.Revit.UI

' 네임스페이스는 프로젝트와 일치시켜 주세요.
Namespace UI.Hub

    ' Web(JS) ↔ Revit(VB) 브릿지의 중심. 외부이벤트 큐, 메시지 라우팅, 공통 로그/브로드캐스트 담당
    Partial Public Class UiBridgeExternalEvent

        ' -----------------------------
        ' 상태/싱글톤
        ' -----------------------------
        Friend Shared _host As HubHostWindow ' 다른 partial(Export 등)에서 사용
        Private Shared ReadOnly _self As New UiBridgeExternalEvent() ' 인스턴스 메서드 호출용
        Private Shared ReadOnly _gate As New Object()
        Private Shared ReadOnly _queue As New Queue(Of Action(Of UIApplication))()

        Private Shared _extEv As ExternalEvent
        Private Shared _handler As IExternalEventHandler

        ' -----------------------------
        ' 공개 진입점
        ' -----------------------------

        ''' <summary>
        ''' CmdOpenHub 에서 최초 1회 호출. Host 저장 + ExternalEvent 준비 + 초기 브로드캐스트.
        ''' </summary>
        Public Shared Sub Initialize(host As HubHostWindow)
            _host = host

            If _extEv Is Nothing Then
                _handler = New BridgeHandler(AddressOf ProcessQueue)
                _extEv = ExternalEvent.Create(_handler)
            End If

            ' 초기 상태 브로드캐스트(항상 위, 연결)
            BroadcastTopmost()
            SendToWeb("host:connected", New With {.ok = True})
        End Sub

        ''' <summary>
        ''' Host(또는 Web 메시지 핸들러)에서 호출: 이름/페이로드를 외부이벤트 큐에 넣고 Revit UI 스레드에서 처리.
        ''' </summary>
        Public Shared Sub Raise(name As String, payload As Object)
            Enqueue(Sub(app) Dispatch(app, name, payload))
        End Sub

        ' -----------------------------
        ' 큐/외부이벤트
        ' -----------------------------
        Private Shared Sub Enqueue(work As Action(Of UIApplication))
            SyncLock _gate
                _queue.Enqueue(work)
            End SyncLock
            If _extEv IsNot Nothing Then
                _extEv.Raise()
            End If
        End Sub

        Private Shared Sub ProcessQueue(app As UIApplication)
            Dim todo As Action(Of UIApplication) = Nothing
            Do
                SyncLock _gate
                    If _queue.Count > 0 Then
                        todo = _queue.Dequeue()
                    Else
                        todo = Nothing
                    End If
                End SyncLock
                If todo Is Nothing Then Exit Do
                Try
                    todo(app)
                Catch ex As Exception
                    SendToWeb("host:error", New With {.message = ex.Message})
                End Try
            Loop
        End Sub

        ' -----------------------------
        ' 라우팅
        ' -----------------------------
        Private Shared Sub Dispatch(app As UIApplication, name As String, payload As Object)
            ' 공통(웹 UI 쪽 요청)
            Select Case name
                Case "ui:query-topmost"
                    BroadcastTopmost()
                    Return

                Case "ui:set-topmost"
                    Dim turnOn As Boolean = False
                    Try
                        Dim raw = GetProp(payload, "on")
                        If raw IsNot Nothing Then turnOn = Convert.ToBoolean(raw)
                    Catch
                    End Try
                    Try
                        If _host IsNot Nothing Then _host.Topmost = turnOn
                    Catch
                    End Try
                    BroadcastTopmost()
                    Return

                Case "ui:toggle-topmost"
                    Try
                        If _host IsNot Nothing Then _host.Topmost = Not _host.Topmost
                    Catch
                    End Try
                    BroadcastTopmost()
                    Return
            End Select

            ' 기능 디스패치: 이벤트명 → 내부 핸들러명
            Dim map As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            ' Duplicate Inspector
            map.Add("dup:run", "HandleDupRun")
            map.Add("duplicate:export", "HandleDuplicateExport")
            map.Add("duplicate:delete", "HandleDuplicateDelete")
            map.Add("duplicate:restore", "HandleDuplicateRestore")
            map.Add("duplicate:select", "HandleDuplicateSelect")
            ' Connector Diagnostics
            map.Add("connector:run", "HandleConnectorRun")
            map.Add("connector:open-excel", "HandleConnectorOpenExcel")
            map.Add("connector:save-excel", "HandleConnectorSaveExcel")
            ' Export Points with Angle
            map.Add("export:browse-folder", "HandleExportBrowse")
            map.Add("export:preview", "HandleExportPreview")
            map.Add("export:save-excel", "HandleExportSaveExcel")

            Dim methodName As String = Nothing
            If Not map.TryGetValue(name, methodName) Then
                SendToWeb("host:warn", New With {.message = String.Format("알 수 없는 이벤트 '{0}'", name)})
                Return
            End If

            ' 동일 Partial 클래스 안의 Private 메서드를 리플렉션으로 찾아 호출
            Dim t As Type = GetType(UiBridgeExternalEvent)
            Dim flags As BindingFlags = BindingFlags.Instance Or BindingFlags.NonPublic Or BindingFlags.Public
            Dim m As MethodInfo = t.GetMethod(methodName, flags)

            If m Is Nothing Then
                ' 메서드가 아직 구현되지 않은 경우에도 앱이 죽지 않도록 안전하게 로그만 남김
                SendToWeb("host:warn", New With {.message = String.Format("핸들러 '{0}' 가 구현되어 있지 않습니다.", methodName)})
                Return
            End If

            ' 시그니처: (UIApplication, payload) or (UIApplication) or (payload) or ()
            Dim ps() As ParameterInfo = m.GetParameters()
            Dim args() As Object
            Select Case ps.Length
                Case 2
                    args = New Object() {app, payload}
                Case 1
                    If ps(0).ParameterType Is GetType(UIApplication) Then
                        args = New Object() {app}
                    Else
                        args = New Object() {payload}
                    End If
                Case Else
                    args = New Object() {}
            End Select

            Try
                m.Invoke(_self, args)
            Catch ex As TargetInvocationException
                Dim msg As String
                If ex.InnerException IsNot Nothing Then
                    msg = ex.InnerException.Message
                Else
                    msg = ex.Message
                End If
                SendToWeb("host:error", New With {.message = String.Format("핸들러 실행 오류({0}): {1}", methodName, msg)})
            Catch ex As Exception
                SendToWeb("host:error", New With {.message = String.Format("핸들러 실행 오류({0}): {1}", methodName, ex.Message)})
            End Try
        End Sub

        ' -----------------------------
        ' 공통 유틸/브로드캐스트
        ' -----------------------------

        Friend Shared Sub SendToWeb(channel As String, payload As Object)
            Try
                If _host IsNot Nothing Then
                    _host.SendToWeb(channel, payload)
                End If
            Catch
                ' Host 없는 초기 구간 등에서는 조용히 무시
            End Try
        End Sub

        Private Shared Sub BroadcastTopmost()
            Try
                Dim onTop As Boolean = False
                If _host IsNot Nothing Then onTop = _host.Topmost
                SendToWeb("host:topmost", New With {.on = onTop})
            Catch
            End Try
        End Sub

        ' payload 속성 안전 추출(익명/Dictionary 수용)
        Private Shared Function GetProp(obj As Object, prop As String) As Object
            If obj Is Nothing Then Return Nothing

            Dim d = TryCast(obj, IDictionary(Of String, Object))
            If d IsNot Nothing Then
                Dim v As Object = Nothing
                If d.TryGetValue(prop, v) Then Return v
                Return Nothing
            End If

            Dim t = obj.GetType()
            Dim p = t.GetProperty(prop, BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.IgnoreCase)
            If p Is Nothing Then Return Nothing
            Return p.GetValue(obj, Nothing)
        End Function

        ' (필요시) 외부에서 직접 로그를 남길 때 사용
        Public Shared Sub HostLog(kind As String, text As String)
            SendToWeb("host:log", New With {.kind = kind, .text = text})
        End Sub

    End Class

    ' 외부이벤트 핸들러(큐를 비우는 역할만 수행)
    Friend Class BridgeHandler
        Implements IExternalEventHandler

        Private ReadOnly _run As Action(Of UIApplication)
        Public Sub New(run As Action(Of UIApplication))
            _run = run
        End Sub

        Public Sub Execute(uiApp As UIApplication) Implements IExternalEventHandler.Execute
            If _run IsNot Nothing Then
                _run.Invoke(uiApp)
            End If
        End Sub

        ' Revit API는 Function GetName() As String 을 요구합니다.
        Public Function GetName() As String Implements IExternalEventHandler.GetName
            Return "KKY Hub Bridge"
        End Function
    End Class

End Namespace
