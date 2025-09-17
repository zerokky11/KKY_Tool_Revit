Imports System
Imports System.IO
Imports System.Windows
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.Wpf
Imports System.Web.Script.Serialization
Imports System.Reflection

Namespace UI.Hub
    Public Class HubHostWindow
        Inherits Window

        Private ReadOnly _web As New WebView2()
        Private ReadOnly _serializer As New JavaScriptSerializer()

        Public ReadOnly Property Web As WebView2
            Get
                Return _web
            End Get
        End Property

        Public Shared Property Current As HubHostWindow

        Private _initStarted As Boolean = False

        Public Sub New()
            Title = "KKY Tool"
            Width = 1280
            Height = 800
            WindowStartupLocation = WindowStartupLocation.CenterScreen
            Content = _web

            Current = Me
            AddHandler Loaded, AddressOf OnLoaded
        End Sub

        Private Function ResolveUiFolder() As String
            Try
                Dim asm = Assembly.GetExecutingAssembly()
                Dim baseDir = Path.GetDirectoryName(asm.Location)
                Dim ui = Path.Combine(baseDir, "Resources", "HubUI")
                If Directory.Exists(ui) Then Return Path.GetFullPath(ui)
            Catch
            End Try
            Return Nothing
        End Function

        Private Async Sub OnLoaded(sender As Object, e As RoutedEventArgs)
            If _initStarted Then Return
            _initStarted = True
            Try
                Dim userData = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "KKY_Tool_Revit", "WebView2UserData")
                Directory.CreateDirectory(userData)

                Dim env = Await CoreWebView2Environment.CreateAsync(Nothing, userData, Nothing)
                Await _web.EnsureCoreWebView2Async(env)
                Dim core = _web.CoreWebView2

                core.Settings.AreDefaultContextMenusEnabled = False
                core.Settings.IsStatusBarEnabled = False
#If DEBUG Then
                core.Settings.AreDevToolsEnabled = True
#Else
                core.Settings.AreDevToolsEnabled = False
#End If

                ' 가상 호스트 매핑
                Dim uiFolder = ResolveUiFolder()
                If String.IsNullOrEmpty(uiFolder) Then
                    Throw New DirectoryNotFoundException("Resources\HubUI 폴더를 찾을 수 없습니다.")
                End If
                core.SetVirtualHostNameToFolderMapping(
                    "hub.local", uiFolder, CoreWebView2HostResourceAccessKind.Allow)

                ' 메시지 브리지
                AddHandler core.WebMessageReceived, AddressOf OnWebMessage

                ' ===== ExternalEvent 초기화(호스트 참조 주입) =====
                UiBridgeExternalEvent.Initialize(Me)

                ' 허브 진입
                _web.Source = New Uri("https://hub.local/index.html")

                ' 초기 상태 알림
                SendToWeb("host:topmost", New With {.on = Me.Topmost})

            Catch ex As Exception
                Dim hr As Integer = Runtime.InteropServices.Marshal.GetHRForException(ex)
                MessageBox.Show($"WebView 초기화 실패 (0x{hr:X8}) : {ex.Message}",
                                "KKY Tool", MessageBoxButton.OK, MessageBoxImage.Error)
            Finally
                _initStarted = False
            End Try
        End Sub

        Private Sub OnWebMessage(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
            Try
                Dim root As Dictionary(Of String, Object) =
                    _serializer.Deserialize(Of Dictionary(Of String, Object))(e.WebMessageAsJson)

                Dim name As String = Nothing
                If root IsNot Nothing Then
                    If root.ContainsKey("ev") AndAlso root("ev") IsNot Nothing Then
                        name = Convert.ToString(root("ev"))
                    ElseIf root.ContainsKey("name") AndAlso root("name") IsNot Nothing Then
                        name = Convert.ToString(root("name"))
                    End If
                End If
                If String.IsNullOrEmpty(name) Then Return

                Dim payload As Object = Nothing
                If root IsNot Nothing AndAlso root.ContainsKey("payload") Then
                    payload = root("payload")
                End If

                Select Case name
                    Case "ui:ping"
                        SendToWeb("host:pong", New With {.t = Date.Now.Ticks})

                    Case "ui:toggle-topmost"
                        Me.Topmost = Not Me.Topmost
                        SendToWeb("host:topmost", New With {.on = Me.Topmost})

                    Case "ui:query-topmost"
                        SendToWeb("host:topmost", New With {.on = Me.Topmost})

                    Case Else
                        ' 나머지는 ExternalEvent로 위임
                        UiBridgeExternalEvent.Raise(name, payload)
                End Select

            Catch ex As Exception
                SendToWeb("host:error", New With {.message = ex.Message})
            End Try
        End Sub

        ' .NET → JS (양쪽 호환: ev & name 둘 다 포함해서 송신)
        Public Sub SendToWeb(ev As String, payload As Object)
            Dim core = _web.CoreWebView2
            If core Is Nothing Then Return
            Dim msg As New Dictionary(Of String, Object) From {
                {"ev", ev}, {"name", ev}, {"payload", payload}
            }
            Dim json = _serializer.Serialize(msg)
            core.PostWebMessageAsJson(json)
        End Sub
    End Class
End Namespace
