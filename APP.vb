Imports System.Reflection
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.DB
Imports System.Linq

Namespace KKY_Tool_Revit
    Public Class App
        Implements IExternalApplication

        Public Function OnStartup(a As UIControlledApplication) As Result Implements IExternalApplication.OnStartup
            Const TAB_NAME As String = "KKY Tools"
            Const PANEL_NAME As String = "Hub"

            ' 탭 생성(이미 있으면 무시)
            Try : a.CreateRibbonTab(TAB_NAME) : Catch : End Try

            ' 패널 찾기/생성
            Dim ribbonPanel As RibbonPanel = a.GetRibbonPanels(TAB_NAME).FirstOrDefault(Function(p) p.Name = PANEL_NAME)
            If ribbonPanel Is Nothing Then
                ribbonPanel = a.CreateRibbonPanel(TAB_NAME, PANEL_NAME)
            End If

            ' 허브 버튼
            Dim asmPath = Assembly.GetExecutingAssembly().Location
            Dim cmdFullName As String = GetType(DuplicateExport).FullName
            Dim pbd As New PushButtonData("KKY_Hub_Button", "KKY Hub", asmPath, cmdFullName)

            Dim btn = TryCast(ribbonPanel.AddItem(pbd), PushButton)
            If btn IsNot Nothing Then
                btn.ToolTip = "KKY Tool 허브 열기"
                btn.Image = LoadPng("KKY_Tool_Revit.Resources.icons.hub_64.png")
                btn.LargeImage = LoadPng("KKY_Tool_Revit.Resources.icons.hub_64.png")
            End If

            Return Result.Succeeded
        End Function

        Public Function OnShutdown(a As UIControlledApplication) As Result Implements IExternalApplication.OnShutdown
            Return Result.Succeeded
        End Function

        Private Function LoadPng(resName As String) As ImageSource
            Try
                Dim asm = Assembly.GetExecutingAssembly()
                Using s = asm.GetManifestResourceStream(resName)
                    If s Is Nothing Then Return Nothing
                    Dim img As New BitmapImage()
                    img.BeginInit()
                    img.StreamSource = s
                    img.CacheOption = BitmapCacheOption.OnLoad
                    img.EndInit()
                    img.Freeze()
                    Return img
                End Using
            Catch
                Return Nothing
            End Try
        End Function
    End Class
End Namespace
