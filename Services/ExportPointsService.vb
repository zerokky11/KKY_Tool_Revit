Imports Autodesk.Revit.UI
Imports Autodesk.Revit.DB
Imports System.IO
Imports System.Linq

Namespace Services

    Public Class ExportPointsService

        Public Class Row
            Public Property File As String
            Public Property ProjectE As Double
            Public Property ProjectN As Double
            Public Property ProjectZ As Double
            Public Property SurveyE As Double
            Public Property SurveyN As Double
            Public Property SurveyZ As Double
            Public Property TrueNorth As Double
        End Class

        Public Shared Function Run(uiapp As UIApplication, files As Object) As IList(Of Row)
            Dim app = uiapp.Application
            Dim list As New List(Of Row)()

            Dim paths As New List(Of String)()
            If TypeOf files Is IEnumerable(Of Object) Then
                For Each o In CType(files, IEnumerable(Of Object))
                    Dim s = TryCast(o, String)
                    If Not String.IsNullOrWhiteSpace(s) AndAlso File.Exists(s) Then paths.Add(s)
                Next
            ElseIf TypeOf files Is String AndAlso File.Exists(CStr(files)) Then
                paths.Add(CStr(files))
            End If
            paths = paths.Distinct().ToList()

            If paths.Count = 0 Then Return list

            For Each p In paths
                Try
                    Dim opt As New OpenOptions()
                    opt.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach
                    opt.Audit = False
                    Dim mp As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(p)
                    Using doc As Document = app.OpenDocumentFile(mp, opt)
                        Dim row As New Row()
                        row.File = Path.GetFileName(p)
                        Extract(doc, row)
                        list.Add(row)
                        doc.Close(False)
                    End Using
                Catch
                End Try
            Next
            Return list
        End Function

        Public Shared Function ExportToExcel(uiapp As UIApplication, files As Object) As String
            Dim rows = Run(uiapp, files)
            Dim desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            Dim filePath = Path.Combine(desktop, $"ExportPoints_{Date.Now:yyyyMMdd_HHmmss}.xlsx")
            Dim headers = New String() {"File", "ProjectPoint_E(mm)", "ProjectPoint_N(mm)", "ProjectPoint_Z(mm)", "SurveyPoint_E(mm)", "SurveyPoint_N(mm)", "SurveyPoint_Z(mm)", "TrueNorthAngle(deg)"}
            Dim data = rows.Select(Function(r) New Object() {
                r.File, Math.Round(r.ProjectE, 3), Math.Round(r.ProjectN, 3), Math.Round(r.ProjectZ, 3),
                Math.Round(r.SurveyE, 3), Math.Round(r.SurveyN, 3), Math.Round(r.SurveyZ, 3), Math.Round(r.TrueNorth, 3)
            })
            ExcelCore.SaveTable(filePath, "Points", headers, data)
            Return filePath
        End Function

        Private Shared Sub Extract(doc As Document, row As Row)
            Dim basePt As BasePoint = New FilteredElementCollector(doc).OfClass(GetType(BasePoint)).Cast(Of BasePoint)().FirstOrDefault(Function(bp) bp.IsShared = False)
            Dim surveyPt As BasePoint = New FilteredElementCollector(doc).OfClass(GetType(BasePoint)).Cast(Of BasePoint)().FirstOrDefault(Function(bp) bp.IsShared = True)

            Dim project As XYZ = If(basePt IsNot Nothing, basePt.Position, XYZ.Zero)
            Dim survey As XYZ = If(surveyPt IsNot Nothing, surveyPt.Position, XYZ.Zero)

            row.ProjectE = project.X * 304.8
            row.ProjectN = project.Y * 304.8
            row.ProjectZ = project.Z * 304.8

            row.SurveyE = survey.X * 304.8
            row.SurveyN = survey.Y * 304.8
            row.SurveyZ = survey.Z * 304.8

            ' fix2와 동일: LookupParameter("Angle to True North")
            Dim deg As Double = 0.0
            Try
                If basePt IsNot Nothing Then
                    Dim p = basePt.LookupParameter("Angle to True North")
                    If p IsNot Nothing Then
                        deg = p.AsDouble() * (180.0 / Math.PI)
                    End If
                End If
            Catch
            End Try
            If deg = 0.0 Then
                Try
                    Dim pl As ProjectLocation = doc.ActiveProjectLocation
                    Dim pp As ProjectPosition = pl.GetProjectPosition(XYZ.Zero)
                    If pp IsNot Nothing Then
                        deg = pp.Angle * (180.0 / Math.PI)
                    End If
                Catch
                End Try
            End If
            row.TrueNorth = deg
        End Sub

    End Class

End Namespace
