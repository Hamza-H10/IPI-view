Imports System.Windows.Media
Imports LiveCharts
Imports LiveCharts.Defaults
Imports LiveCharts.Wpf


Public Class Form1
    Dim listBH As List(Of BoreHole), bhIndex As Short = -1
    Dim boreHoleSelected As Short = 0, _axisValue As Short = 0
    Dim bsTextPrintData As String, printFont As Font

    Private Sub ReloadList()

        lstBoreholes.Items.Clear()
        If boreHoleSelected = 0 Then
            listBH = GetBoreholes()

            For Each bitem In listBH
                lstBoreholes.Items.Add("[" & bitem.Id.ToString("D2") & "] " & bitem.SiteName & " - " & bitem.Location)
            Next
            lstBoreholes.SelectionMode = SelectionMode.One
            ToolBarEnable(False)
        Else
            ' get directory listing
            Dim di As New IO.DirectoryInfo(GetBoreholeDirectory(boreHoleSelected))
            Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
            Dim fi As IO.FileInfo

            For Each fi In aryFi
                lstBoreholes.Items.Add(fi.Name)
            Next

            lstBoreholes.SelectionMode = SelectionMode.MultiSimple

            ToolBarEnable(True)
        End If
        ResetLabels()
        CartesianChart1.Visible = False
        DataGridView1.Visible = False
        ToolStrip2.Enabled = False

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.ForeColor = System.Drawing.Color.FromArgb(33, 149, 242)
        Label2.ForeColor = System.Drawing.Color.FromArgb(243, 67, 54)
        Label3.ForeColor = System.Drawing.Color.FromArgb(254, 192, 7)
        Label4.ForeColor = System.Drawing.Color.FromArgb(96, 125, 138)
        Label5.ForeColor = System.Drawing.Color.FromArgb(0, 187, 211)
        OpenDatabase()
        '_DeleteAllBoreholes() ' temporary delete all
        tbGraphType.SelectedIndex = 0

        ReloadList()
    End Sub

    Private Sub FrmMain_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        CloseDatabase()
    End Sub

    Private Sub tbImport_Click(sender As Object, e As EventArgs) Handles tbImport.Click
        Dim cnt As Short = 0, cntError As Short = 0, cntRepeat As Short = 0
        Dim msgString As String = "Import Summary:" & vbCrLf
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then

            For Each strFileName As String In OpenFileDialog1.FileNames
                Dim strFileNew As String = strFileName.Split("\").Last
                If (strFileNew.Split(".").Last.ToLower() = "csv") Then
                    Dim strData As String()() = ReadCSVFile(strFileName)
                    If (strData.Length < 5) Then
                        cntError += 1
                    Else
                        ' Catch 4 poarameters for the new borehole
                        Dim borehole_num As Short, depth As Single, strDirName As String

                        borehole_num = Short.Parse(strData(0)(1))
                        strDirName = GetBoreholeDirectory(borehole_num)
                        strFileNew = strDirName & "\" & strFileNew
                        If System.IO.File.Exists(strFileNew) Then
                            cntRepeat += 1
                        Else
                            If Not System.IO.Directory.Exists(strDirName) Then
                                System.IO.Directory.CreateDirectory(strDirName)
                            End If
                            FileCopy(strFileName, strFileNew)

                            depth = Single.Parse(strData(4)(0))
                            Dim bh As New BoreHole With {.Id = borehole_num, .SiteName = strData(1)(1), .Location = strData(2)(1), .Depth = depth, .BaseFile = ""}
                            If (Not AddBorehole(bh)) Then
                                UpdateBorehole(bh)
                            End If
                            ReloadList()
                            cnt += 1
                        End If

                    End If
                End If
            Next
            If cnt > 0 Then msgString &= "You have added " & cnt & " CSV file(s) to the InclinoView successfully." & vbCrLf
            If cntError > 0 Then msgString &= cntError & " file(s) were found to be incorrect format." & vbCrLf
            If cntRepeat > 0 Then msgString &= cntRepeat & " file(s) were already imported into the application, hence ignored." & vbCrLf

            MsgBox(msgString, MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Import")
        End If

    End Sub


    Private Sub tbBack_Click(sender As Object, e As EventArgs) Handles tbBack.Click
        If boreHoleSelected > 0 Then
            boreHoleSelected = 0
            bhIndex = -1
            ReloadList()
        End If

    End Sub

    Private Sub lstBoreholes_DoubleClick(sender As Object, e As EventArgs) Handles lstBoreholes.DoubleClick
        If lstBoreholes.SelectedIndex < 0 Then Exit Sub
        If boreHoleSelected = 0 Then
            bhIndex = lstBoreholes.SelectedIndex
            boreHoleSelected = listBH(bhIndex).Id
            ReloadList()
        Else
            DataGridView1.Visible = False
            CartesianChart1.Visible = False
            DisplayReport()
            ' lstboreholes.selecteditem is a CSV file
        End If
    End Sub

    Private Sub DisplayReport(Optional bnLoadText As Boolean = False)
        Dim ds As New DataTable, strBaseData As String()()
        Dim bnBaseFilePresent As Boolean
        Dim i As Short, ValA As Single, ValB As Single, absValA As Single = 0, absValB As Single = 0
        Dim bsValA As Single, bsValB As Single, bsAbsValA As Single = 0, bsAbsValB As Single = 0
        Dim deviationA As Single, deviationB As Single

        ResetLabels()

        If listBH(bhIndex).BaseFile Is Nothing Or listBH(bhIndex).BaseFile = "" Then
            bnBaseFilePresent = False
            Label6.Text = ""
        Else
            Dim strFile As String = GetBoreholeDirectory(boreHoleSelected) & "\" & listBH(bhIndex).BaseFile
            If IO.File.Exists(strFile) Then
                strBaseData = ReadCSVFile(strFile)
                bnBaseFilePresent = True
                Label6.Text = "Base File : " & listBH(bhIndex).BaseFile.Split(".").First().Replace("_", ":")
            Else
                MsgBox("Base file does not exist. It must have been deleted. Please select another file as base.", vbOKOnly Or vbExclamation, "Graph")
            End If
        End If
        Label1.Text = lstBoreholes.SelectedItem.ToString().Split(".").First().Replace("_", ":")

        Dim strData As String()() = ReadCSVFile(GetBoreholeDirectory(boreHoleSelected) & "\" & lstBoreholes.SelectedItem)

        ds.Columns.Add("Depth", Type.GetType("System.Single"))
        ds.Columns.Add("Face A+", Type.GetType("System.Single"))
        ds.Columns.Add("Face A-", Type.GetType("System.Single"))
        ds.Columns.Add("Face B+", Type.GetType("System.Single"))
        ds.Columns.Add("Face B-", Type.GetType("System.Single"))

        ds.Columns.Add("Mean A", Type.GetType("System.Single"))
        ds.Columns.Add("Mean B", Type.GetType("System.Single"))

        ds.Columns.Add("Absolute A", Type.GetType("System.Single"))
        ds.Columns.Add("Absolute B", Type.GetType("System.Single"))

        ds.Columns.Add("Deviation A", Type.GetType("System.Single"))
        ds.Columns.Add("Deviation B", Type.GetType("System.Single"))

        For i = 4 To strData.Length - 1
            ValA = (Single.Parse(strData(i)(1)) - Single.Parse(strData(i)(2))) / 2
            ValB = (Single.Parse(strData(i)(3)) - Single.Parse(strData(i)(4))) / 2
            ValA = Math.Round(ValA, 2)
            ValB = Math.Round(ValB, 2)
            absValA += ValA
            absValB += ValB
            absValA = Math.Round(absValA, 2)
            absValB = Math.Round(absValB, 2)
            If bnBaseFilePresent Then
                bsValA = (Single.Parse(strBaseData(i)(1)) - Single.Parse(strBaseData(i)(2))) / 2
                bsValB = (Single.Parse(strBaseData(i)(3)) - Single.Parse(strBaseData(i)(4))) / 2
                bsValA = Math.Round(bsValA, 2)
                bsValB = Math.Round(bsValB, 2)
                bsAbsValA += bsValA
                bsAbsValB += bsValB
                bsAbsValA = Math.Round(bsAbsValA, 2)
                bsAbsValB = Math.Round(bsAbsValB, 2)
                deviationA = Math.Round(absValA - bsAbsValA, 2)
                deviationB = Math.Round(absValB - bsAbsValB, 2)
                ds.Rows.Add(New Object() {Single.Parse(strData(i)(0)), Single.Parse(strData(i)(1)), Single.Parse(strData(i)(2)),
                        Single.Parse(strData(i)(3)), Single.Parse(strData(i)(4)), ValA, ValB, absValA, absValB, deviationA, deviationB})
            Else
                ds.Rows.Add(New Object() {Single.Parse(strData(i)(0)), Single.Parse(strData(i)(1)), Single.Parse(strData(i)(2)),
                        Single.Parse(strData(i)(3)), Single.Parse(strData(i)(4)), ValA, ValB, absValA, absValB})

            End If

        Next

        If bnLoadText Then
            Dim row As Integer, strItem As String

            For i = 0 To ds.Columns.Count - 1
                If i > 6 Then
                    bsTextPrintData &= ds.Columns.Item(i).ColumnName.PadLeft(12)
                Else
                    bsTextPrintData &= ds.Columns.Item(i).ColumnName.PadLeft(8)
                End If
            Next
            bsTextPrintData &= vbCrLf
            bsTextPrintData &= "".PadRight(104, "=") & vbCrLf

            For row = 0 To ds.Rows.Count - 1
                For i = 0 To ds.Columns.Count - 1
                    strItem = ""
                    If Not IsDBNull(ds.Rows(row).Item(i)) Then
                        strItem = FormatNumber(ds.Rows(row).Item(i), 2)
                    End If
                    If i > 6 Then
                        bsTextPrintData &= "  " & strItem.PadLeft(8) & "  "
                    Else
                        bsTextPrintData &= " " & strItem.PadLeft(6) & " "
                    End If
                Next
                bsTextPrintData &= vbCrLf
            Next
        Else
            DataGridView1.DataSource = ds
            DataGridView1.Visible = True

            For i = 0 To DataGridView1.Columns.Count - 1
                If i > 6 Then
                    DataGridView1.Columns(i).Width = 80
                Else
                    DataGridView1.Columns(i).Width = 60
                End If
            Next

            DataGridView1.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 11, FontStyle.Bold Or FontStyle.Italic)
            DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            If Not ToolStrip2.Enabled Then ToolStrip2.Enabled = True
            'PrintToolStripButton.Enabled = True
            tbAxisX.Enabled = False
            tbAxisY.Enabled = False
            tbZoom.Enabled = False
            tbGraphType.Enabled = False
        End If

    End Sub

    Private Sub DisplayGraph()
        Dim cnt As Short, maxX As Double = 50.0
        Dim strBaseData As String()()

        Dim axisSectionSeries = New SectionsCollection
        axisSectionSeries.Add(New AxisSection With {
                .SectionWidth = 0,
                .StrokeThickness = 2,
                .Stroke = Brushes.Gray,
                .Value = 0
             })

        Dim YAxis = New Axis With {
            .LabelFormatter = Function(x) (x).ToString() & "m",
            .Title = "Depth (Mtr)",
            .Separator = New Separator With {
                .IsEnabled = True,
                .StrokeThickness = 1
            }
        }
        Dim XAxis = New Axis With {
            .Title = "Displacement (mm)",
            .LabelFormatter = Function(y) (Math.Round(y, 2)).ToString(),
            .MaxValue = 80,
            .MinValue = -80,
            .Separator = New Separator With {
                .IsEnabled = True,
                .[Step] = 10,
                .StrokeThickness = 1
            },
            .Sections = axisSectionSeries
        }
        Dim seriesCollection = New SeriesCollection

        ResetLabels()
        CartesianChart1.BackColor = System.Drawing.Color.White
        'CartesianChart1.Zoom = ZoomingOptions.X
        CartesianChart1.Series.Clear()
        CartesianChart1.AxisX.Clear()
        CartesianChart1.AxisY.Clear()
        CartesianChart1.AxisY.Add(YAxis)
        CartesianChart1.AxisX.Add(XAxis)
        If DataGridView1.Visible Then DataGridView1.Visible = False
        If CartesianChart1.Visible = False Then CartesianChart1.Visible = True
        If Not ToolStrip2.Enabled Then ToolStrip2.Enabled = True
        'PrintToolStripButton.Enabled = False
        tbAxisX.Enabled = True
        tbAxisY.Enabled = True
        tbZoom.Enabled = True
        tbGraphType.Enabled = True

        If tbGraphType.Text = "Deviation" Then
            If listBH(bhIndex).BaseFile Is Nothing Or listBH(bhIndex).BaseFile = "" Then
                MsgBox("No base file selected for this borehole. Go back and select a base file to view deviation.", vbOKOnly Or vbExclamation, "Graph")
                Return
            End If
            Dim strFile As String = GetBoreholeDirectory(boreHoleSelected) & "\" & listBH(bhIndex).BaseFile
            If Not IO.File.Exists(strFile) Then
                MsgBox("Base file does not exist. It must have been deleted. Please select another file as base.", vbOKOnly Or vbExclamation, "Graph")
                Return
            End If
            strBaseData = ReadCSVFile(strFile)
        End If

        For Each lstItem As String In lstBoreholes.SelectedItems
            Dim strData As String()() = ReadCSVFile(GetBoreholeDirectory(boreHoleSelected) & "\" & lstItem)
            Dim strFile = lstItem.Split(".").First().Replace("_", ":")

            If tbGraphType.Text = "Deviation" Then
                If strData.Length <> strBaseData.Length Then
                    MsgBox("Scale or length mismatch between Base file and Selected file.", vbExclamation Or vbOKOnly, "Graph")
                    Return
                End If
            End If
            'Dim invertedYMapper
            '= LiveCharts.Configurations.Mappers.Xy(Of ObservablePoint)().X(Function(point) point.Y).Y(Function(point) -point.X)
            Dim lineSeries = New VerticalLineSeries With {
                .Title = "[" & strFile & "]",
                .Values = New ChartValues(Of ObservablePoint),
                .Fill = Brushes.Transparent
            }
            Dim i As Short = 0, Val As Single, absVal As Single = 0, Val2 As Single, absVal2 As Single = 0
            For i = 4 To strData.Length - 1
                Val = (Single.Parse(strData(i)(1 + _axisValue)) -
                    Single.Parse(strData(i)(2 + _axisValue))) / 2
                If tbGraphType.Text = "Absolute" Then
                    Val += absVal
                    absVal = Val
                ElseIf tbGraphType.Text = "Deviation" Then
                    Val += absVal
                    absVal = Val
                    Val2 = (Single.Parse(strBaseData(i)(1 + _axisValue)) -
                        Single.Parse(strBaseData(i)(2 + _axisValue))) / 2
                    Val2 += absVal2
                    absVal2 = Val2
                    Val = absVal - absVal2
                End If
                If Math.Ceiling(Math.Abs(Val)) > maxX Then
                    maxX = Math.Ceiling(Math.Abs(Val))
                End If
                lineSeries.Values.Add(New ObservablePoint(Val, -Single.Parse(strData(i)(0))))

            Next

            maxX += maxX * 0.2
            maxX = Math.Ceiling(maxX)
            XAxis.MinValue = -maxX
            XAxis.MaxValue = maxX
            If maxX > 150 Then XAxis.Separator.Step = 40

            ' set the inverted mapping...
            'lineSeries.Configuration = invertedYMapper

            seriesCollection.Add(lineSeries)

            ' correct the labels
            'XAxis.LabelFormatter = Function(x) (x * -1).ToString() & "m"

            'Dim tooltip = New DefaultTooltip With {
            '    .SelectionMode = TooltipSelectionMode.OnlySender
            '}
            'CartesianChart1.DataTooltip = tooltip
            Select Case cnt
                Case 0
                    Label1.Text = strFile
                Case 1
                    Label2.Text = strFile
                Case 2
                    Label3.Text = strFile
                Case 3
                    Label4.Text = strFile
                Case 4
                    Label5.Text = strFile
            End Select
            cnt += 1
        Next
        If tbGraphType.SelectedItem = "Deviation" Then
            Label6.Text = "Base File : " & listBH(bhIndex).BaseFile.Split(".").First().Replace("_", ":")
        Else
            Label6.Text = ""
        End If
        CartesianChart1.Series = seriesCollection
    End Sub

    Private Sub tbAxisX_Click(sender As Object, e As EventArgs) Handles tbAxisX.Click
        If tbAxisY.Checked = True Then
            tbAxisY.Checked = False
            _axisValue = 0
            DisplayGraph()
        Else
            tbAxisX.Checked = True
        End If
    End Sub

    Private Sub tbAxisY_Click(sender As Object, e As EventArgs) Handles tbAxisY.Click
        If tbAxisX.Checked = True Then
            tbAxisX.Checked = False
            _axisValue = 2
            DisplayGraph()
        Else
            tbAxisY.Checked = True
        End If
    End Sub

    Private Sub tbViewGraph_Click(sender As Object, e As EventArgs) Handles tbViewGraph.Click
        If boreHoleSelected = 0 Then Return
        If lstBoreholes.SelectedItems.Count = 0 Then Return
        If lstBoreholes.SelectedItems.Count > 5 Then
            MsgBox("You have selected " & lstBoreholes.SelectedItems.Count & " files. You can select maximum 5 files for plotting graph", vbOKOnly Or vbExclamation, "Graph")
            Return
        End If
        DataGridView1.Visible = False
        DisplayGraph()
    End Sub

    Private Sub ResetLabels()
        Label1.Text = ""
        Label2.Text = ""
        Label3.Text = ""
        Label4.Text = ""
        Label5.Text = ""
        Label6.Text = "Select one or more files and click on View Graph"

        If bhIndex >= 0 Then
            lblBoreholeNumber.Text = "Borehole : " & boreHoleSelected.ToString().PadLeft(2, "0")
            lblDepth.Text = "Depth : " & listBH(bhIndex).Depth & "m"
            lblSiteName.Text = "Site : " & listBH(bhIndex).SiteName
            lblLocation.Text = "Location : " & listBH(bhIndex).Location
        Else
            lblBoreholeNumber.Text = ""
            lblDepth.Text = ""
            lblSiteName.Text = ""
            lblLocation.Text = ""
        End If
    End Sub

    Private Sub ToolBarEnable(ByRef enb As Boolean)
        Dim bnOneFileSelected As Boolean = (lstBoreholes.SelectedItems.Count = 1)
        tbBack.Enabled = enb
        If enb And lstBoreholes.SelectedItems.Count > 0 Then
            tbViewGraph.Enabled = True
            tbDelete.Enabled = True
            tbReport.Enabled = bnOneFileSelected
            tbBaseFile.Enabled = bnOneFileSelected
        Else
            tbViewGraph.Enabled = False
            tbDelete.Enabled = False
            tbBaseFile.Enabled = False
            tbReport.Enabled = False
        End If
    End Sub

    Private Sub lstBoreholes_Click(sender As Object, e As EventArgs) Handles lstBoreholes.Click
        ToolBarEnable(Not (boreHoleSelected = 0))
    End Sub

    Private Sub TbGraphType_SelChange(sender As Object, e As EventArgs) Handles tbGraphType.SelectedIndexChanged
        ' to be implemented
        DisplayGraph()
    End Sub

    Private Sub tbZoom_Click(sender As Object, e As EventArgs) Handles tbZoom.Click
        If tbZoom.Checked Then
            CartesianChart1.Zoom = ZoomingOptions.X
        Else
            CartesianChart1.Zoom = ZoomingOptions.None
        End If
    End Sub

    Private Sub tbDelete_Click(sender As Object, e As EventArgs) Handles tbDelete.Click
        If lstBoreholes.SelectedItems.Count = 0 Then Return

        If MsgBox("Are you sure you want to delete " & lstBoreholes.SelectedItems.Count & " selected file(s)?", vbYesNo Or vbQuestion, "Delete") = vbYes Then
            For Each strFile As String In lstBoreholes.SelectedItems
                IO.File.Delete(GetBoreholeDirectory(boreHoleSelected) & "\" & strFile)
            Next
            ReloadList()
        End If
    End Sub

    Private Sub PrintToolStripButton_Click(sender As Object, e As EventArgs) Handles PrintToolStripButton.Click
        If PrintDialog1.ShowDialog() <> DialogResult.OK Then Exit Sub
        PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings

        bsTextPrintData = "Borehole  : " & boreHoleSelected.ToString().PadLeft(2, "0") & vbCrLf
        bsTextPrintData &= "Depth     : " & listBH(bhIndex).Depth & "m" & vbCrLf
        bsTextPrintData &= "Site      : " & listBH(bhIndex).SiteName & vbCrLf
        bsTextPrintData &= "Location  : " & listBH(bhIndex).Location & vbCrLf
        bsTextPrintData &= "Date/Time : " & Label1.Text
        If CartesianChart1.Visible Then
            If Label2.Text <> "" Then bsTextPrintData &= ", " & Label2.Text
            If Label3.Text <> "" Then bsTextPrintData &= ", " & Label3.Text
            If Label4.Text <> "" Then bsTextPrintData &= ", " & Label4.Text
            If Label5.Text <> "" Then bsTextPrintData &= ", " & Label5.Text
            PrintDocument1.DefaultPageSettings.Margins.Left = 20
            PrintDocument1.DefaultPageSettings.Margins.Top = 20
            PrintDocument1.DefaultPageSettings.Margins.Right = 15
        Else
            PrintDocument1.DefaultPageSettings.Margins.Left = 90
            PrintDocument1.DefaultPageSettings.Margins.Top = 90
            PrintDocument1.DefaultPageSettings.Margins.Right = 75
        End If
        bsTextPrintData &= vbCrLf
        If Label6.Text <> "" Then bsTextPrintData &= Label6.Text & vbCrLf

        ' Report Printing code below
        printFont = New Font("Courier New", 9, FontStyle.Regular)
        PrintDocument1.DefaultPageSettings.Landscape = True
        If DataGridView1.Visible Then DisplayReport(True)
        'PrintDocument1.Print()

        'Show the Print Preview Dialog.
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.PrintPreviewControl.Zoom = 1
        PrintPreviewDialog1.ShowDialog()
        'PrintDocument1.Dispose()
    End Sub

    Private Sub tbReport_Click(sender As Object, e As EventArgs) Handles tbReport.Click
        DataGridView1.Visible = False
        CartesianChart1.Visible = False
        DisplayReport()
    End Sub

    Private Sub tbBaseFile_Click(sender As Object, e As EventArgs) Handles tbBaseFile.Click
        listBH(bhIndex).BaseFile = lstBoreholes.SelectedItem
        UpdateBorehole(listBH(bhIndex))
        ReloadList()
    End Sub

    Private Sub ListBox1_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) _
        Handles lstBoreholes.DrawItem

        ' Draw the background of the ListBox control for each item.
        e.DrawBackground()
        If e.Index < 0 Then Return

        ' Define the default color of the brush as black.
        Dim myBrush As System.Drawing.Brush = System.Drawing.Brushes.Chocolate

        'If bhIndex >= 0 Then
        'If listBH(bhIndex).BaseFile = lstBoreholes.Items(e.Index).ToString() Then
        'myBrush = System.Drawing.Brushes.Purple
        'Else
        'myBrush = System.Drawing.Brushes.Coral
        'End If
        'End If

        ' Draw the current item text based on the current 
        ' Font and the custom brush settings.
        e.Graphics.DrawString(lstBoreholes.Items(e.Index).ToString(), e.Font, myBrush, e.Bounds, StringFormat.GenericDefault)

        ' If the ListBox has focus, draw a focus rectangle around  _ 
        ' the selected item.
        e.DrawFocusRectangle()
    End Sub


    Private Sub PrintForm1_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim strFormat As New StringFormat()
        Dim rectDraw As New RectangleF(e.MarginBounds.Left, e.MarginBounds.Top, e.MarginBounds.Width, e.MarginBounds.Height)

        strFormat.Trimming = StringTrimming.Word
        If DataGridView1.Visible Then
            Dim numChars As Integer
            Dim numLines As Integer
            Dim stringForPage As String
            Dim sizeMeasure As New SizeF(e.MarginBounds.Width, e.MarginBounds.Height - printFont.GetHeight(e.Graphics))

            e.Graphics.MeasureString(bsTextPrintData, printFont, sizeMeasure, strFormat, numChars, numLines)
            stringForPage = bsTextPrintData.Substring(0, numChars)
            e.Graphics.DrawString(stringForPage, printFont, System.Drawing.Brushes.Black, rectDraw, strFormat)
            If numChars < bsTextPrintData.Length Then
                bsTextPrintData = bsTextPrintData.Substring(numChars)
                e.HasMorePages = True
            Else
                e.HasMorePages = False
            End If
        Else
            Dim MyChartPanel As Bitmap = New Bitmap(SplitContainer2.Panel1.Width, SplitContainer2.Panel1.Height)
            SplitContainer2.Panel1.DrawToBitmap(MyChartPanel, New Rectangle(0, 0, SplitContainer2.Panel1.Width, SplitContainer2.Panel1.Height))
            Dim p1 As Point
            p1.X = 5
            p1.Y = 110
            If MyChartPanel.Size.Width < e.PageBounds.Width Then
                p1.X = (e.PageBounds.Width - MyChartPanel.Size.Width) / 2
            End If
            e.Graphics.DrawString(bsTextPrintData, printFont, System.Drawing.Brushes.Black, rectDraw, strFormat)
            e.Graphics.DrawImage(MyChartPanel, p1)
            e.HasMorePages = False
        End If
    End Sub
End Class
