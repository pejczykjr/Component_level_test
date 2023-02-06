Attribute VB_Name = "Charts"
Option Explicit

'This Sub adds charts and centers the whole spreadsheet.
Sub dataFormat(ByRef src As Workbook, measurementFileName As String)

'   VARIABLES

    'These keep information about automatic range of values on y-axis.
    Dim maxScale, minScale As String

'   FUNCTIONAL PART

    'Adjusts height and width of all active cells in worksheet.
    With ActiveSheet.UsedRange
        .EntireRow.AutoFit
        .EntireColumn.AutoFit
        .HorizontalAlignment = xlCenter
    End With
    
    'Adding chart for Insertion loss.
    If (InStr(1, measurementFileName, "il")) <> 0 Then
        
        'Selects a random empty cell. It is performed to create empty chart.
        ActiveSheet.Range("F7").Select
        ActiveSheet.Shapes.AddChart2(227, xlXYScatterLinesNoMarkers).Select

        With ActiveChart
            
            'Shows legend of the chart, sets its position to bottom and sets title for the chart.
            .HasLegend = True
            .Legend.Position = xlBottom
            .ChartTitle.Text = "Insertion Loss " + Range("F3").Value
            
            'Adds 2 series to chart.
            .SeriesCollection.NewSeries
            .SeriesCollection.NewSeries
            
            'First series which is limit line. Names it with value from C1 cell. Here are also added
            'X values for limit range to cover x-axis with frequency values. Line
            'is coloured with red. Thickness of line is 0.75.
            With .SeriesCollection(1)
                .Name = Range("C1").Value
                .XValues = "='" & measurementFileName & "'!$A$2:$A$" + CStr(Cells(Rows.Count, 1).End(xlUp).Row) 'x
                .Values = "='" & measurementFileName & "'!$C$2:$C$" + CStr(Cells(Rows.Count, 3).End(xlUp).Row) 'y
                .Format.Line.Weight = 0.75
                .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
            End With
    
            'Second series which is measurement line. Names it with value from F3 cell. Here are also added
            'X values for measurement range to cover x-axis with frequency values. Thickness of line is 0.75
            With .SeriesCollection(2)
                .Name = Range("F3").Value
                .XValues = "='" & measurementFileName & "'!$A$2:$A$" + CStr(Cells(Rows.Count, 1).End(xlUp).Row) 'x
                .Values = "='" & measurementFileName & "'!$B$2:$B$" + CStr(Cells(Rows.Count, 2).End(xlUp).Row) 'y
                .Format.Line.Weight = 0.75
                
                'Setting colour for line.
                With .Format.Line.ForeColor
                    If (InStr(1, measurementFileName, "brown")) <> 0 Then
                        .RGB = RGB(153, 76, 0)
                    ElseIf (InStr(1, measurementFileName, "green")) <> 0 Then
                        .RGB = RGB(0, 255, 0)
                    ElseIf (InStr(1, measurementFileName, "orange")) <> 0 Then
                        .RGB = RGB(255, 153, 51)
                    ElseIf (InStr(1, measurementFileName, "blue")) <> 0 Then
                        .RGB = RGB(0, 0, 255)
                    End If
                End With
            End With
    
            'It shows x-axis title and limits x-axis values to last frequency.
            With .Axes(xlCategory)
                .HasTitle = True
                .AxisTitle.Text = "Frequency [MHz]"
                .MaximumScale = Cells(Cells(Rows.Count, 1).End(xlUp).Row, 1).Value
            End With
            
            'It shows y-axis title.
            With .Axes(xlValue)
                .HasTitle = True
                .AxisTitle.Text = "Power ratio [dB]"
            End With
                
            'Sets scale of the chart and places it in F7 cell (left top corner in there).
            With ActiveSheet.Shapes("Chart 1")
                .ScaleWidth 1.7243055556, msoFalse, msoScaleFromTopLeft
                .ScaleHeight 1.3252314815, msoFalse, msoScaleFromTopLeft
                .Top = Range("F7").Top
                .Left = Range("F7").Left
            End With
            
        End With
            
    'Adding chart for NEXT forward.
    ElseIf (InStr(1, measurementFileName, "next")) <> 0 Then
        
        'Selects a random empty cell. It is performed to create empty chart.
        Range("J7").Select
        ActiveSheet.Shapes.AddChart2(227, xlXYScatterLinesNoMarkers).Select
        
        With ActiveChart
            
            'Shows legend of the chart, sets its position to bottom and sets title for the chart.
            .HasLegend = True
            .Legend.Position = xlBottom
            'In J3 there is e.g. Blue to Orange, we take Blue (first one).
            .ChartTitle.Text = "NEXT forward " + Left(Range("J3").Value, InStr(Range("J3").Value, " ") - 1) + "-to-All"
            
            'Adds 4 series to chart.
            .SeriesCollection.NewSeries
            .SeriesCollection.NewSeries
            .SeriesCollection.NewSeries
            .SeriesCollection.NewSeries
            'This is fifth, because it will be displayed as a line at 10MHz (we don't consider values at less frequency).
            .SeriesCollection.NewSeries
            
            'First series which is limit line. Names it with value from E1 cell. Here are also added
            'X values for limit range to cover x-axis with frequency values. Line
            'is coloured with red. Thickness of line is 0.75.
            With .SeriesCollection(1)
                .Name = Range("E1").Value
                .XValues = "='" & measurementFileName & "'!$A$2:$A$" + CStr(Cells(Rows.Count, 1).End(xlUp).Row) 'x
                .Values = "='" & measurementFileName & "'!$E$2:$E$" + CStr(Cells(Rows.Count, 5).End(xlUp).Row) 'y
                .Format.Line.Weight = 0.75
                .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
            End With
            
            'Second series which is measurement line. Names it with value from J3 cell. Here are also added
            'X values for measurement range to cover x-axis with frequency values. Thickness of line is 0.75
            With .SeriesCollection(2)
                .Name = Range("J3").Value
                .XValues = "='" & measurementFileName & "'!$A$2:$A$" + CStr(Cells(Rows.Count, 1).End(xlUp).Row) 'x
                .Values = "='" & measurementFileName & "'!$B$2:$B$" + CStr(Cells(Rows.Count, 2).End(xlUp).Row) 'y
                .Format.Line.Weight = 0.75
                
                'Setting colour for line.
                With .Format.Line.ForeColor
                    If Range("J3").Value Like "*to Brown*" Then
                        .RGB = RGB(153, 76, 0)
                    ElseIf Range("J3").Value Like "*to Green*" Then
                        .RGB = RGB(0, 255, 0)
                    ElseIf Range("J3").Value Like "*to Orange*" Then
                        .RGB = RGB(255, 153, 51)
                    ElseIf Range("J3").Value Like "*to Blue*" Then
                        .RGB = RGB(0, 0, 255)
                    End If
                End With
            End With
            
            'Third series which is measurement line. Names it with value from L3 cell. Here are also added
            'X values for measurement range to cover x-axis with frequency values. Thickness of line is 0.75
            With .SeriesCollection(3)
                .Name = Range("L3").Value
                .XValues = "='" & measurementFileName & "'!$A$2:$A$" + CStr(Cells(Rows.Count, 1).End(xlUp).Row) 'x
                .Values = "='" & measurementFileName & "'!$C$2:$C$" + CStr(Cells(Rows.Count, 3).End(xlUp).Row) 'y
                .Format.Line.Weight = 0.75
                
                'Setting colour for line.
                With .Format.Line.ForeColor
                    If Range("L3").Value Like "*to Brown*" Then
                        .RGB = RGB(153, 76, 0)
                    ElseIf Range("L3").Value Like "*to Green*" Then
                        .RGB = RGB(0, 255, 0)
                    ElseIf Range("L3").Value Like "*to Orange*" Then
                        .RGB = RGB(255, 153, 51)
                    ElseIf Range("L3").Value Like "*to Blue*" Then
                        .RGB = RGB(0, 0, 255)
                    End If
                End With
            End With
            
            'Fourth series which is measurement line. Names it with value from N3 cell. Here are also added
            'X values for measurement range to cover x-axis with frequency values. Thickness of line is 0.75
            With .SeriesCollection(4)
                .Name = Range("N3").Value
                .XValues = "='" & measurementFileName & "'!$A$2:$A$" + CStr(Cells(Rows.Count, 1).End(xlUp).Row) 'x
                .Values = "='" & measurementFileName & "'!$D$2:$D$" + CStr(Cells(Rows.Count, 4).End(xlUp).Row) 'y
                .Format.Line.Weight = 0.75
                
                'Setting colour for line.
                With .Format.Line.ForeColor
                    If Range("N3").Value Like "*to Brown*" Then
                        .RGB = RGB(153, 76, 0)
                    ElseIf Range("N3").Value Like "*to Green*" Then
                        .RGB = RGB(0, 255, 0)
                    ElseIf Range("N3").Value Like "*to Orange*" Then
                        .RGB = RGB(255, 153, 51)
                    ElseIf Range("N3").Value Like "*to Blue*" Then
                        .RGB = RGB(0, 0, 255)
                    End If
                End With
            End With

            'Gets scale of y-axis.
            maxScale = CStr(.Axes(xlValue).MaximumScale)
            minScale = CStr(.Axes(xlValue).MinimumScale)
            
            'This is fifth series which is straight line at 10MHz.
            With .SeriesCollection(5)
                .XValues = "={10,10}"
                .Values = "={" + minScale + "," + maxScale + "}"
                .Format.Line.Weight = 0.25
                .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
                .ApplyDataLabels
                .Points(2).DataLabel.delete
                .HasLeaderLines = False
                
                'Shows only Point(1) which is one at the bottom (there was also Point(2) at the top, but it
                'was unnecessary, it is being deleted in With above). Also here displays argument on x-axis
                'and hides value.
                With .Points(1).DataLabel
                    .Position = xlLabelPositionBelow
                    .ShowCategoryName = True
                    .ShowValue = False
                End With
    
            End With
            
            'It deletes SeriesCollection(5) from chart legend.
            .Legend.LegendEntries(5).delete
            'Sets y-axis values back to defalut (after adding a line it was changing).
            .Axes(xlValue).MaximumScale = maxScale
            .Axes(xlValue).MinimumScale = minScale
        
            'It shows x-axis title and limits x-axis values to last frequency.
            With .Axes(xlCategory)
                .HasTitle = True
                .AxisTitle.Text = "Frequency [MHz]"
                .MaximumScale = Cells(Cells(Rows.Count, 1).End(xlUp).Row, 1).Value
                .TickLabelPosition = xlLow
            End With
            
            'It shows y-axis title.
            With .Axes(xlValue)
                .HasTitle = True
                .AxisTitle.Text = "Power ratio [dB]"
            End With
                
            'Sets scale of the chart and places it in J7 cell (left top corner in there).
            With ActiveSheet.Shapes("Chart 1")
                .ScaleWidth 1.7243055556, msoFalse, msoScaleFromTopLeft
                .ScaleHeight 1.3252314815, msoFalse, msoScaleFromTopLeft
                .Top = Range("J7").Top
                .Left = Range("J7").Left
            End With
        
        End With
        
    'Adding chart for Return Loss.
    ElseIf (InStr(1, measurementFileName, "rl")) <> 0 Then
    
        'Selects a random empty cell. It is performed to create empty chart.
        Range("L7").Select
        ActiveSheet.Shapes.AddChart2(227, xlXYScatterLinesNoMarkers).Select
        
        With ActiveChart
            
            'Shows legend of the chart, sets its position to bottom and sets title for the chart.
            .HasLegend = True
            .Legend.Position = xlBottom
            'If fw measurement, set Forward to name. In other case set Reverse.
            If (InStr(1, measurementFileName, "fw")) <> 0 Then
                .ChartTitle.Text = "Return Loss Forward"
            ElseIf (InStr(1, measurementFileName, "rev")) <> 0 Then
                .ChartTitle.Text = "Return Loss Reverse"
            End If
            
            'Adds 5 series to chart.
            .SeriesCollection.NewSeries
            .SeriesCollection.NewSeries
            .SeriesCollection.NewSeries
            .SeriesCollection.NewSeries
            .SeriesCollection.NewSeries
            
            'First series which is limit line. Names it with value from F1 cell. Here are also added
            'X values for limit range to cover x-axis with frequency values. Line
            'is coloured with red. Thickness of line is 0.75.
            With .SeriesCollection(1)
                .Name = Range("F1").Value
                .XValues = "='" & measurementFileName & "'!$A$2:$A$" + CStr(Cells(Rows.Count, 1).End(xlUp).Row) 'x
                .Values = "='" & measurementFileName & "'!$F$2:$F$" + CStr(Cells(Rows.Count, 6).End(xlUp).Row) 'y
                .Format.Line.Weight = 0.75
                .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
            End With
            
            'Second series which is measurement line. Names it with value from L3 cell. Here are also added
            'X values for measurement range to cover x-axis with frequency values. Thickness of line is 0.75
            With .SeriesCollection(2)
                .Name = Range("L3").Value
                .XValues = "='" & measurementFileName & "'!$A$2:$A$" + CStr(Cells(Rows.Count, 1).End(xlUp).Row) 'x
                .Values = "='" & measurementFileName & "'!$B$2:$B$" + CStr(Cells(Rows.Count, 2).End(xlUp).Row) 'y
                .Format.Line.Weight = 0.75

                'Setting colour for line.
                With .Format.Line.ForeColor
                    If Range("L3").Value Like "*Brown*" Then
                        .RGB = RGB(153, 76, 0)
                    ElseIf Range("L3").Value Like "*Green*" Then
                        .RGB = RGB(0, 255, 0)
                    ElseIf Range("L3").Value Like "*Orange*" Then
                        .RGB = RGB(255, 153, 51)
                    ElseIf Range("L3").Value Like "*Blue*" Then
                        .RGB = RGB(0, 0, 255)
                    End If
                End With
            End With
            
            'Third series which is measurement line. Names it with value from N3 cell. Here are also added
            'X values for measurement range to cover x-axis with frequency values. Thickness of line is 0.75
            With .SeriesCollection(3)
                .Name = Range("N3").Value
                .XValues = "='" & measurementFileName & "'!$A$2:$A$" + CStr(Cells(Rows.Count, 1).End(xlUp).Row) 'x
                .Values = "='" & measurementFileName & "'!$C$2:$C$" + CStr(Cells(Rows.Count, 3).End(xlUp).Row) 'y
                .Format.Line.Weight = 0.75
                
                'Setting colour for line.
                With .Format.Line.ForeColor
                    If Range("N3").Value Like "*Brown*" Then
                        .RGB = RGB(153, 76, 0)
                    ElseIf Range("N3").Value Like "*Green*" Then
                        .RGB = RGB(0, 255, 0)
                    ElseIf Range("N3").Value Like "*Orange*" Then
                        .RGB = RGB(255, 153, 51)
                    ElseIf Range("N3").Value Like "*Blue*" Then
                        .RGB = RGB(0, 0, 255)
                    End If
                End With
            End With
            
            'Fourth series which is measurement line. Names it with value from P3 cell. Here are also added
            'X values for measurement range to cover x-axis with frequency values. Thickness of line is 0.75
            With .SeriesCollection(4)
                .Name = Range("P3").Value
                .XValues = "='" & measurementFileName & "'!$A$2:$A$" + CStr(Cells(Rows.Count, 1).End(xlUp).Row) 'x
                .Values = "='" & measurementFileName & "'!$D$2:$D$" + CStr(Cells(Rows.Count, 4).End(xlUp).Row) 'y
                .Format.Line.Weight = 0.75
                
                'Setting colour for line.
                With .Format.Line.ForeColor
                    If Range("P3").Value Like "*Brown*" Then
                        .RGB = RGB(153, 76, 0)
                    ElseIf Range("P3").Value Like "*Green*" Then
                        .RGB = RGB(0, 255, 0)
                    ElseIf Range("P3").Value Like "*Orange*" Then
                        .RGB = RGB(255, 153, 51)
                    ElseIf Range("P3").Value Like "*Blue*" Then
                        .RGB = RGB(0, 0, 255)
                    End If
                End With
            End With
            
            'Fifth series which is measurement line. Names it with value from R3 cell. Here are also added
            'X values for measurement range to cover x-axis with frequency values. Thickness of line is 0.75
            With .SeriesCollection(5)
                .Name = Range("R3").Value
                .XValues = "='" & measurementFileName & "'!$A$2:$A$" + CStr(Cells(Rows.Count, 1).End(xlUp).Row) 'x
                .Values = "='" & measurementFileName & "'!$E$2:$E$" + CStr(Cells(Rows.Count, 5).End(xlUp).Row) 'y
                .Format.Line.Weight = 0.75
                
                'Setting colour for line.
                With .Format.Line.ForeColor
                    If Range("R3").Value Like "*Brown*" Then
                        .RGB = RGB(153, 76, 0)
                    ElseIf Range("R3").Value Like "*Green*" Then
                        .RGB = RGB(0, 255, 0)
                    ElseIf Range("R3").Value Like "*Orange*" Then
                        .RGB = RGB(255, 153, 51)
                    ElseIf Range("R3").Value Like "*Blue*" Then
                        .RGB = RGB(0, 0, 255)
                    End If
                End With
            End With
            
            'It shows x-axis title and limits x-axis values to last frequency.
            With .Axes(xlCategory)
                .HasTitle = True
                .AxisTitle.Text = "Frequency [MHz]"
                .MaximumScale = Cells(Cells(Rows.Count, 1).End(xlUp).Row, 1).Value
                .TickLabelPosition = xlLow
            End With
            
            'It shows y-axis title.
            With .Axes(xlValue)
                .HasTitle = True
                .AxisTitle.Text = "Power ratio [dB]"
            End With
                
            'Sets scale of the chart and places it in J7 cell (left top corner in there).
            With ActiveSheet.Shapes("Chart 1")
                .ScaleWidth 1.7243055556, msoFalse, msoScaleFromTopLeft
                .ScaleHeight 1.3252314815, msoFalse, msoScaleFromTopLeft
                .Top = Range("L7").Top
                .Left = Range("L7").Left
            End With
            
        End With
    
    End If

    'Closes the source file.
    'True - saves the source file.
    src.Close True
    Set src = Nothing

End Sub
