Attribute VB_Name = "ModuleCalc"
Sub Calc(o, t)
    Dim i, j As Integer ' Rows
    Dim t1 As Integer ' Temperature
    Dim m1 As Integer ' Mass
    Dim b As Boolean
    Dim y As Integer
    
    i = 1
    Do Until o.Cells(i, 1) = ""
        i = i + 1
    Loop
    i = i - 1
    o.Cells(i, 1) = ""
    o.Cells(i, 2) = ""
    i = i - 1
    
    ' Change Number format
    o.Range(o.Cells(3, 2), o.Cells(i, 6)).NumberFormat = "General"
    
    ' Normalize Mass
    o.Cells(1, 7) = "Mass Perc."
    o.Cells(2, 7) = "[%]"
    For j = 3 To i
        o.Cells(j, 7) = o.Cells(j, 5) / o.Cells(3, 5) * 100
    Next
    
    ' Calculate Mass Loss
    o.Cells(1, 8) = "Mass lost"
    o.Cells(2, 8) = "[mg]"
    For j = 3 To i
        o.Cells(j, 8) = o.Cells(3, 5) - o.Cells(j, 5)
    Next
    
    ' Save maximum percentage value
    o.Range("J1") = "Mass perc.Max"
    o.Range("J2") = "[%]"
    o.Range("J3") = Application.WorksheetFunction.Max(o.Range(o.Cells(3, 7), o.Cells(i, 7)))
    
    ' Normalization according to maximum weight
    o.Cells(1, 9) = "Mass normalized"
    o.Cells(2, 9) = "[%]"
    For j = 3 To i
        o.Cells(j, 9) = o.Cells(j, 7) / o.Range("J3") * 100
    Next
    
    'saving the minimum normalized value for minimum scale on y-axis
    Z = Application.WorksheetFunction.Min(Range(Cells(3, 9), Cells(j, 9)))
    
    'Setting names and values for the Wt.OA-value
    o.Range("L1") = "T"
    o.Range("L2") = "°C"
    o.Range("L3") = "150"
    o.Range("L4") = t
    o.Range("M1") = "Mass"
    o.Range("M2") = "[mg]"
    o.Range("N1") = "Wt.% OA"
    o.Range("N2") = "[%]"
    
    j = 3
    Do Until o.Cells(j, 6) = 150
        j = j + 1
    Loop
    o.Range("M3") = o.Cells(j, 5)
    
    j = 3
    Do Until o.Cells(j, 6) = t
        j = j + 1
    Loop
    o.Range("M4") = o.Cells(j, 5)

    'Wt.OA -Value
    o.Range("N3") = (o.Range("M3") - o.Range("M4")) / o.Range("M3") * 100
    o.Range("N1:N3").Font.Bold = True
    
'    'Calculate the most appropriate minimum scale for the y-axis of the upcoming diagram
    b = False
    y = 95
    Do While y >= 0 And b = False
        If Z >= y And Z < y + 5 Then
            b = True
        Else
            y = y - 5
        End If
    Loop
    
    
    'Add a xy-diagram
    Set co = o.ChartObjects.Add(550, 70, 800, 500)
    Set c = co.Chart
    
    'Setting the y-lines on the y-axis, if the minimum is >95: 95,96,...100; if >90: 92,94...100; else: 75,80,85...100
    If y = 95 Then
        x = 1
    ElseIf y = 90 Then
        x = 2
    Else
        x = 5
    End If
    
    'Define parameters of the diagram
    With c
        .ChartType = xlXYScatter
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = o.Range(Cells(3, 6), Cells(i, 6))            'x-values
        .SeriesCollection(1).Values = o.Range(Cells(3, 9), Cells(i, 9))             'y-values
        .SeriesCollection(1).Name = "Mass Perc. normalized"
        .HasLegend = False                                                          'no Legend name
        .HasTitle = yes
        .ChartTitle.Characters.Text = o.Name                                        'Adding the title, it is the name of the map
        .Axes(xlCategory, xlPrimary).HasTitle = True                                'Setting the name of the x-axis
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Temperature [°C]" 'Name of the x-axis
        .Axes(xlCategory, xlPrimary).MinimumScale = 0
        .Axes(xlCategory, xlPrimary).MaximumScale = o.Cells(i, 6)                   'Scale of the y-axis, it is the last value of the temperature data
        .Axes(xlValue, xlPrimary).MajorUnit = 100                                   'primary counting on x-axis
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Mass normalized [%]" 'name of the y-axis
        .Axes(xlValue, xlPrimary).MinimumScale = y                                  'minimum scale as calculated above
        .Axes(xlValue, xlPrimary).MaximumScale = 100                                'maximum scale is 100%, no calculation needed
        .Axes(xlValue, xlPrimary).MajorUnit = x                                     'primary step is calculated as described above
    End With
End Sub
