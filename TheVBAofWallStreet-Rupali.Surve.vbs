'Below VBA script is comprised of the homework assignment including challenge activity.

Sub VBA_StockAnalysis()
'this code works - satisfies homework requirement and challege 2
'to repeat the ticker summary calculation for each sheet in the workbook
'This is part of challenge 2
For Each ws In Worksheets
        ws.Activate
        
          
        'calling the sub routine which calculates ticker summary
            Call TickerSummary
            'Call StockSummaryChallengeSolution
    Next ws
   
End Sub


Sub TickerSummary()
'Trying with one sheet
'Declaring variables to hold current ticker name, next ticker name, ticker summary count and total volume per ticker
Dim Thisticker As String
Dim NextTicker As String
Dim Summaryticker As LongLong

'setting up start of ticker at row 2 as 1st row is a header.
Summaryticker = 2

Dim Tickerstockvolume As LongLong
'Tickerstockvolume = 0


'Counting total number of rows. This will help to stop counting after the last row.
Dim Totalrows As LongLong
Totalrows = Cells(Rows.Count, 1).End(xlUp).Row


'defining variables to hold current ticker open and close values for yearly and percent change
    
    Dim Thistickeropen As Double
    Dim Thistickerclose As Double
    Dim Yearlychange As Double
    Dim Percentchange As Double

'setting initial value for ticker open value
Thistickeropen = Cells(2, 3).Value


'-----end of declare


' defining a loop so we can go through ticker in the current row, check the next row ticker value
'if both values match then raise current ticker count by 1 and add current ticket volume to the Tickerstockvolume
'if values don`t match then restart the loop at new value for the ticker and calculate summary per same logic


For i = 2 To Totalrows
    
    
    'check the details of the ticker from the curernt row
    Thisticker = Cells(i, 1).Value
    
    'check the details for the ticker from the next row
    NextTicker = Cells(i + 1, 1).Value
    
    Tickerstockvolume = Tickerstockvolume + Cells(i, 7).Value
    
       
    'check if the value of the current cell and next cell are different so can write to summary
       
    If Thisticker <> NextTicker Then
            
        
    'when ticker value changes, calculate summery and write to the summary table/columns
    
    'Cells(Summaryticker, 9).Value = Thisticker
    'Cells(Summaryticker, 12).Value = Tickerstockvolume

    Range("I" & Summaryticker).Value = Thisticker
    Range("L" & Summaryticker).Value = Tickerstockvolume
    
    '----------------This is the start for yearly change and percent change calculations---------------------------------------
            
     Thistickerclose = Cells(i, 6).Value

    
    'calculate yearly change which is (close price at the end of the year - open price at the beginning of the year)
    
           Yearlychange = Thistickerclose - Thistickeropen


    
    'write yearlychange to the summary table/column
    Range("J" & Summaryticker).Value = Yearlychange
    
    'calculate % change ...set % change = 0 if open price = 0 to address mathematical problem with % calculation
    '% change here will be (yearlychange)/Open)*100 ---since using number format to covert using yearlychange/Open
    
    If Thistickeropen = 0 Then
        Percentchange = 0
    Else
        Percentchange = Yearlychange / Thistickeropen
    End If
    'Debug.Print Percentchange
    
    'Write the % change for each ticker in the summary table
        Range("K" & Summaryticker).Value = Percentchange
        Range("K" & Summaryticker).NumberFormat = "0.00%"

    '----------------This is the end for yearly change and percent change calculations---------------------------------
    
    'since now we have summary for the current ticker, write it to the summary table and advance to the next row for the ticker
    
    Summaryticker = Summaryticker + 1
    
    'since now we have new ticker, resetting the total
    
    Tickerstockvolume = 0

    'reset open value for next ticker
    Thistickeropen = Cells(i + 1, 3).Value
    

'            Percentchange = 0
'            Tickerstockvolume = 0
                        
    
    
    End If
Next i
      

'Challenge 1 - To calculate Greatest % increase, Greatest % decrease and Greatest total volume of all tickers,
'if Percentchange is -ve then price decreased, if +ve means price increased
'to accommodate this logic using maximum and minimum values from the percentchange
    
Dim TotRowsSummary As LongLong
    TotRowsSummary = Cells(Rows.Count, 9).End(xlUp).Row
    
        For i = 2 To TotRowsSummary
    
        If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & TotRowsSummary)) Then
            Cells(2, 16).Value = Cells(i, 9).Value
            Cells(2, 17).Value = Cells(i, 11).Value
            Cells(2, 17).NumberFormat = "0.00%"
       
        ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & TotRowsSummary)) Then
            Cells(3, 16).Value = Cells(i, 9).Value
            Cells(3, 17).Value = Cells(i, 11).Value
            Cells(3, 17).NumberFormat = "0.00%"
       
        ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & TotRowsSummary)) Then
            Cells(4, 16).Value = Cells(i, 9).Value
            Cells(4, 17).Value = Cells(i, 12).Value
        
        End If
        
        Next i

    
        
      'calling the subroutines which formats summary data
            Call Summarytitles
            Call conditionalformat
            Call arrow
      
End Sub


Sub Summarytitles()

'Defining new column headers for the ticker summary.

    'clearing cell values and formatting
'    Range("I:Q").Value = ""
'    Range("I:Q").Interior.ColorIndex = 0
     
    'for the ticker summary
    Range("i1").Value = "Ticker"
    Range("j1").Value = "Yearly Change"
    Range("k1").Value = "Percent Change"
    Range("l1").Value = "Total Stock Volume"
    
    Range("i1").Font.ColorIndex = 5
    Range("i1").Interior.ColorIndex = 36
    Range("j1").Font.ColorIndex = 5
    Range("j1").Interior.ColorIndex = 36
    Range("k1").Font.ColorIndex = 5
    Range("k1").Interior.ColorIndex = 36
    Range("l1").Font.ColorIndex = 5
    Range("l1").Interior.ColorIndex = 36
    Range("i1:l1").BorderAround (1)
    Range("i1:l1").Borders.LineStyle = xlContinuous
    Range("i1:l1").Borders.ColorIndex = 0
    Range("i1:l1").Borders.TintAndShade = 0
    Range("i1:l1").Borders.Weight = xlThin
        
    'for the challenge 1
    
    Range("p1").Value = "Ticker"
    Range("q1").Value = "Value"
    Range("o2").Value = "Greatest % Increase"
    Range("o3").Value = "Greatest % Decrease"
    Range("o4").Value = "Greatest Total Volume"
    
    Range("p1").Font.ColorIndex = 5
    Range("p1").Interior.ColorIndex = 36
    Range("q1").Font.ColorIndex = 5
    Range("q1").Interior.ColorIndex = 36
    Range("o2").Font.ColorIndex = 5
    Range("o2").Interior.ColorIndex = 36
    Range("o3").Font.ColorIndex = 5
    Range("o3").Interior.ColorIndex = 36
    Range("o4").Font.ColorIndex = 5
    Range("o4").Interior.ColorIndex = 36
    Range("o1:q4").BorderAround (1)
    Range("o1:q4").Borders.LineStyle = xlContinuous
    Range("o1:q4").Borders.ColorIndex = 0
    Range("o1:q4").Borders.TintAndShade = 0
    Range("o1:q4").Borders.Weight = xlThin

    Range("I:Q").Columns.AutoFit
    


'Range("i1").Font.ColorIndex = 5
'Range("i1").Interior.ColorIndex = 6
'Range("i1").Font.Size = 15
'Range("i1").Font.Name = "Calibri"
'Range("i1").ColumnWidth = 20
'Range("i1").Font.Bold = True

End Sub

'Conditional formatting to highlight positive change in green and negative change in red.

Sub conditionalformat()

Dim summaryrows As LongLong
         
     'find the last row of the summary table


    summaryrows = Cells(Rows.Count, 9).End(xlUp).Row
    
    'set color depending on yearly change
    
    For i = 2 To summaryrows
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
    Next i

End Sub

'formating for object arrow
Sub arrow()

    ActiveSheet.Shapes.AddShape(msoShapeRightArrow, 850, 91.5, 80, 22.5).Select
    Selection.ShapeRange.IncrementRotation 327.8691666667
    Selection.ShapeRange.IncrementLeft -5.2500787402
    Selection.ShapeRange.IncrementTop -7.5000787402
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(187, 33, 21)
        .Transparency = 0
        .Solid
    End With
    With Selection.ShapeRange.Shadow
        .Type = msoShadow25
        .Visible = msoTrue
        .Style = msoShadowStyleOuterShadow
        .Blur = 4
        .OffsetX = 2.4492935983E-16
        .OffsetY = 4
        .RotateWithShape = msoFalse
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.150000006
        .Transparency = 0
    End With
    With Selection.ShapeRange.Shadow
        .Type = msoShadow25
        .Visible = msoTrue
        .Style = msoShadowStyleOuterShadow
        .Blur = 4
        .OffsetX = 2.4492935983E-16
        .OffsetY = 4
        .RotateWithShape = msoFalse
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
        .Size = 100
    End With
    With Selection.ShapeRange.Shadow
        .Type = msoShadow25
        .Visible = msoTrue
        .Style = msoShadowStyleOuterShadow
        .Blur = 4
        .OffsetX = 2.4492935983E-16
        .OffsetY = 4
        .RotateWithShape = msoFalse
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.150000006
        .Transparency = 0
        .Size = 100
    End With
    Selection.ShapeRange.Reflection.Type = msoReflectionType1
    Selection.ShapeRange.Reflection.Type = msoReflectionTypeNone
    With Selection.ShapeRange.Glow
        .Color.RGB = RGB(255, 0, 0)
        .Transparency = 0
        .Radius = 10
    End With
    With Selection.ShapeRange.Glow
        .Color.ObjectThemeColor = msoThemeColorAccent4
        .Color.TintAndShade = 0
        .Color.Brightness = 0.8000000119
        .Transparency = 0
        .Radius = 10
    End With
    With Selection.ShapeRange.Glow
        .Color.ObjectThemeColor = msoThemeColorAccent2
        .Color.TintAndShade = 0
        .Color.Brightness = 0
        .Transparency = 0.6000000238
        .Radius = 5
    End With
    With Selection.ShapeRange.Glow
        .Color.ObjectThemeColor = msoThemeColorBackground1
        .Color.TintAndShade = 0
        .Color.Brightness = -0.150000006
        .Transparency = 0.6000000238
        .Radius = 5
    End With
    Application.CommandBars("Format Object").Visible = False
    Selection.ShapeRange.ThreeD.ContourColor.RGB = RGB(255, 0, 0)
    Range("N10").Select
End Sub

