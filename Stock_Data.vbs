Sub Lite()
    'This first sub I named "lite" as a first step to make sure I could break down the solution and see if I could get the first 2 lines working'
    
    Dim x As Double
    Dim Total As Double
    
    Dim TotalV As Double
    
    'I made a Clear all statement to make sure all lines between columns I:Q were clear'
        Columns("I:Q").Select
        Selection.Clear
    'These two lines are to print out the header for the two columns'
        Cells(1, 9).Value = Cells(1, 1).Value
        Cells(1, 10).Value = "Total Stock Value"
    'The next few lines contain an If statement to print out the ticker if not it will print out total value'
    x = 2
    Cells(x, 9).Value = Cells(x, 1).Value
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To LastRow
    
    If Cells(i, 1).Value = Cells(x, 9).Value Then
    
    TotalV = TotalV + Cells(i, 7).Value
    
         Else
         
    Cells(x, 10).Value = TotalV
    
    TotalV = Cells(i, 7).Value
    
    x = x + 1
    Cells(x, 9).Value = Cells(i, 1).Value
    
    
    
    
    End If
        
        Next i
    
    Cells(x, 10).Value = TotalV
        
        
    'resize'
    Columns("I:Q").EntireColumn.AutoFit
    ''
    Cells(1, 1).Select

End Sub


Sub Moderate()

    'This second sub I made was a second step to see if I could print the "yearly change" along with the "Percent Change"'
    
    Dim DateMinOpen As Variant
    Dim DateMaxClose As Variant
    Dim i As Double
    
    
    Dim x As Double
    
    'Like in the forst sub these next lines are to clear column I to Q'
        Columns("I:Q").Select
        Selection.Clear
    'This is for the Headers on the moderate button'
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
    ''''''''''
    x = 2
    i = 2
    
    Cells(x, 9).Value = Cells(x, 1).Value
    
    DateMinOpen = Cells(i, 3).Value
    
    
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
    
    
    
    'end setup'
    
    
    If Cells(i, 1).Value = Cells(x, 9).Value Then
    
    
    TotalV = TotalV + Cells(i, 7).Value
    
    
    DateMaxClose = Cells(i, 6).Value
    
    
         Else
         
    
    'The next few lines of If statements are for calculated fields'
    Cells(x, 10).Value = DateMaxClose - DateMinOpen
    
                    If DateMaxClose <= 0 Then
                
                        Cells(x, 11).Value = 0
                        
                        Else
    
                        Cells(x, 11).Value = (DateMaxClose / DateMinOpen) - 1
                        
    
                    End If
                    
                        Cells(x, 11).Style = "Percent"
                            
                If Cells(x, 10).Value >= 0 Then
                                    
                    Cells(x, 10).Interior.ColorIndex = 4
                                        
                        Else
                                    
                    Cells(x, 10).Interior.ColorIndex = 3
                        
                End If
                    
    Cells(x, 12).Value = TotalV
    
    'The next few lines are to reset the variables'
    
    DateMinOpen = Cells(i, 3).Value
    
    TotalV = Cells(i, 7).Value
    
    x = x + 1
    Cells(x, 9).Value = Cells(i, 1).Value
    
    End If
    
    Next i
    
    'calculated fields final'
    Cells(x, 10).Value = DateMaxClose - DateMinOpen
    
                    If DateMaxClose <= 0 Then
                
                        Cells(x, 11).Value = 0
                        
                        Else
    
                        Cells(x, 11).Value = (DateMaxClose / DateMinOpen) - 1
                        
    
                    End If
                    
                        Cells(x, 11).Style = "Percent"
                            
                If Cells(x, 10).Value >= 0 Then
                                    
                    Cells(x, 10).Interior.ColorIndex = 4
                                        
                        Else
                                    
                    Cells(x, 10).Interior.ColorIndex = 3
                        
                End If
                    
    Cells(x, 12).Value = TotalV
    
    
    'resize the next line is to resize the column (Won't lie i looked up this Line, couldn't figure it out on my own :('
    Columns("I:Q").EntireColumn.AutoFit
    
    Cells(1, 1).Select

End Sub


Sub Full()
'This sub is the complete assignment with the "greatest Increase" "Greatest Decrease" and "Greatest Total Volume"'



    For Each ws In Worksheets
            Dim WorksheetName As String
            WorksheetName = ws.Name
            
            Sheets(ws.Name).Select
    
    
    Dim DateMinOpen As Variant
    Dim DateMaxClose As Variant
    Dim i As Double
    
    
    Dim x As Double
    
    'Once again these two lines are to clear columns I to Q'
        Columns("I:Q").Select
        Selection.Clear
    'Next few lines are for the headers'
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Volume"
    
    x = 2
    i = 2
    
    Cells(x, 9).Value = Cells(x, 1).Value
    
    DateMinOpen = Cells(i, 3).Value
    
    
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
    
    
    
    
    
    
    
    
    If Cells(i, 1).Value = Cells(x, 9).Value Then
    
    
    TotalV = TotalV + Cells(i, 7).Value
    
    
    DateMaxClose = Cells(i, 6).Value
    
    
         Else
         
    
    
    Cells(x, 10).Value = DateMaxClose - DateMinOpen
    
                    If DateMaxClose <= 0 Then
                
                        Cells(x, 11).Value = 0
                        
                        Else
                        
    
                        Cells(x, 11).Value = (DateMaxClose / DateMinOpen) - 1
                        
    
                        
                    End If
                    
                        Cells(x, 11).Style = "Percent"
                            
                If Cells(x, 10).Value >= 0 Then
                                    
                    Cells(x, 10).Interior.ColorIndex = 4
                                        
                        Else
                                    
                    Cells(x, 10).Interior.ColorIndex = 3
                        
                End If
                    
    Cells(x, 12).Value = TotalV
    
    
    DateMinOpen = Cells(i, 3).Value
    
    TotalV = Cells(i, 7).Value
    
    x = x + 1
    Cells(x, 9).Value = Cells(i, 1).Value
    
    End If
    
    Next i
    
    
    Cells(x, 10).Value = DateMaxClose - DateMinOpen
    
                    If DateMaxClose <= 0 Then
                
                        Cells(x, 11).Value = 0
                        
                        Else
    
                        Cells(x, 11).Value = (DateMaxClose / DateMinOpen) - 1
                        
    
                    End If
                    
                        Cells(x, 11).Style = "Percent"
                            
                If Cells(x, 10).Value >= 0 Then
                                    
                    Cells(x, 10).Interior.ColorIndex = 4
                                        
                        Else
                                    
                    Cells(x, 10).Interior.ColorIndex = 3
                        
                End If
                    
    Cells(x, 12).Value = TotalV
    
    '''Start Hard Section''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Volume_Greatest_Decrease = 100000
            Ticker_Greatest_Decrease = 100000
            
            LastRow = Cells(Rows.Count, 9).End(xlUp).Row
            
            For x = 2 To LastRow
            
            
            If Cells(x, 11).Value > Volume_Greatest_Increase Then
                
                Ticker_Greatest_Increase = Cells(x, 9).Value
                Volume_Greatest_Increase = Cells(x, 11).Value
            
            End If
            
            
            If Cells(x, 11).Value < Volume_Greatest_Decrease Then
                
                Ticker_Greatest_Decrease = Cells(x, 9).Value
                Volume_Greatest_Decrease = Cells(x, 11).Value
            
            End If
            
            
            If Cells(x, 12).Value > Volume_Greatest_Total_Volume Then
                
                Ticker_Greatest_Total_Volume = Cells(x, 9).Value
                Volume_Greatest_Total_Volume = Cells(x, 12).Value
            
            End If
            
            Next x
            
    Cells(2, 16).Value = Ticker_Greatest_Increase
    Cells(2, 17).Value = Volume_Greatest_Increase
    Cells(2, 17).Style = "Percent"
    Cells(3, 16).Value = Ticker_Greatest_Decrease
    Cells(3, 17).Value = Volume_Greatest_Decrease
    Cells(3, 17).Style = "Percent"
    Cells(4, 16).Value = Ticker_Greatest_Total_Volume
    Cells(4, 17).Value = Volume_Greatest_Total_Volume
    'resize''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Columns("I:Q").EntireColumn.AutoFit
    
    Cells(1, 1).Select
    
    Next ws


End Sub
Sub reset()
'This sub is to reset all the worksheets'
    For Each ws In Worksheets
        Dim WorksheetName As String
        WorksheetName = ws.Name
        
                Sheets(ws.Name).Select
        'Clear all'
        Columns("I:Q").Select
        Selection.Clear
        'resize'
        Columns("I:Q").EntireColumn.AutoFit
            Cells(1, 1).Select
    
    Next ws
    
End Sub

