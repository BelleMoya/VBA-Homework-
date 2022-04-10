Attribute VB_Name = "Module1"
Option Explicit

Sub NewCode()
'
' NewCode Macro

'Part 1: Create a loop to sort through the ticker and have placed in column "I"
'Part 2:Find the yearly price change by subtracting the opening value from the the closing value
'Part 3: Find the percentage change for each stock
'Part 4:Find the total stock volume
'Part 5: Conditional format yearly change; green-positive change red-negative change



    Cells(1, 9) = "Ticker "
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Volume"
    
    
    Dim Ticker As String
    Dim Total_Volume As Double
    Dim Summary_Row As Integer
    Dim Start_Value As Double
    Dim End_Value As Double
    Dim LastRow As Long
    Dim i As Double
    
    

LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Total_Volume = 0
Summary_Row = 2
Start_Value = Cells(2, 3).Value


For i = 2 To LastRow




If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
    
    End_Value = Cells(i, 6).Value
    
    
    Ticker = Cells(i, 1).Value
    
    


    Range("I" & Summary_Row).Value = Ticker
    
    Range("L" & Summary_Row).Value = Total_Volume
    
    Range("J" & Summary_Row).Value = End_Value - Start_Value
    
   If (Range("J" & Summary_Row) > 0) Then
                    
            
            Range("J" & Summary_Row).Interior.ColorIndex = 4
                  
                ElseIf (Range("J" & Summary_Row) <= 0) Then
                
                    
                    Range("J" & Summary_Row).Interior.ColorIndex = 3
                End If

If Start_Value = 0 Then
    
    Range("K" & Summary_Row).Value = (0 & "%")
     
        Else
     
            Range("K" & Summary_Row).Value = ((End_Value - Start_Value) / (Start_Value) * 100) & "%"
         End If
         
         
            Summary_Row = Summary_Row + 1
            
            Total_Volume = 0
     
            Start_Value = Cells(i + 1, 3)
            

    Else
           Total_Volume = Total_Volume + Cells(i, 7).Value
           
           End If
           
           
   
           
           
    Next i
    


    
 
    
End Sub



