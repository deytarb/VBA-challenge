Attribute VB_Name = "Module1"
   Sub test()
    
    
  
Dim ws As Worksheet
Dim Header As Boolean
Header = False
Dim SHT As Boolean
SHT = True
    
    
For Each ws In Worksheets
    
' mains labels
Dim labels As String
labels = " "
        
Dim alltickers As Double
alltickers = 0
       
Dim opn As Double
Dim cls As Double
Dim yc As Double
Dim perchng As Double
       
Dim higher As String
Dim lower As String
Dim higherpct As Double
Dim lowerpct As Double
Dim Volhigher As String
Dim volmx As Double
       
        
opn = 0
cls = 0
yc = 0
perchng = 0
higher = " "
lower = " "
higherpct = 0
lowerpct = 0
Volhigher = " "
volmx = 0
        
        
Dim alltable As Long
alltable = 2
        
Dim last As Long
Dim i As Long
        
last = ws.Cells(Rows.Count, 1).End(xlUp).Row

If Header Then

 'where put the labels
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
Else
Header = True
End If
              
opn = ws.Cells(2, 3).Value
        For i = 2 To last
        
      
     
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
labels = ws.Cells(i, 1).Value
             
              
cls = ws.Cells(i, 6).Value
yc = cls - opn

If opn <> 0 Then
perchng = (yc / opn) * 100
Else
                 
End If
                

alltickers = alltickers + ws.Cells(i, 7).Value
              
                
ws.Range("I" & alltable).Value = labels
ws.Range("J" & alltable).Value = yc
                
If (yc > 0) Then

ws.Range("J" & alltable).Interior.ColorIndex = 4
ElseIf (yc <= 0) Then

ws.Range("J" & alltable).Interior.ColorIndex = 3
End If
ws.Range("K" & alltable).Value = (CStr(perchng) & "%")
ws.Range("L" & alltable).Value = alltickers
         
               
alltable = alltable + 1

yc = 0
               
cls = 0
opn = ws.Cells(i + 1, 3).Value
              
If (perchng > higherpct) Then
higherpct = perchng
higher = labels
ElseIf (perchng < lowerpct) Then
lowerpct = perchng
lower = labels
End If
                 
If (alltickers > volmx) Then
volmx = alltickers
Volhigher = labels
End If
                
perchng = 0
alltickers = 0
                
Else

alltickers = alltickers + ws.Cells(i, 7).Value
End If
    
Next i

            
If Not SHT Then
            
ws.Range("Q2").Value = (CStr(higherpct) & "%")
ws.Range("Q3").Value = (CStr(lowerpct) & "%")
ws.Range("P2").Value = higher
ws.Range("P3").Value = lower
ws.Range("Q4").Value = volmx
ws.Range("P4").Value = Volhigher
                
Else
SHT = False
End If
        
Next ws
     
End Sub

