Attribute VB_Name = "Module1"
'#
'# Dan Cusick
'# Data Science Boot Camp
'#
Sub vbaofwallstreet()
   Dim ws As Worksheet
   
   Dim ticker As String
   Dim priceopen As Double
   Dim priceclose As Double
   Dim pricechange As Double
   Dim percentchange As Double
   Dim volume As Double
   
   Dim biggestwinner As String
   Dim biggestwinnernum As Double
   Dim biggestloser As String
   Dim biggestlosernum As Double
   Dim largestvolume As String
   Dim largestvolumenum As Double
   
   Dim ResultRowIndex As Long
   Dim PassIndex As Long
   
   biggestwinnernum = 0
   biggestlosernum = 0
   largestvolumenum = 0
   
   
   '# For each ws in Worksheets
   For Each ws In Worksheets
   
     '# Worksheets(ws).Activate
     LastRowIndex = ws.Cells(Rows.Count, "A").End(xlUp).Row
   
     '# Label Cells I,J,K,L
     ws.Cells(1, "I").Value = "Ticker"
     ws.Cells(1, "J").Value = "Yearly Change"
     ws.Cells(1, "K").Value = "Percent Change"
     ws.Cells(1, "L").Value = "Total Stock Volume"
     
   
     '# Set Indexes
     ResultRowIndex = 1
     PassIndex = 1
   
     '# Init Volume
     volume = 0
      
     '# Process a sheet
     For i = 2 To LastRowIndex '+ 1
       
       '# Test for first row of a symbol, assign ticker, open
       If PassIndex = 1 Then
         
         '# Reset the Ticker volume, priceopen and passindex
         ticker = ws.Cells(i, "A").Value
         priceopen = ws.Cells(i, "C").Value
         PassIndex = 0
       
       End If
     
       '# Sum Volume
       volume = ws.Cells(i, "G").Value + volume
       
       If (ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value) Then

         '# Increment Result Row Index
         ResultRowIndex = ResultRowIndex + 1
         
         '# Calculate Yearly Change
         pricechange = ws.Cells(i, "F").Value - priceopen
         
         '# Calculate the Percent Change but do not divide by 0
         If priceopen = 0 Then
           percentchange = 100
         Else
           percentchange = ((pricechange / priceopen) * 100)
         End If
         
         '# Output Ticker Result
         ws.Cells(ResultRowIndex, "I").Value = ticker
         ws.Cells(ResultRowIndex, "J").Value = Format(pricechange, "#,##0.00")
         ws.Cells(ResultRowIndex, "K").Value = Format(percentchange, "% " + "0.00")
         ws.Cells(ResultRowIndex, "L").Value = volume
         
         '# Sum Greatest Increase, Decrease and Volume
         If pricechange > biggestwinnernum Then
           biggestwinnernum = percentchange
           biggestwinner = ticker
         End If
         If pricechange < biggestlosernum Then
           biggestlosernum = percentchange
           biggestloser = ticker
         End If
         If volume > largestvolumenum Then
           largestvolumenum = volume
           largestvolume = ticker
         End If
         
         volume = 0
         PassIndex = 1
         
       End If  '# Roll to next Ticker Symbol
  
     Next i
     ws.Columns("I:L").AutoFit
 
   Next ws  '# Step to next sheet
   
   '# Label Greatest % winner, % loser and volume
   Worksheets("A").Activate
   Worksheets("A").Cells(1, "P").Value = "Ticker"
   Worksheets("A").Cells(1, "Q").Value = "Value"
   
   Worksheets("A").Cells(2, "O").Value = "Greatest % Increase"
   Worksheets("A").Cells(3, "O").Value = "Greatest % Decrease"
   Worksheets("A").Cells(4, "O").Value = "Greatest Total Volume"
   
   Worksheets("A").Cells(2, "P").Value = biggestwinner
   Worksheets("A").Cells(2, "Q").Value = Format(biggestwinnernum, "#,##0.00") + "%"
   
   Worksheets("A").Cells(3, "P").Value = biggestloser
   Worksheets("A").Cells(3, "Q").Value = Format(biggestlosernum, "#,##0.00") + "%"
   
   Worksheets("A").Cells(4, "P").Value = largestvolume
   Worksheets("A").Cells(4, "Q").Value = largestvolumenum
   
   '# Handy ditty this debug.print ditty to immediate window!
   '# Debug.Print "Greatest % Decrease ", biggestloser, " ", biggestlosernum
   '# Debug.Print "Greatest Total Volume", largestvolume, " ", largestvolumenum
   Worksheets("A").Columns("O:Q").AutoFit
End Sub
