VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysisData()

'Declare and set worksheet
'Dim ws As Worksheet
'Set ws = ThisWorkbook.Worksheets("StockAnalysisData")

'   Declare the variables
    
    Dim ticker As String
    Dim yearlychange As Double
    Dim percentagechange As Double
    Dim totalstockvolume As Double
    Dim lastrow As Long
    Dim summaryrow As Long
    Dim openprice As Double
    Dim closeprice As Double
    Dim i As Long

' Initialize values
    totalstockvolume = 0
    yearlychange = 0
    percentagechange = 0
    summaryrow = 2
    
    
'Add Header Name to Display Data
    Range("I1").Value = "Ticker"
    Range("J1").Value = "YearlyChange"
    Range("K1").Value = "PercentChange"
    Range("L1").Value = "TotalStockVolume"
    

'Find the last row of data
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'   Loop through all rows of data

    For i = 2 To lastrow
    
        ticker = Cells(i, 1).Value
        openprice = Cells(i, 3).Value
        closeprice = Cells(i, 6).Value
        
 
  ' Track if ticker changes then print results
    If Cells(i + 1, 1).Value <> ticker Then
        
        yearlychange = closeprice - openprice
        percentagechange = yearlychange / openprice * 100
        totalstockvolume = Cells(i, 12).Value
        
   ' calculating yearlychange, percentagechange and totalstockvolume.
     
        Cells(summaryrow, 9).Value = ticker
        Cells(summaryrow, 10).Value = yearlychange
        Cells(summaryrow, 11).Value = percentagechange
        Cells(summaryrow, 12).Value = totalstockvolume
        
        summaryrow = summaryrow + 1
        totalstockvolume = 0
 
        Else
       
            totalstockvolume = totalstockvolume + Cells(i, 12).Value
     
     End If
   
    Next i
    

End Sub


