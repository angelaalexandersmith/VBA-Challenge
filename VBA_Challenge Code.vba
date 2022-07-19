Sub Vba_Challenge()

Dim i As LongLong

Dim Ticker As String

Dim Yearly_Change As Double

    Yearly_Change = 0

Dim Percentage_Change As Variant
    
    Percentage_Change = 0

Dim Total_Volume As LongLong

    Total_Volume = 0

Dim Summary_Table_Row As Integer

Summary_Table_Row = 2


Dim Year_opening_price, Year_closing_price As Double

i = 2

Dim sngOpenPrice As Double
    
    sngOpenPrice = 0
    
Dim sngClosePrice As Double
    
    sngClosePrice = 0
    

For i = 2 To 753001

Ticker = Cells(i, 1).Value

sngClosePrice = Cells(i, 6).Value

sngOpenPrice = Cells(i, 3).Value


Yearly_Change = sngClosePrice - sngOpenPrice

Total_Volume = Total_Volume + Cells(i, 7).Value

Percentage_Change = Yearly_Change / sngOpenPrice

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    ' if the ticker changes
    
    sngClosePrice = Cells(i, 6).Value
              
    Range("I" & Summary_Table_Row).Value = Ticker

    Range("J" & Summary_Table_Row).Value = Yearly_Change

    Range("K" & Summary_Table_Row).Value = Percentage_Change

    Range("L" & Summary_Table_Row).Value = Total_Volume

    Summary_Table_Row = Summary_Table_Row + 1

    Yearly_Change = 0

    Percentage_Change = 0

    Total_Volume = 0
    
    sngClosePrice = 0
    
    Else
    
    sngOpenPrice = sngOpenPrice
    
    sngOpenPrice = 0
    
    
    
    

 

 End If

 Next i




End Sub


