Attribute VB_Name = "Module1"
' This code includes 2 subs

' IMPORTANT : User need to launch sub (Allsheets)

'   - First one (Allsheets) it will select every Worksheet in the Workbook, and will apply (call) the second sub
'        * the code can be used whether the workbook has 1 or N sheets
'    - the second sub (vbachallenge) will perform all the needed transformation and calculation
'
'sub vbachallenge works as follows :
'1) all the processing is done in columns AA to AN
'2) the needed columns are copy/pasted  in columns I to Q at the end
'
' the processing works as follows
'1) on the basis of a distinct extraction of ticker value (using xls formula)
'2) we identify for each ticker the first and last date of trading
'3) we retrieve for each ticker the price of OPENING and CLOSING related to the dates identified earlier
'4) we calculate the price evolution during the period and the %
'5) we sum the column of stocks traded during the periode (using sumifs xls formula)
'6) we apply a conditional format on price evolution and %
'
'steps 2 to 6 are done by looping and use of XLS formulas
'
'7) we create the new table : Greatest % Increase, Greatest % Decrease, Greatest Total Volume
'8) we calculate the max and min of price % evolution
'9) we retrieve the ticker related using xls VLOOKUP formula
'
'10) we copy past (value and format) the needed columns fo the challenge to columns I to Q
'11) all the columns used for processing are deleted
'
'
'


' This sub AllSheets will select each worksheet of the workbook, and will apply (call) sub vbachallenge



Sub AllSheets()
    Dim ws As Worksheet
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
    
        ' Activate the worksheet
        ws.Activate
        
        ' Apply (Call) "vbachallenge" to the current sheet
        Call vbachallenge
    Next ws
    
End Sub


Sub vbachallenge()


' this step will give us number of rows (not empty) of the sheet and stock it in variabl lastrow
'       - we will use this variable to customize different formula

Dim lastrow As Long
lastrow = Range("A2").End(xlDown).Row
Range("aa1").Value = lastrow 'we are not obliged, but this step is usefull to control the code and follow its processing

'create headers
Cells(1, 28).Value = "Ticker"
Cells(1, 29).Value = "min date"
Cells(1, 30).Value = "max date"
Cells(1, 31).Value = "price open"
Cells(1, 32).Value = "price close"
Cells(1, 33).Value = "Quartly change"
Cells(1, 34).Value = "Percent change"
Cells(1, 35).Value = "Total stock volume"


'create distinct values of tickers using xls function UNIQUE on column A
Cells(2, 28).Formula2 = "=unique(A2:A" & lastrow & ")"

Columns("AB").Copy
    Columns("AB").PasteSpecial Paste:=xlPasteValues
    
    'identify the last row number to be used (same use as previously : to customize the formulas)

        Cells(2, 28).Select
        Dim lastdistinct As Long
        lastdistinct = Cells(2, 28).End(xlDown).Row
        Range("aa2").Value = lastdistinct 'we are not obliged, but this step is usefull to control the code and follow its processing

'Calculate Quartly change ; Percent change ; Total stock volume  Total stock volume
    
    'define variables
    
    Dim k As Integer
    
    Dim maxdate As Double
    Dim mindate As Double
    Dim Qchange As Double
    Dim Perchange As Double
    Dim Totstock As Double

    'loop to generate values
    '   this loop through the distinct tickers, will implement xls formulas in each cell
    
    For k = 1 To lastdistinct - 1
    
        'identify each max date and mindate for each ticker + store them in column 27,28 (AA, AB)
        ' this step is important is a stock was not traded during some days of the period
        
        Cells(k + 1, 29).Formula2 = "=minifs(B:B,A:A,AB" & (k + 1) & ")"
        Cells(k + 1, 30).Formula2 = "=maxifs(B:B,A:A,AB" & (k + 1) & ")"

        'find opening and closing prices and store them in column 29, 30 (AC, AD)
        ' using sumifs formula (cause there is only 1 value per day for each stock
        
        Cells(k + 1, 31).Formula2 = "=sumifs(C:C,A:A,AB" & (k + 1) & ",B:B,AC" & (k + 1) & ")"
        Cells(k + 1, 32).Formula2 = "=sumifs(F:F,A:A,AB" & (k + 1) & ",B:B,AD" & (k + 1) & ")"

         'COMPUTE difference between open and close prices for each row store it column AG
        Cells(k + 1, 33).Formula2 = "=AF" & (k + 1) & "-AE" & (k + 1)

            ' conditional format the cells
            '       - regarding if value is >0 (green) or > 0 (red)
            
            If Cells(k + 1, 33).Value > 0 Then
            Cells(k + 1, 33).Interior.Color = vbGreen

            ElseIf Cells(k + 1, 33).Value < 0 Then
            Cells(k + 1, 33).Interior.Color = vbRed

            End If


        'COMPUTE % store it column 34 (AH)

        'we set a condition if vale of AE =0
        ' if original extraction did selet all stocks even those not traded
        ' this will generate an error (divide by 0), so we replace by 0 and we do not calculate

            If Cells(k + 1, 31).Value = 0 Then
                Cells(k + 1, 34).Value = 0
                Else: Cells(k + 1, 34).Formula2 = "= (AF" & (k + 1) & "/AE" & (k + 1) & ")-1"
            End If

            ' conditional format the cells
            '       - regarding if value is >0 (green) or > 0 (red)

            If Cells(k + 1, 34).Value > 0 Then
            Cells(k + 1, 34).Interior.Color = vbGreen

            ElseIf Cells(k + 1, 34).Value < 0 Then
            Cells(k + 1, 34).Interior.Color = vbRed

            End If

         'compute total stock volume store and store in column 35 (AI)
        Cells(k + 1, 35).Formula2 = "=sumifs(G:G,A:A,AB" & (k + 1) & ")"

     Next k

'copy/past values of tickers
'   - we need this step to use a VLOOKUPin next steps

Columns("AB").Copy
Columns("AJ").PasteSpecial Paste:=xlPasteValues


'CREATE NEW VALUES

'CREATE HEADERS

Cells(1, 39).Value = "Ticker"
Cells(1, 40).Value = "Value"
Cells(2, 38).Value = "Greatest % Increase"
Cells(3, 38).Value = "Greatest % Decrease"
Cells(4, 38).Value = "Greatest Total Volume"
Cells(1, 41).Value = "Ticker" 'we need this second column to process a vlokup

'calculate values of the great % increase
Cells(2, 40).Formula2 = "=max(AH:AH)"

    'retrieve ticker Using VLOOKUP formula
    
    Cells(2, 41).Formula2 = "=vlookup(AN2,AH2:AJ" & lastdistinct & ",3,false)"
    Cells(2, 39).Value = Cells(2, 41).Value


'calculate values of the great & decrease
Cells(3, 40).Formula2 = "=min(AH:AH)"
    
    'retrieve ticker Using VLOOKUP formula
    
    Cells(3, 41).Formula2 = "=vlookup(AN3,AH2:AJ" & lastdistinct & ",3,false)"
    Cells(3, 39).Value = Cells(3, 41).Value

'calculate values great stock volume
Cells(4, 40).Formula2 = "=max(AI:AI)"

    'retrieve ticker Using VLOOKUP formula
       
    Cells(4, 41).Formula2 = "=vlookup(AN4,AI2:AJ" & lastdistinct & ",2,false)"
    Cells(4, 39).Value = Cells(4, 41).Value



'create requested table by copying the needed columns

Columns("AB").Copy
    Columns("I").PasteSpecial Paste:=xlPasteValues
    Columns("I").PasteSpecial Paste:=xlPasteFormats

Columns("AG").Copy
    Columns("J").PasteSpecial Paste:=xlPasteValues
    Columns("J").PasteSpecial Paste:=xlPasteFormats

Columns("AH").Copy
    Columns("K").PasteSpecial Paste:=xlPasteValues
    Columns("K").PasteSpecial Paste:=xlPasteFormats

Columns("AI").Copy
    Columns("L").PasteSpecial Paste:=xlPasteValues
    Columns("L").PasteSpecial Paste:=xlPasteFormats

Columns("AL:AN").Copy
    Columns("O:Q").PasteSpecial Paste:=xlPasteValues
    Columns("O:Q").PasteSpecial Paste:=xlPasteFormats



'format of numeric values

Range("J:J").NumberFormatLocal = "##0,#0"
Range("k:k").NumberFormatLocal = "0,00%"
Range("L:L").NumberFormatLocal = "# ##0"

Range("Q2").NumberFormatLocal = "0,00%"
Range("Q3").NumberFormatLocal = "0,00%"
Range("Q4").NumberFormatLocal = "# ##0"

' some fancy format, we are not buffles
Range("I1:Q1").Font.Bold = True
Range("O2:O4").Font.Bold = True

Columns("J:Q").AutoFit

'delete "one use" columns
'   - we can uncomment this line if we need to keep track of the calculations done

Columns("R:AZ").Delete

'back to a normal life
Range("I1").Activate

'thank you
 
End Sub
 

