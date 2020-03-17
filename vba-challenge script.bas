Attribute VB_Name = "Module2"
Sub symbol():
    Dim Worksheetname As String
    Dim ticker As String
    Dim vol As Double
        vol = 0
    Dim summary_row As Integer
    Dim year_open As Double
    Dim year_close As Double
    Dim last_row As Double
    Dim yearly_change As Double
    Dim yearly_percentage As Double


    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly_change"
    Cells(1, 12).Value = "Yearly_percentage"
    Cells(1, 13).Value = "Total Stock Vol"

    Worksheetname = ActiveSheet.Name
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        MsgBox (LastRow)
    summary_row = 2
    vol = 0
    year_open = Cells(2, 3).Value
For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            vol = vol + Cells(i, 7).Value
            Cells(summary_row, 10).Value = ticker
            Cells(summary_row, 13).Value = vol
            year_close = Cells(i, 6).Value

            yearly_change = year_close - year_open
                Cells(summary_row, 11).Value = yearly_change
                summary_row = summary_row + 1

                vol = 0
            yearly_percentage = yearly_change / year_open
         Else
             vol = vol + Cells(i, 7).Value

        End If
        
Next i

End Sub
