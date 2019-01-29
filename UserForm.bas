Dim a1995(7) As Double 'population in 1995 for 7 countries
Dim a2010(7) As Double 'population in 2010 for 7 countries

'Hide frmCalculate
Private Sub cmdExit_Click()
    Hide
End Sub

Private Sub cmdOK_Click()

'update data
lstSelectedCountries.Clear

'calculate the number of selected items
count = 0
For i = 0 To lstEntry.ListCount - 1
    If lstEntry.Selected(i) Then count = count + 1
Next

If count < 2 Then
    MsgBox ("Please, select more then 2 countries (Ctrl + Country)")
    Exit Sub
End If

Application.ScreenUpdating = False
Index = 120

'fill the box numer 3 and worksheet
For i = 0 To lstEntry.ListCount - 1
    If lstEntry.Selected(i) Then
    
        Cells(Index, 1) = Trim(lstEntry.List(i))
        Cells(Index, 2) = Round((a2010(i) - a1995(i)) / a1995(i) * 100)
        
        lstSelectedCountries.AddItem lstEntry.List(i) + vbTab + Str(Cells(Index, 2)) + "%" + vbTab + Str(a1995(i)) + vbTab + Str(a2010(i))
        Index = Index + 1
    End If
Next

    'select the range on the worksheet with data for the pie chart (from A120)
    start = 120
    sRange = "Sheet1!$A$" + Trim(Str(start)) + ":$B$" + Trim(Str(start + count - 1))
    
    'create pie chart on on the worksheet
    
    ActiveSheet.Shapes.AddChart(5, xlPie).Select
    ActiveChart.SetSourceData Source:=Range(sRange)
    ActiveChart.ApplyLayout (2)
    ActiveChart.HasTitle = False

   
    'save pie chart in the file (in current directory)
    Dim imageName As String
    imageName = ThisWorkbook.Path & "\temp.gif"
    ActiveChart.Export Filename:=imageName
    
    'load the picture from the file to form
    frmCalculate.Image1.Picture = LoadPicture(imageName)
    
    'remove pie chart from the worksheet
    ActiveSheet.ChartObjects(1).Delete
    
    'remove data for pie chart from the worksheet
    Range(sRange).Clear
    
    Application.ScreenUpdating = True
   
End Sub



'filling the box number 1 and arrays with population of 7 countries
Private Sub UserForm_Initialize()
lstSelectedCountries.Clear
lstEntry.AddItem "KUWAIT         "
lstEntry.AddItem "UNITED STATES  "
lstEntry.AddItem "UK             "
lstEntry.AddItem "SAUDI ARABIA   "
lstEntry.AddItem "OMAN           "
lstEntry.AddItem "IRAQ           "

'KUWAIT
a1995(0) = Cells(59, 2)
a2010(0) = Cells(59, 17)

'UNITED STATES
a1995(1) = Cells(116, 2)
a2010(1) = Cells(116, 17)

'UK
a1995(2) = Cells(83, 2)
a2010(2) = Cells(83, 17)

'SAUDI ARABIA
a1995(3) = Cells(63, 2)
a2010(3) = Cells(63, 17)

'OMAN
a1995(4) = Cells(61, 2)
a2010(4) = Cells(61, 17)

'IRAQ
a1995(5) = Cells(56, 2)
a2010(5) = Cells(56, 17)

End Sub
