Option Explicit

Public Function FindDWG(X, Y, Xa As Range, Xb As Range, Ya As Range, Yb As Range, PI As Range)
  Application.ScreenUpdating = False
  Application.Volatile False

  
  Dim ixa As Integer
  Dim ixb As Integer
  Dim iya As Integer
  Dim iyb As Integer
  Dim ipi As Integer
   ixa = Xa.Cells.Count
   ixb = Xb.Cells.Count
   iya = Ya.Cells.Count
   iyb = Yb.Cells.Count
   ipi = PI.Cells.Count
  Dim idxa As Integer
  Dim idxb As Integer
  Dim idya As Integer
  Dim idyb As Integer
  Dim idpi As Integer
  
  Dim strTemp As String
    strTemp = ""
   
  Dim strXa() As String
    ReDim strXa(0 To ixa - 1)
    
  Dim strXb() As String
    ReDim strXb(0 To ixb - 1)
    
  Dim strYa() As String
    ReDim strYa(0 To iya - 1)
    
  Dim strYb() As String
    ReDim strYb(0 To iyb - 1)
    
  Dim strPI() As String
    ReDim strPI(0 To ipi - 1)
     
  Dim Cxa As Range
  Dim Cxb As Range
  Dim Cya As Range
  Dim Cyb As Range
  Dim Cpi As Range
  
  For Each Cxa In Xa
    strXa(idxa) = Cxa.Value
    idxa = idxa + 1
  Next Cxa
  
  For Each Cxb In Xb
    strXb(idxb) = Cxb.Value
    idxb = idxb + 1
  Next Cxb
  
  For Each Cya In Ya
    strYa(idya) = Cya.Value
    idya = idya + 1
  Next Cya
   
  For Each Cyb In Yb
    strXa(idyb) = Cyb.Value
    idyb = idyb + 1
  Next Cyb
    
  For Each Cpi In PI
    strPI(idpi) = Cpi.Value
    idpi = idpi + 1
  Next Cpi
    
    
  Dim i As Integer
   For i = 0 To ixa - 1
      If X >= strXa(i) And X <= strXb(i) And Y <= strYa(i) And Y >= strYb(i) And strTemp = "" Then
          strTemp = strPI(i)
      ElseIf X >= strXa(i) And X <= strXb(i) And Y <= strYa(i) And Y >= strYb(i) And strTemp <> "" Then
         strTemp = strTemp & "/" & strPI(i)
         
      End If
   Next
  
  FindDWG = strTemp
  Erase strPI
  Erase strXa
  Erase strXb
  Erase strYa
  Erase strYb
     
End Function

Public Function Add(X As Range, Y As Double)
  Application.ScreenUpdating = False
  Application.Volatile
  Dim strX() As String
  Dim intX As Integer
  Dim i As Integer
  Dim c As Range
  intX = X.Cells.Count
  ReDim strX(0 To intX - 1)
   For Each c In X
    strX(i) = c.Value
    i = i + 1
   Next c
  Add = strX(0) + Y
          
End Function
Public Function FindDWGS(X, Y, Xa As Range, Xb As Range, Ya As Range, Yb As Range, PI As Range)
  Application.ScreenUpdating = False
  Application.Volatile False

  
  Dim ixa As Integer
  Dim ixb As Integer
  Dim iya As Integer
  Dim iyb As Integer
  Dim ipi As Integer
   ixa = Xa.Cells.Count
   ixb = Xb.Cells.Count
   iya = Ya.Cells.Count
   iyb = Yb.Cells.Count
   ipi = PI.Cells.Count
  Dim idxa As Integer
  Dim idxb As Integer
  Dim idya As Integer
  Dim idyb As Integer
  Dim idpi As Integer
  
  Dim strTemp As String
    strTemp = ""
   
  Dim strXa() As String
    ReDim strXa(0 To ixa - 1)
    
  Dim strXb() As String
    ReDim strXb(0 To ixb - 1)
    
  Dim strYa() As String
    ReDim strYa(0 To iya - 1)
    
  Dim strYb() As String
    ReDim strYb(0 To iyb - 1)
    
  Dim strPI() As String
    ReDim strPI(0 To ipi - 1)
     
  Dim Cxa As Range
  Dim Cxb As Range
  Dim Cya As Range
  Dim Cyb As Range
  Dim Cpi As Range
  
  For Each Cxa In Xa
    strXa(idxa) = Cxa.Value
    idxa = idxa + 1
  Next Cxa
  
  For Each Cxb In Xb
    strXb(idxb) = Cxb.Value
    idxb = idxb + 1
  Next Cxb
  
  For Each Cya In Ya
    strYa(idya) = Cya.Value
    idya = idya + 1
  Next Cya
   
  For Each Cyb In Yb
    strXa(idyb) = Cyb.Value
    idyb = idyb + 1
  Next Cyb
    
  For Each Cpi In PI
    strPI(idpi) = Cpi.Value
    idpi = idpi + 1
  Next Cpi
    
    
  Dim i As Integer
   For i = 0 To ipi - 1
      If X >= strXa(i) And X <= strXb(i) And Y <= strYa(i) And Y >= strYb(i) Then
          strTemp = strPI(i)
         
      End If
   Next
  
  FindDWGS = strTemp
  Erase strPI
  Erase strXa
  Erase strXb
  Erase strYa
  Erase strYb
     
End Function
