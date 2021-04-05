Option Explicit

' 도면화 모듈 1단

Sub DWG_Data입력_도면생성()



'우측 하단 도각 입력 사항 정리
    Dim shtData As Worksheet
     Set shtData = ActiveSheet
    
    Dim strJN As String
    
    Dim strPName As String
    
    Dim strEQName As String
    
    Dim strDWGName As String
    
    Dim strNOofEQ As String
    
    Dim strDesign As String
    
    Dim strCheck As String
    
    Dim strQv As String
     
    Dim strTag1st As String

     
    
    
    Dim shtDWG_1st As Worksheet
     Set shtDWG_1st = Worksheets("도면_Single")
    Dim shtDWG_New_1st As Worksheet
    
    
    Dim strHole_1 As String
    
    Dim strPi_1 As String
    
    Dim strPo_1 As String
    
    Dim strdP_1 As String
   
    Dim strML_1 As String
    
    Dim strDim_D As String
    
    Dim strDim_B As String
  
    Dim strFileName As String
    
    
    With shtData
        
        strJN = .Range("E14")
        
            strPName = .Range("E16")
        
            strEQName = .Range("E17")
        
            strDWGName = "ORIFICE DWG FOR " & strEQName & "(Single)"
        
            strNOofEQ = .Range("E18")
        
            strDesign = .Range("E19")
        
            strCheck = .Range("E20")
        
            strQv = .Range("T17")
        
            strTag1st = .Range("C44")
        
            strPi_1 = .Range("H44")
            
            strPo_1 = .Range("J44")
            
            strdP_1 = .Range("L44")
            
            strML_1 = .Range("N44")
            
            strHole_1 = .Range("F44")
            
            strDim_D = .Range("Z20")
            
            strDim_B = .Range("AA20")
            
            
            
            strFileName = "DWG_" & strTag1st & "_1단"
            
        
    End With
    

     

    
        Dim i As Integer
            For i = 1 To ThisWorkbook.Sheets.Count
                If Sheets(i).Name = strFileName Then
                  strFileName = "Sheets" & ThisWorkbook.Sheets.Count + 1
                  MsgBox "동일한 Item에 대한 Sheet가 있어 Sheet명을 임의로 부여합니다." & "(Sheet" & ThisWorkbook.Sheets.Count + 1 & ")"
                End If
            Next i
                
 
     shtDWG_1st.Copy after:=shtData
    
     Set shtDWG_New_1st = ActiveSheet
     
     With shtDWG_New_1st
     
         .Name = strFileName
         .Range("F38") = strTag1st
         .Range("H38") = strHole_1
         .Range("I38") = strDim_D
         .Range("J38") = strDim_B
         .Range("K38") = "FRONT" & Chr(10) & strTag1st & Chr(10) & "HoleDia :Φ" & strHole_1 & "㎜"
         .Range("M38") = "BACK"
         .Range("O38") = strML_1 & " ㎜"
         .Range("Q38") = strNOofEQ
         .Range("Q42") = strJN
         .Range("Q44") = strPName
         .Range("Q46") = strTag1st
         .Range("Q49") = strDWGName
         .Range("Q50") = Year(Now) & "/" & Month(Now) & "/" & Day(Now)
         .Range("Q51") = strDesign
         .Range("Q52") = strCheck
         
     End With
     
     
    
    
   
End Sub


