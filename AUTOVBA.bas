Attribute VB_Name = "Module1"
Sub filtre6()
Attribute filtre6.VB_ProcData.VB_Invoke_Func = "i\n14"
    
    Dim i       As Long
    Dim j       As Long
    Dim dic     As Object
    Dim arrEle  As Variant
    Dim arrData As Variant
    Set dic = CreateObject("Scripting.Dictionary")
    ActiveSheet.Cells(1, 1).Select
    If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.ShowAllData
    End If
    
    With ActiveSheet
        .AutoFilterMode = False
        arrData = .Range("B1:B" & .Cells(.Rows.Count, "A").End(xlUp).Row)
        For Each arrEle In Array("CA *", "CIO*", "CE*", "CIC*", "BPGO*", "BNP*", "SG*", "CM*", "BP*", "CREDIT MUTUEL*", "Crédit-Agricole", "CREDIT*")
            For j = 1 To UBound(arrData)
                If arrData(j, 1) Like arrEle Then dic(arrData(j, 1)) = vbNullString
            Next
        Next
        Range("A1").AutoFilter Field:=2, Criteria1:=dic.Keys, Operator:=xlFilterValues
        Range("A1").AutoFilter Field:=5, Criteria1:="6*"
    End With
End Sub
Sub filtre7()
Attribute filtre7.VB_ProcData.VB_Invoke_Func = "o\n14"
    
    Dim i       As Long
    Dim j       As Long
    Dim dic     As Object
    Dim arrEle  As Variant
    Dim arrData As Variant
    Set dic = CreateObject("Scripting.Dictionary")
    ActiveSheet.Cells(1, 1).Select
    If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.ShowAllData
    End If
    
    With ActiveSheet
        .AutoFilterMode = False
        arrData = .Range("B1:B" & .Cells(.Rows.Count, "B").End(xlUp).Row)
        For Each arrEle In Array("CA *", "CIO*", "CE*", "CIC*", "BPGO*", "BNP*", "SG*", "CM*", "BP*", "CREDIT MUTUEL*", "Crédit-Agricole", "CREDIT*")
            For j = 1 To UBound(arrData)
                If arrData(j, 1) Like arrEle Then dic(arrData(j, 1)) = vbNullString
            Next
        Next
        Range("A1").AutoFilter Field:=2, Criteria1:=dic.Keys, Operator:=xlFilterValues
        Range("A1").AutoFilter Field:=5, Criteria1:="7*"
    End With
End Sub
Sub filtreOD()
Attribute filtreOD.VB_ProcData.VB_Invoke_Func = "p\n14"
    
    Dim i       As Long
    Dim j       As Long
    Dim dic     As Object
    Dim arrEle  As Variant
    Dim arrData As Variant
    Set dic = CreateObject("Scripting.Dictionary")
    ActiveSheet.Cells(1, 1).Select
    If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.ShowAllData
    End If
    
    With ActiveSheet
        .AutoFilterMode = False
        arrData = .Range("B1:B" & .Cells(.Rows.Count, "B").End(xlUp).Row)
        For Each arrEle In Array("OD", "OPERATIONS DIVERSES*")
            For j = 1 To UBound(arrData)
                If arrData(j, 1) Like arrEle Then dic(arrData(j, 1)) = vbNullString
            Next
        Next
        Range("A1").AutoFilter Field:=2, Criteria1:=dic.Keys, Operator:=xlFilterValues
        Range("A1").AutoFilter Field:=5, Criteria1:=Array("411*", "401*", "411000", "401000", "41100", "40100", "4110", "4010", "411", "401", "4110000", "4010000", "411000000", "401000000"), Operator:=xlFilterValues
    End With
End Sub
Sub filtreACHAT()
Attribute filtreACHAT.VB_ProcData.VB_Invoke_Func = "u\n14"
    
    Dim i       As Long
    Dim j       As Long
    Dim dic     As Object
    Dim arrEle  As Variant
    Dim arrData As Variant
    Set dic = CreateObject("Scripting.Dictionary")
    ActiveSheet.Cells(1, 1).Select
    If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.ShowAllData
    End If
    
    With ActiveSheet
        .AutoFilterMode = False
        arrData = .Range("B1:B" & .Cells(.Rows.Count, "B").End(xlUp).Row)
        For Each arrEle In Array("ACHATS")
            For j = 1 To UBound(arrData)
                If arrData(j, 1) Like arrEle Then dic(arrData(j, 1)) = vbNullString
            Next
        Next
        Range("A1").AutoFilter Field:=2, Criteria1:=dic.Keys, Operator:=xlFilterValues
        Range("A1").AutoFilter Field:=5, Criteria1:="7*"
    End With
End Sub
