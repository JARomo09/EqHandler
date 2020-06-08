Attribute VB_Name = "Macro11"
Dim Indx As Integer
Sub main()
    Dim SwAPP As SldWorks.SldWorks
    Dim SwMDL As SldWorks.ModelDoc2
    Dim EqusMgr As SldWorks.EquationMgr
    
    Set SwAPP = Application.SldWorks
    Set SwMDL = SwAPP.ActiveDoc
    Set EqusMgr = SwMDL.GetEquationMgr
    
    For i = 0 To EqusMgr.GetCount
        If InStr(1, CStr(EqusMgr.Equation(i)), "@") > 0 Or InStr(1, CStr(EqusMgr.Equation(i)), "@", vbTextCompare) <> Null Then
        Else
            UserForm1.ListBox1.AddItem (EqusMgr.Equation(i))
        End If
    Next
    UserForm1.ListBox1.Font.Size = 8
    UserForm1.Show
End Sub

Sub FtchIndx()
    Dim SwAPP As SldWorks.SldWorks
    Dim SwMDL As SldWorks.ModelDoc2
    Dim EqusMgr As SldWorks.EquationMgr
    
    Set SwAPP = Application.SldWorks
    Set SwMDL = SwAPP.ActiveDoc
    Set EqusMgr = SwMDL.GetEquationMgr
    
    For i = 0 To EqusMgr.GetCount
        If EqusMgr.Equation(i) = UserForm1.ListBox1.Value Then
            UserForm1.TextBox1.Value = EqusMgr.Value(i)
            Macro11.Indx = i
            Exit For
        Else: End If
    Next
End Sub

Sub WriteValue()
    Dim SwAPP As SldWorks.SldWorks
    Dim SwMDL As SldWorks.ModelDoc2
    Dim EqusMgr As SldWorks.EquationMgr
    Dim vModels As Variant
    Dim x As Integer
    Dim count As Integer
    
    Set SwAPP = Application.SldWorks
    Set SwMDL = SwAPP.ActiveDoc
    Set EqusMgr = SwMDL.GetEquationMgr
    count = SwAPP.GetDocumentCount
    vModels = SwAPP.GetDocuments

    x = InStr(1, EqusMgr.Equation(Macro11.Indx), "=", vbTextCompare)
    EqusMgr.Equation(Macro11.Indx) = Mid(EqusMgr.Equation(Macro11.Indx), 1, x) & UserForm1.TextBox1.Value
    'SwMDL.ForceRebuild3 (True)
    For Index = LBound(vModels) To UBound(vModels)
        Set swModel = vModels(Index)
        swModel.ForceRebuild3 True
    Next Index
    Unload UserForm1
    For Index = LBound(vModels) To UBound(vModels)
        Set swModel = vModels(Index)
        swModel.ForceRebuild3 True
    Next Index
End Sub
