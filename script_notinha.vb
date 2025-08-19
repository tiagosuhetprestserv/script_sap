Sub AtualizarSAPComPlanilha()

    Dim SapGuiAuto As Object
    Dim SAPApp As Object
    Dim SAPCon As Object
    Dim session As Object
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long

    ' Define a planilha ativa
    Set ws = ThisWorkbook.Sheets("NotasSAP") ' Altere para o nome correto da aba

    ' Inicializa conexão com o SAP
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set SAPCon = SAPApp.Children(0)
    Set session = SAPCon.Children(0)

    ' Encontra a última linha com dados
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop pelas linhas da planilha
    For i = 2 To ultimaLinha
        Dim numNota As String
        Dim tituloNota As String
        Dim mensagemNota As String
        Dim andamento As String

        numNota = ws.Cells(i, 1).Value
        tituloNota = ws.Cells(i, 2).Value
        mensagemNota = ws.Cells(i, 3).Value
        andamento = ws.Cells(i, 4).Value

        ' Interações com a tela do SAP
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = numNota
        session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").caretPosition = Len(numNota)
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[1]/usr/txtDOCUMENT_TITLE").Text = tituloNota & " - " & mensagemNota
        session.findById("wnd[1]/usr/txtDOCUMENT_TITLE").caretPosition = Len(tituloNota & " - " & mensagemNota)
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[0]/tbar[0]/btn[11]").press
    Next i

End Sub
