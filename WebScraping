Sub WebScraping()
    Dim navegador As Object
    Dim elementoLink As Object
    Dim elementoSelect, frame As Object
    Dim opcao As Object
    Dim rows As Object
    Dim row As Object
    Dim dado As String
    Dim ws As Worksheet
    Dim tabela As QueryTable
    
    Set navegador = CreateObject("InternetExplorer.Application")
    navegador.Navigate "http://estatisticas.cetip.com.br/astec/series_v05/paginas/lum_web_v05_series_introducao.asp?str_Modulo=Ativo&int_Idioma=1&int_Titulo=6&int_NivelBD=2"
    
    navegador.Visible = True
    
    Do While navegador.Busy Or navegador.readyState <> 4
        Application.Wait Now + TimeValue("00:00:01")
    Loop
    Set elementoLink = navegador.document.querySelector("a[onclick='MenuAtivo(); return false;']")
    If Not elementoLink Is Nothing Then
        elementoLink.Click
    Else
        MsgBox "Erro ao encontrar o link."
        navegador.Quit
        Exit Sub
    End If
    Do While navegador.Busy Or navegador.readyState <> 4
        Application.Wait Now + TimeValue("00:00:01")
    Loop
 
    Set elementoSelect = navegador.document.getElementsByTagName("iframe")(0).contentDocument.getElementsByTagName("frame")(0).contentDocument.getElementsByTagName("select")(0)
    For Each opcao In elementoSelect.Options
        If opcao.Value = "CBIO" Then
            opcao.Selected = True
            Exit For
        Else
            opcao.Selected = False
        End If
    Next opcao
    elementoSelect.FireEvent ("onChange")
    
    Do While navegador.Busy Or navegador.readyState <> 4
        Application.Wait Now + TimeValue("00:00:01")
    Loop
    
    Set elementoSelect = navegador.document.getElementsByTagName("iframe")(0).contentDocument.getElementsByTagName("frame")(1).contentDocument.getElementsByTagName("select")(0)
    For Each opcao In elementoSelect.Options
        If opcao.Text = "Negociações Definitivas" Then
            opcao.Selected = True
            Exit For
        Else
            opcao.Selected = False
        End If
    Next opcao
    elementoSelect.FireEvent ("onChange")

    Do While navegador.Busy Or navegador.readyState <> 4
        Application.Wait Now + TimeValue("00:00:01")
    Loop
    
    Set elementoDiaInicial = navegador.document.getElementsByTagName("iframe")(0).contentDocument.getElementsByTagName("frame")(2).contentDocument.getElementsByName("DT_DIA_DE")(0)
    elementoDiaInicial.Value = Day(Now() - 4)
    Set elementoMesInicial = navegador.document.getElementsByTagName("iframe")(0).contentDocument.getElementsByTagName("frame")(2).contentDocument.getElementsByName("DT_MES_DE")(0)
    elementoMesInicial.Value = Month(Now() - 4)
    Set elementoAnoInicial = navegador.document.getElementsByTagName("iframe")(0).contentDocument.getElementsByTagName("frame")(2).contentDocument.getElementsByName("DT_ANO_DE")(0)
    elementoAnoInicial.Value = Year(Now() - 4)
    Set btnPesquisar = navegador.document.getElementsByTagName("iframe")(0).contentDocument.getElementsByTagName("frame")(2).contentDocument.getElementsByTagName("a")(1)
    btnPesquisar.Click
    
    Do While navegador.Busy Or navegador.readyState <> 4
        Application.Wait Now + TimeValue("00:00:01")
    Loop
    

    Dim htmlString As String
    Set rows = navegador.document.getElementsByClassName("ConsultaDados_R_02")
    
    Dim linhaAtual As Long
    linhaAtual = 1
    
    Set ws = Worksheets("Cetip")
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.rows.Count, linhaAtual).End(xlUp).row
    
    For Each row In rows

        ws.Cells(ultimaLinha + 1, linhaAtual).Value = row.innerText
        linhaAtual = linhaAtual + 1
    Next row
    
    Do While navegador.Busy Or navegador.readyState <> 4
        Application.Wait Now + TimeValue("00:00:01")
    Loop

    navegador.Quit
    
End Sub
