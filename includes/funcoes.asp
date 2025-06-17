<%

Function EmailValido(email)
    EmailValido = (InStr(email, "@") > 0 And InStr(email, ".") > 0)
End Function

Function EscreverLog(msg)
    If Application("LOG_ATIVO") Then
        On Error Resume Next
        Dim fso, arquivo
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        Set arquivo = fso.OpenTextFile(Application("LOG_CAMINHO"), 8, True)
        arquivo.WriteLine Now() & " - " & msg
        arquivo.Close
        Set arquivo = Nothing
        Set fso = Nothing
    End If
End Function


Function GerarMD5(texto)
    On Error Resume Next

    Dim md5, resultado
    resultado = ""

    Set md5 = CreateObject("CAPICOM.HashedData")
    If Not md5 Is Nothing Then
        md5.Algorithm = 3 
        md5.Hash texto
        resultado = md5.Value
        Set md5 = Nothing
    End If

    If resultado = "" Then
        resultado = texto
        Call EscreverLog("GerarMD5: fallback ativado, usando valor puro.")
    End If

    GerarMD5 = resultado
End Function

Function SanitizarEntrada(valor)
    valor = Replace(valor, "'", "''")
    valor = Replace(valor, "<", "")
    valor = Replace(valor, ">", "")
    SanitizarEntrada = Trim(valor)
End Function

Sub TratarErro(msg)
    EscreverLog "Erro: " & msg & " | " & Err.Description
    Response.Write "<p>Ocorreu um erro ao processar a solicitação. Tente novamente mais tarde.</p>"
End Sub
%>
