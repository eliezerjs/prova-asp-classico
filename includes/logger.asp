<%
Sub LogErro(mensagem)
    Dim fso, arquivo, caminhoLog, dataHora
    dataHora = Now()
    caminhoLog = Server.MapPath("../logs/erros.txt") 

    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(Server.MapPath("../logs")) Then
        fso.CreateFolder(Server.MapPath("../logs"))
    End If

    Set arquivo = fso.OpenTextFile(caminhoLog, 8, True) 
    arquivo.WriteLine("[" & dataHora & "] " & mensagem)
    arquivo.Close

    Set arquivo = Nothing
    Set fso = Nothing
End Sub
%>
