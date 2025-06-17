<!--#include file="config.asp"-->
<!--#include file="logger.asp"-->
<%
Function Conectar()
    On Error Resume Next

    If Application("DB_PATH") = "" Then
        Application("DB_PATH") = Server.MapPath("../dados/banco.mdb")
    End If

    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")

    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application("DB_PATH")

    If Err.Number <> 0 Then
        If Application("LOG_ATIVO") = True Then
            Call EscreverLog("Erro de conexão: " & Err.Description)
        End If
        Set Conectar = Nothing
        Err.Clear
    Else
        Set Conectar = conn
    End If
End Function
%>
