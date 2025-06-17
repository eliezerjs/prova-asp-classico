<!--#include file="../includes/conexao.asp"-->
<!--#include file="../includes/logger.asp"-->
<%
	On Error Resume Next

	If Err.Number <> 0 Then
		Call LogErro("Erro em excluir.asp: " & Err.Description & " | Linha: " & Erl)
		Response.Write("Ocorreu um erro. Verifique os logs.")
		Err.Clear
	End If

	Dim id
	id = Request.QueryString("id")
	If id <> "" Then
		Set conn = Conectar()
		conn.Execute("DELETE FROM usuarios WHERE id=" & id)
		conn.Close
	End If
	Response.Redirect "listar.asp"
%>
