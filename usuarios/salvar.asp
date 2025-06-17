<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Response.ContentType = "text/html; charset=UTF-8"
%>
<!--#include file="../includes/conexao.asp"-->
<!--#include file="../includes/funcoes.asp"-->
<!--#include file="../includes/logger.asp"-->

<%
On Error Resume Next

Dim id, nome, email, senha, senhaCriptografada
id    = Trim(Request.Form("id"))
nome  = Trim(Request.Form("nome"))
email = Trim(Request.Form("email"))
senha = Trim(Request.Form("senha"))

Sub MostrarErro(msg)
    Response.Write "<!DOCTYPE html><html><head>" & _
    "<meta charset='UTF-8'>" & _
    "<style>" & _
    "body{font-family:Arial;background:#f9f9f9;padding:40px;text-align:center;}" & _
    ".erro{background:#ffe6e6;border:1px solid #ffcccc;color:#b30000;padding:20px;" & _
    "border-radius:8px;max-width:400px;margin:40px auto;font-size:16px}" & _
    ".voltar{margin-top:15px}" & _
    ".voltar a{color:#007bff;text-decoration:none;font-weight:bold;}" & _
    "</style></head><body>" & _
    "<div class='erro'>" & msg & "</div>" & _
    "<div class='voltar'><a href='javascript:history.back()'>Voltar</a></div>" & _
    "</body></html>"
    Response.End
End Sub

If Not EmailValido(email) Then MostrarErro "Email inválido."
If senha = "" Then MostrarErro "Senha não pode ser vazia."

senhaCriptografada = GerarMD5(senha)
If senhaCriptografada = "" Then
    Call EscreverLog("GerarMD5 retornou vazio. Usando senha em texto.")
    senhaCriptografada = senha
End If

Dim conn, cmd, rs, sql
Set conn = Conectar()
If conn Is Nothing Then MostrarErro "Erro ao conectar com o banco de dados."

' Verifica e-mail duplicado
sql = "SELECT COUNT(*) AS total FROM usuarios WHERE email = ?"
If IsNumeric(id) And id <> "" Then sql = sql & " AND id <> ?"

Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn
cmd.CommandText = sql
cmd.CommandType = 1
cmd.Parameters.Append cmd.CreateParameter("email", 200, 1, 100, email)

If IsNumeric(id) And id <> "" Then
    cmd.Parameters.Append cmd.CreateParameter("id", 3, 1, , CLng(id))
End If

Set rs = cmd.Execute()
If rs("total") > 0 Then MostrarErro "E-mail já cadastrado."

rs.Close
Set cmd = Nothing

' Gravação
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn

If id = "" Then
    cmd.CommandText = "INSERT INTO usuarios (nome, email, senha) VALUES (?, ?, ?)"
    cmd.Parameters.Append cmd.CreateParameter("nome", 200, 1, 100, nome)
    cmd.Parameters.Append cmd.CreateParameter("email", 200, 1, 100, email)
    cmd.Parameters.Append cmd.CreateParameter("senha", 200, 1, 100, senhaCriptografada)
Else
    cmd.CommandText = "UPDATE usuarios SET nome = ?, email = ?, senha = ? WHERE id = ?"
    cmd.Parameters.Append cmd.CreateParameter("nome", 200, 1, 100, nome)
    cmd.Parameters.Append cmd.CreateParameter("email", 200, 1, 100, email)
    cmd.Parameters.Append cmd.CreateParameter("senha", 200, 1, 100, senhaCriptografada)
    cmd.Parameters.Append cmd.CreateParameter("id", 3, 1, , CLng(id))
End If

cmd.Execute , , 129

If Err.Number <> 0 Then
    Call LogErro("Erro ao executar INSERT/UPDATE: " & Err.Description)
    MostrarErro "Erro ao salvar o usuário: " & Err.Description
End If

conn.Close
Response.Redirect "listar.asp"
%>
