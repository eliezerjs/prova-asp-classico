<%
	Response.CodePage = 65001
	Response.CharSet = "UTF-8"
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

If Not EmailValido(email) Then
    Response.Write "<p style='color:red'>Email inválido.</p>"
    Response.End
End If

If senha = "" Then
    Response.Write "<p style='color:red'>Senha não pode ser vazia.</p>"
    Response.End
End If

senhaCriptografada = GerarMD5(senha)

If senhaCriptografada = "" Then
    Call EscreverLog("Erro: GerarMD5 retornou vazio. Texto original: " & senha)
    Response.Write "<p style='color:red'>Erro ao criptografar a senha. Valor gerado está vazio.</p>"
    Response.End
End If

Dim conn, cmd, rs, sql
Set conn = Conectar()
If conn Is Nothing Then
    Response.Write "<p style='color:red'>Erro na conexão com o banco.</p>"
    Response.End
End If

Dim idLong, isValidId
isValidId = False

If IsNumeric(id) Then
    idLong = CLng(id)
    isValidId = True
End If

sql = "SELECT COUNT(*) AS total FROM usuarios WHERE email = ?"
If isValidId Then
    sql = sql & " AND id <> ?"
End If

Set cmd = Server.CreateObject("ADODB.Command")
With cmd
    .ActiveConnection = conn
    .CommandText = sql
    .CommandType = 1
    .Parameters.Append .CreateParameter("email", 200, 1, 100, email)

    If isValidId Then
        .Parameters.Append .CreateParameter("id", 3, 1, , idLong)
    End If

    Set rs = .Execute()
End With

If rs("total") > 0 Then
    Response.Write "<p style='color:red'>E-mail já cadastrado.</p>"
    conn.Close
    Response.End
End If

Set cmd = Nothing
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn

If Not isValidId Then
    cmd.CommandText = "INSERT INTO usuarios (nome, email, senha) VALUES (?, ?, ?)"
    cmd.Parameters.Append cmd.CreateParameter("nome", 200, 1, 100, nome)
    cmd.Parameters.Append cmd.CreateParameter("email", 200, 1, 100, email)
    cmd.Parameters.Append cmd.CreateParameter("senha", 200, 1, 100, senhaCriptografada)
Else
    cmd.CommandText = "UPDATE usuarios SET nome=?, email=?, senha=? WHERE id=?"
    cmd.Parameters.Append cmd.CreateParameter("nome", 200, 1, 100, nome)
    cmd.Parameters.Append cmd.CreateParameter("email", 200, 1, 100, email)
    cmd.Parameters.Append cmd.CreateParameter("senha", 200, 1, 100, senhaCriptografada)
    cmd.Parameters.Append cmd.CreateParameter("id", 3, 1, , idLong)
End If

Dim sqlDebug
sqlDebug = cmd.CommandText

If cmd.Parameters.Count > 0 Then
    Dim i
    For i = 0 To cmd.Parameters.Count - 1
        sqlDebug = Replace(sqlDebug, "?", "'" & cmd.Parameters(i).Value & "'", 1, 1)
    Next
End If

Call LogErro("Comando executado: " & sqlDebug)
Response.Write "<pre style='color:blue'>" & sqlDebug & "</pre><hr>"

cmd.Execute , , 129

If Err.Number <> 0 Then
    Call LogErro("Erro ao executar INSERT/UPDATE: " & Err.Description)
    Response.Write "<p style='color:red'>Erro ao salvar o usuário: " & Err.Description & "</p>"
    Err.Clear
    conn.Close
    Response.End
End If

conn.Close
Response.Redirect "listar.asp"
%>
