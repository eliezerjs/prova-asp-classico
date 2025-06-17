<%
	Response.CodePage = 65001
	Response.CharSet = "UTF-8"
%>

<!--#include file="../includes/conexao.asp"-->

<%
	Dim id, nome, email, senha, conn, rs
	id = Request.QueryString("id")
	nome = ""
	email = ""
	senha = ""

	If id <> "" And IsNumeric(id) Then
		Set conn = Conectar()
		If Not conn Is Nothing Then
			Set rs = conn.Execute("SELECT nome, email, senha FROM usuarios WHERE id=" & CLng(id))
			If Not rs.EOF Then
				nome = rs("nome")
				email = rs("email")
				senha = rs("senha")
			End If
			rs.Close
			conn.Close
		End If
	End If
%>

<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <%
		Dim tituloAcao
		If id <> "" And IsNumeric(id) Then
			tituloAcao = "Editar"
		Else
			tituloAcao = "Novo"
		End If
		%>

		<title><%=tituloAcao%> Usuário</title>
		<h2><%=tituloAcao%> Usuário</h2>

    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f7f9fb;
            padding: 40px;
        }

        .form-container {
            background-color: #fff;
            max-width: 400px;
            margin: auto;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }

        h2 {
            margin-top: 0;
            margin-bottom: 20px;
            text-align: center;
            color: #333;
        }

        label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
        }

        input[type="text"], input[type="email"], input[type="password"] {
            width: 100%;
            padding: 10px;
            margin-bottom: 15px;
            border: 1px solid #ccc;
            border-radius: 4px;
        }

        input[type="submit"] {
            width: 100%;
            padding: 10px;
            background-color: #007bff;
            color: white;
            font-weight: bold;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }

        input[type="submit"]:hover {
            background-color: #0056b3;
        }

        .back {
            text-align: center;
            margin-top: 15px;
        }

        .back a {
            text-decoration: none;
            color: #007bff;
        }
    </style>
</head>
<body>

<div class="form-container">
    
    <form action="salvar.asp" method="post">
        <input type="hidden" name="id" value="<%=id%>">

        <label>Nome:</label>
        <input type="text" name="nome" value="<%=Server.HTMLEncode(nome)%>" required>

        <label>Email:</label>
        <input type="email" name="email" value="<%=Server.HTMLEncode(email)%>" required>

        <label>Senha:</label>
        <input type="password" name="senha" value="<%=Server.HTMLEncode(senha)%>" required>

        <input type="submit" value="Salvar">
    </form>

    <div class="back">
        <a href="listar.asp">Voltar para a lista</a>
    </div>
</div>

</body>
</html>
