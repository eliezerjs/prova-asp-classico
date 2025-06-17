<%
Response.CodePage = 65001
Response.CharSet = "UTF-8"
%>
<!--#include file="../includes/conexao.asp"-->

<%
Dim conn, rs
Set conn = Conectar()
If conn Is Nothing Then
    Response.Write "<p style='color:red'>Erro ao conectar com o banco de dados.</p>"
    Response.End
End If

Set rs = conn.Execute("SELECT id, nome, email FROM usuarios")
%>

<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Usuários</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f2f4f8;
            padding: 40px;
        }

        h2 {
            color: #333;
            margin-bottom: 20px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            background: #fff;
            box-shadow: 0 2px 6px rgba(0,0,0,0.1);
        }

        th, td {
            padding: 12px 16px;
            text-align: left;
            border-bottom: 1px solid #e0e0e0;
        }

        th {
            background-color: #f5f5f5;
        }

        a.button {
            display: inline-block;
            padding: 8px 16px;
            background-color: #007bff;
            color: white;
            text-decoration: none;
            border-radius: 4px;
            font-weight: bold;
        }

        a.button:hover {
            background-color: #0056b3;
        }

        .actions a {
            margin-right: 10px;
            color: #007bff;
            text-decoration: none;
        }

        .actions a:hover {
            text-decoration: underline;
        }

        .novo {
            margin-top: 20px;
        }
		
		.validar-cpf {
			margin-top: 40px;
			padding: 20px;
			background: #fff;
			border-radius: 8px;
			max-width: 400px;
			box-shadow: 0 1px 5px rgba(0,0,0,0.1);
		}
		
		
    </style>
	
<script>
function validarCPF() {
    const cpf = document.getElementById("cpf").value.trim();
    const res = document.getElementById("resultado");

    if (cpf.length !== 11 || isNaN(cpf)) {
        res.style.color = "red";
        res.textContent = "CPF inválido. Digite 11 números.";
        return;
    }

    res.style.color = "#333";
    res.textContent = "Validando...";

    fetch("../api/validar_cpf.asp?cpf=" + cpf)
        .then(r => r.json())
        .then(data => {
            if (data.success && data.data.summary.valid > 0) {
                res.style.color = "green";
                res.textContent = "CPF válido!";
            } else {
                res.style.color = "red";
                res.textContent = "CPF inválido!";
            }
        })
        .catch(err => {
            console.error(err);
            res.style.color = "red";
            res.textContent = "Erro ao validar o CPF.";
        });
}
</script>



</head>
<body>

<div class="validar-cpf">
    <h3>Validação de CPF</h3>
    <input type="text" id="cpf" placeholder="Digite o CPF (somente números)" maxlength="11" style="padding:8px; width:200px;">
    <button onclick="validarCPF()" class="button">Validar CPF</button>
    <p id="resultado" style="margin-top:10px;"></p>
</div>


<h2>Usuários Cadastrados</h2>

<table>
    <thead>
        <tr>
            <th>ID</th>
            <th>Nome</th>
            <th>Email</th>
            <th>Ações</th>
        </tr>
    </thead>
    <tbody>
        <%
        Do Until rs.EOF
            Response.Write "<tr>"
            Response.Write "<td>" & rs("id") & "</td>"
            Response.Write "<td>" & Server.HTMLEncode(rs("nome")) & "</td>"
            Response.Write "<td>" & Server.HTMLEncode(rs("email")) & "</td>"
            Response.Write "<td class='actions'>" & _
                "<a href='form.asp?id=" & rs("id") & "'>Editar</a>" & _
                "<a href='excluir.asp?id=" & rs("id") & "' onclick='return confirm(""Deseja realmente excluir?"")'>Excluir</a>" & _
                "</td>"
            Response.Write "</tr>"
            rs.MoveNext
        Loop
        rs.Close
        conn.Close
        %>
    </tbody>
</table>

<div class="novo">
    <a class="button" href="form.asp">Novo Usuário</a>
</div>

</body>
</html>
