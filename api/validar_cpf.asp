<%
	Response.ContentType = "application/json"
	Response.CodePage = 65001
	Response.CharSet = "UTF-8"

	Dim cpf
	cpf = Trim(Request.QueryString("cpf"))

	If Len(cpf) <> 11 Or Not IsNumeric(cpf) Then
		Response.Write "{""success"":false,""message"":""CPF inválido ou malformado""}"
		Response.End
	End If

	Dim http, body, resp
	Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
	body = "{""values"": """ & cpf & """, ""format"": false}"
	http.Open "POST", "https://cpfgenerator.org/api/cpf/validator", False
	http.setRequestHeader "Content-Type", "application/json"
	http.Send body

	If http.Status = 200 Then
		resp = http.responseText
	Else
		resp = "{""success"":false,""message"":""Erro na API externa""}"
	End If

	Response.Write resp
%>
