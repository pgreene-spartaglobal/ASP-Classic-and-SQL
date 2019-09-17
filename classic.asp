<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title></title>
</head>
<body>
    <h1>classic.asp</h1>
    <a href="index.html">index.html</a>

    <h2>
    <% Response.Write("classic.asp is connected") %>
    </h2>

    <p>
        <%
            Dim conn
            Set conn = Server.CreateObject("ADODB.Connection")
            'Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=CSharpGame;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False
            conn.ConnectionString = "Provider=SQLNCLI11;Server=(localdb)\MSSQLLocalDB;Trusted_Connection=yes;timeout=30;AttachDbFileName=C:\Users\Philip Greene\CSharpGame.mdf"
            
            conn.Open

            If conn.errors.count = 0 Then
            Response.Write("Connected OK")
            End If
        %>
    </p>

</body>
</html>