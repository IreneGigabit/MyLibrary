<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="EncodeTester.aspx.cs" Inherits="TestWeb.EncodeTester" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE9"/>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8"/>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
		<input type="text" name="gov_no" size="50" value="<%=gov_no%>"/>
		<input type="submit" name="btn1" value="送出">
    </div>
    </form>
</body>
</html>
