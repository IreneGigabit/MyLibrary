<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="RepeaterTester.aspx.cs" Inherits="TestWeb.ListTester" %>
<%@Import Namespace = "MyLibrary"%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<script runat="server">
	private void Page_Load(System.Object sender, System.EventArgs e) {
		string cnstr="Server=web08;Database=account;User ID=web_usr;Password=web1823";
		branch_Rptr.SetItem(cnstr, "select top 10 seq,cappl_name from dmp where cappl_name like '%&#%' ", true, "22514");
	}
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE9"/>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8"/>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    單位別：<select id="sfx_a_branch" name="sfx_a_branch" size="1" title="單位別">
        <asp:Repeater id="branch_Rptr" runat="server">
            <ItemTemplate>
                <option value="<%#Eval("Value")%>"<%#Eval("Attr")%>><%#Eval("Disp")%></option>
            </ItemTemplate>
        </asp:Repeater>
    </select>
</body>
</html>
