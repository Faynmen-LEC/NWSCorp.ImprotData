<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ImportJob.aspx.cs" Inherits="NWSCorp.ImprotData.ImportJob" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:FileUpload ID="ImportData" runat="server" accept=".xls,.xlsx" />
            <br />
            <asp:RadioButton ID="RadioButton1" runat="server" Text="English version" GroupName="language" Checked="true"/>  
            <asp:RadioButton ID="RadioButton2" runat="server" Text="繁體中文版本" GroupName="language" />
            <asp:RadioButton ID="RadioButton3" runat="server" Text="简体中文版本" GroupName="language" />
            <br />
            <asp:Button ID="NextButton" runat="server" Text="Next" OnClick="NextButton_Click" />
        </div>
        <div>
            <asp:Label ID="Label1" runat="server" Text="Leclerc："><%= Logstring %></asp:Label>
        </div>
    </form>
</body>
</html>
