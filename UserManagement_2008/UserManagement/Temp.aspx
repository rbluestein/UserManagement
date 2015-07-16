<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Temp.aspx.vb" Inherits="Temp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:Button ID="Button1" runat="server" Text="Button" />
        
        <input type="button" value="Submitty" onclick="fnSubmit()" />
    
    </div>
    </form>
    <script type="text/javascript">
        function fnSubmit() {
            form1.submit()
            //document.getElementsByTagName("form")[0].submit()
        }
    </script>
</body>
</html>
