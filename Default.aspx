<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<%@ Register TagPrefix="Web" Namespace="WebChart" Assembly="WebChart" %>
<Web:ChartControl ID="ChartControl1" runat="Server" Height="366px" 
                        Width="500px">
 </Web:ChartControl>    
    </div>
    </form>
</body>
</html>
