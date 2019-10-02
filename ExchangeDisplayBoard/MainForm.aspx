<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MainForm.aspx.cs" Inherits="ExchangeDisplayBoard.MainForm" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Display Board</title>
  <link rel="stylesheet" href="main.css"/>
   <!--<meta http-equiv="refresh" content="600" /> -->
   
<script src="Scripts/jquery-3.3.1.js"></script>
<script type="text/javascript">
    //$(document).ready(function () {
      //  setInterval(AutoScrollEdge(),10000);
    //});
    
    function AutoScroll() {
        document.getElementById('divGV').scroll();
        document.getElementById('divGV').animate({
            scrollTop: 1000
        }, 2000);
        $("#divGV").animate({
          scrollTop: 0
        }, 5000).delay(4000);
        return false;
    }

    function AutoScrollEdge()
    {
        var objDiv = document.getElementById("divGV");

        $('#divGV').scroll();
        $("#divGV").animate({
          scrollTop: objDiv.scrollHeight
        }, 9000);
        $("#divGV").animate({
          scrollTop: 0
        }, 1000);
        return false;
    }

</script>
</head>   

<body>   
    <form id="form1" runat="server">
        <div class="Header">
            <table style="width:100%;">
                <tr>
                    <td>
                    <asp:Image ID="Image1" runat="server" ImageUrl="logo.jpg" style="padding : 5pt" Visible="False"/>
                        <asp:Label ID="lblRoomName" runat="server" CssClass="TextHeader" Text="Room Name"></asp:Label>
                    </td>
                    <td aria-required="True" style="text-align:right; ">
                        <asp:Label ID="lblDateHeader" runat="server" Text="Monday, May 29th 2019" CssClass="TextHeader"></asp:Label>
                    </td>
                </tr>
            </table>
        </div>
        <div id="sPLine" class="SPLine" runat="server"> 
        </div>
        <div class="divTitle">
            <table style="border-style: none; width: 80%; margin:0 auto">
                <tr>
                    <td style="width:10%;">
                        &nbsp;</td>
                    <td class="gvHeader" style="width:20%;">Time</td>
                    <td class="gvHeader" style="width:30%;">Purpose</td>
                    <td class="gvHeader" style="width:30%;">Resource</td>
                    <td class="gvHeader" style="width:10%;"></td>
                </tr>
            </table>
        </div>            
        <div id="divGV" class="gvDiv" >
            <asp:GridView ID="gvApps" runat="server" BorderStyle="None" AutoGenerateColumns="false" Width="80%" HorizontalAlign="Center" OnDataBound="gvApps_DataBound" 
                OnRowDataBound="gvApps_RowDataBound" ShowHeader="False">
                <Columns>                    
                    <asp:BoundField ItemStyle-Width="10%" />
                    <asp:BoundField HeaderText="Time" DataField="dateTime" ItemStyle-Width="20%"  HeaderStyle-CssClass="gvHeader" ItemStyle-CssClass="gvFirstRowItem"/>
                    <asp:BoundField HeaderText="Purpose" DataField="description" ItemStyle-Width= "30%" HeaderStyle-CssClass="gvHeader" ItemStyle-CssClass="gvMiddleRowItem" />
                    <asp:BoundField HeaderText="Resource" DataField="resource" ItemStyle-Width="30%" HeaderStyle-CssClass="gvHeader" ItemStyle-CssClass="gvLastRowItem" />
                    <asp:BoundField ItemStyle-Width="10%" />
                </Columns>
            </asp:GridView>
        </div>
        <asp:Timer ID="Timer1" runat="server" OnTick="Timer1_Tick" Interval="20000">
        </asp:Timer>
        <asp:ScriptManager ID="ScriptManager1" runat="server" EnableCdn="true">
        </asp:ScriptManager>
    </form>
</body>    
</html>
