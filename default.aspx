<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="SSAS2012_MyBI.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
<script type="text/javascript">
    function ExcelLoad(instance,catalog,cube) {
        
      window.open("http://<servername>/ExcelLoad.aspx?instance=" + instance + "&catalog=" + catalog + "&cube=" + cube);
      return false;
    }

   
</script>

</head>
<body>
    <form id="form1" runat="server">
    <div>     
                          <asp:Label ID="Label1" runat="server" Text="Enter Source"></asp:Label>
                        <br />
                        <asp:TextBox ID="TextBoxSource" runat="server" Width="307px"></asp:TextBox>
                        &nbsp;<asp:Button ID="ButtonSubmit" runat="server" OnClick="ButtonSubmit_Click" Text="Submit" /> 

                   

                        <p>
                        <asp:GridView ID="GridViewResults" runat="server" OnRowDataBound="gvResults_RowDataBound">
                        </asp:GridView>
                  </p>
                <br />

                            </div>
    </form>

</body>
</html>
