<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="css/bootstrap-3.3.6/css/bootstrap.css"
        rel="stylesheet" />
</head>
<body>
    <form id="form1" runat="server">
    <div class="container">
        <div class="col-md-12">  
            <b>Please select Excel file </b>
        </div>
        <div class="col-md-6">
            <asp:FileUpload ID="FileUploadExcel" runat="server" class="form-control" />
        </div>
        <div class="col-md-3">
            <asp:Button ID="btnImport" runat="server" Text="Import" class="btn btn-primary" />
        </div>
        <br />
        <div class="col-md-12">
           <b> <asp:Label ID="lblmsg" runat="server" Text=""></asp:Label></b>
        </div>
        <br />
        <br />
        <br />
        <div class="col-md-12">
            <asp:GridView ID="gvExcelData" runat="server" class="table table-striped table-hover" ></asp:GridView>
        </div>
    </div>
    </form>
</body>
</html>
