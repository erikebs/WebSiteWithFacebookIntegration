<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebSiteWithFacebookIntegration.Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>

    <!-- Bootstrap core CSS -->
    <link href="~/Content/bootstrap.min.css" rel="stylesheet" />
    <script src="scripts/jquery-1.9.1.min.js"></script>
    <script src="scripts/angular.min.js"></script>
    <script src="scripts/angular-animate.min.js"></script>
    <script src="scripts/angular-aria.min.js"></script>
    <script src="scripts/angular-messages.min.js"></script>


    <!-- Include Required Prerequisites -->
    <script type="text/javascript" src="//cdn.jsdelivr.net/jquery/1/jquery.min.js"></script>
    <script type="text/javascript" src="//cdn.jsdelivr.net/momentjs/latest/moment.min.js"></script>
    <link rel="stylesheet" type="text/css" href="//cdn.jsdelivr.net/bootstrap/latest/css/bootstrap.css" />

    <!-- Include Date Range Picker -->
    <script type="text/javascript" src="//cdn.jsdelivr.net/bootstrap.daterangepicker/2/daterangepicker.js"></script>
    <link rel="stylesheet" type="text/css" href="//cdn.jsdelivr.net/bootstrap.daterangepicker/2/daterangepicker.css" />


    <script type="text/javascript">
        $(function () {
            $('input[name="daterange"]').daterangepicker();
        });



    </script>

</head>

<body>
    <br />
    <form runat="server" class="container">

        <asp:ScriptManager EnablePartialRendering="true" ID="ScriptManager1" runat="server"></asp:ScriptManager>

        <div class="form-group">
            <div class="row">
                <div class="col-sm-3">
                    <asp:Button ID="btnLogin" CssClass="btn btn-primary btn-sm active" runat="server" Text="Log in with Facebook" OnClick="btnLogin_Click" />
                    <asp:Label ID="lblUsrName" runat="server" Text=""></asp:Label>
                </div>
            </div>

            <br />

            <div class="row">
                <div class="form-inline">
                    <div class="col-sm-8">
                        <label for="txtDtIni">From - To:</label>
                        <input type="text" class="form-control" runat="server" id="daterange" name="daterange" value="01/01/2015 - 01/31/2015" />
                        <asp:Button ID="btnSearch" CssClass="btn btn-primary btn-sm active" runat="server" Text="Search" OnClick="btnSearch_Click" />
                    </div>
                </div>
            </div>

            <br />

            <div class="row">

                <div class="col-sm-4">
                    <span class="control-label">Your Friends</span><br />
                    <asp:ListBox ID="lstFriends" CssClass="form-control" runat="server" Height="200px"></asp:ListBox>
                    <span>Friend: <span class="badge">
                        <asp:Label ID="lblQtdFriends" CssClass="control-label" runat="server" Text=""></asp:Label></span></span>
                </div>

                <div class="col-sm-8">
                    <span class="control-label">Your Groups</span><br />
                    <asp:ListBox ID="lstGroups" runat="server" CssClass="form-control" Height="200px" AutoPostBack="True" OnSelectedIndexChanged="lstGroups_SelectedIndexChanged"></asp:ListBox>
                    <span>Group: <span class="badge">
                        <asp:Label ID="lblQtdGroup" CssClass="control-label" runat="server" Text=""></asp:Label></span></span>
                </div>
            </div>
            <br />
            <asp:UpdatePanel ID="UpdatePanel1" runat="server"
                UpdateMode="Conditional">
                <ContentTemplate>
                    <div class="row">
                        <div class="col-sm-12">
                            <span class="control-label">Group's Posts</span><br />
                            <asp:ListBox ID="lstPosts" runat="server" CssClass="form-control" Height="200px" AutoPostBack="True" OnSelectedIndexChanged="lstPosts_SelectedIndexChanged"></asp:ListBox>
                            <span>Group's Post: <span class="badge">
                                <asp:Label ID="lblPostsGroup" runat="server" Text=""></asp:Label></span></span>
                        </div>
                    </div>
                    <br />
                    <div class="row">
                        <div class="col-sm-3">
                            <span class="control-label">Who Likes?</span><br />
                            <asp:ListBox ID="lstLikes" runat="server" CssClass="form-control" Height="200px"></asp:ListBox>
                            <span>Like: <span class="badge">
                                <asp:Label ID="lblLike" runat="server" Text=""></asp:Label></span>
                        </div>
                        <div class="col-sm-9">
                            <span class="control-label">Who and what commented?</span><br />
                            <asp:ListBox ID="lstComments" CssClass="form-control" runat="server" Height="200px"></asp:ListBox>
                            <span>Comment: <span class="badge">
                                <asp:Label ID="lblComments" runat="server" Text=""></asp:Label></span></span>
                        </div>
                    </div>
                </ContentTemplate>
                <Triggers>
                    <asp:AsyncPostBackTrigger ControlID="lstPosts" EventName="SelectedIndexChanged" />
                </Triggers>
            </asp:UpdatePanel>

            <div class="row">
                <div class="col-sm-12 text-right">
                    <asp:Button ID="btnExtract" CssClass="btn btn-success btn-sm active disabled" Enabled="false" runat="server" Text="Export to Excel" OnClick="btnExtract_Click" />
                </div>
            </div>

        </div>

    </form>


</body>
</html>


