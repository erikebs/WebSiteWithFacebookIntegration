using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Facebook;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using WebSiteWithFacebookIntegration.Models;

namespace WebSiteWithFacebookIntegration
{
    public partial class Default : System.Web.UI.Page
    {

        string[] headerColumns = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ" };
        Dictionary<string, List<string>> likesConsolidate;
        List<string> likesAll = new List<string>();

        protected void Page_Load(object sender, EventArgs e)
        {

            if (!Page.IsPostBack)
            {

                this.daterange.Value = $"{DateTime.Today.AddDays(-8).ToString("MM/dd/yyyy")} - {DateTime.Today.ToString("MM/dd/yyyy")}";

                if (Request.QueryString["code"] != null)
                {
                    string accessCode = Request.QueryString["code"].ToString();

                    var fb = new FacebookClient();

                    var appId = ConfigurationManager.AppSettings["appId"].ToString();
                    var appSecret = ConfigurationManager.AppSettings["appSecret"].ToString();

                    if (Request.Url.Authority.IndexOf("localhost") > -1)
                    {
                        appId = ConfigurationManager.AppSettings["localhost_appId"].ToString();
                        appSecret = ConfigurationManager.AppSettings["localhost_appSecret"].ToString();
                    }

                    // throws OAuthException 
                    dynamic result = fb.Post("oauth/access_token", new
                    {
                        client_id = appId,
                        client_secret = appSecret,
                        redirect_uri = $"http://{Request.Url.Authority}/Default.aspx",
                        code = accessCode
                    });

                    var accessToken = result.access_token;
                    var expires = result.expires;

                    // Store the access token in the session
                    Session["AccessToken"] = accessToken;

                    // update the facebook client with the access token 
                    fb.AccessToken = accessToken;

                    var client = new FacebookClient(Session["AccessToken"].ToString());
                    
                    // Calling Graph API for user info
                    dynamic me = fb.Get("me?fields=friends,name,email,first_name");


                    // You can store it in the database
                    string id = me.id;
                    string name = me.name;
                    string email = me.email;
                    string firstName = me.first_name;

                    // set logged user name in label name and email or firstName
                    lblUsrName.Text = name + "(" + (!String.IsNullOrEmpty(email) ? email : firstName) + ")";

                    dynamic fbresult = client.Get($"/{id}/friends?limit=5000&offset=0");

                    var data = fbresult["data"].ToString();
                    List<FacebookUser> friends = JsonConvert.DeserializeObject<List<FacebookUser>>(data);

                    foreach (var friend in friends.OrderBy(q => q.name).ToList())
                        lstFriends.Items.Add(friend.name);

                    lblQtdFriends.Text = lstFriends.Items.Count.ToString();

                    dynamic groups = fb.Get("/me/groups");

                    foreach (var g in groups.data)
                        lstGroups.Items.Add(new ListItem(g.name, g.id));

                    lblQtdGroup.Text = lstGroups.Items.Count.ToString();
                    btnLogin.Text = "Log Out";

                    FormsAuthentication.SetAuthCookie(email, false);
                }
            }
        }

        protected void btnLogin_Click(object sender, EventArgs e)
        {
            if (btnLogin.Text == "Log Out")
                logout();
            else
            {
                var fb = new FacebookClient();

                var appId = ConfigurationManager.AppSettings["appId"].ToString();

                if (Request.Url.Authority.IndexOf("localhost") > -1)
                {
                    appId = ConfigurationManager.AppSettings["localhost_appId"].ToString();
                }

                var loginUrl = fb.GetLoginUrl(new
                {
                    client_id = appId,
                    redirect_uri = $"http://{Request.Url.Authority}/Default.aspx",
                    response_type = "code",
                    scope = "user_friends,user_managed_groups,read_custom_friendlists,publish_actions,user_posts"
                });
                Response.Redirect(loginUrl.AbsoluteUri);
            }
        }

        private void logout()
        {
            var fb = new FacebookClient();

            var logoutUrl = fb.GetLogoutUrl(new
            {
                access_token = Session["AccessToken"],
                next = $"http://{Request.Url.Authority}/Default.aspx"
            });

            // User Logged out, remove access token from session
            Session.Remove("AccessToken");

            Response.Redirect(logoutUrl.AbsoluteUri);
        }

        protected void lstGroups_SelectedIndexChanged(object sender, EventArgs e)
        {
            var idGroup = lstGroups.SelectedValue.ToString();
            loadGroupFeeds(idGroup);
        }


        private void loadGroupFeeds(string idGroup)
        {

            if (String.IsNullOrEmpty(idGroup))
                return;
            
            lstPosts.Items.Clear();
            lstLikes.Items.Clear();
            lstComments.Items.Clear();

            var fb = new FacebookClient();
            fb.AccessToken = Session["AccessToken"].ToString();

            IFormatProvider enUsDateFormat = new CultureInfo("en-US").DateTimeFormat;

            var dtBegin = this.daterange.Value.Substring(0, 10);
            var dtEnd = this.daterange.Value.Substring(13, 10);

            DateTime inicio = DateTime.Parse(dtBegin, enUsDateFormat);
            DateTime termino = DateTime.Parse(dtEnd, enUsDateFormat);

            dynamic feed = fb.Get($"/v2.3/{idGroup}/feed?fields=created_time,message");
            
            bool hasPage = true;
            List<dynamic> filteredPosts = new List<dynamic>();

            do
            {
                var posts = ((IEnumerable)feed.data).Cast<dynamic>()
                            .Where(p => p.message != null);

                foreach (var item in posts.OrderByDescending(i => i.created_time))
                {
                    DateTime dtCreation = DateTime.ParseExact(item.created_time, "yyyy-MM-ddTHH:mm:ss+ffff", enUsDateFormat);

                    //Filter until final date
                    if (dtCreation.Date >= inicio.Date && dtCreation.Date <= termino.Date)
                        filteredPosts.Add(item);
                }

                if (feed?.paging?.next != null)
                    feed = fb.Get(feed.paging.next);
                else
                    hasPage = false;

            } while (hasPage);

            foreach (var item in filteredPosts.OrderBy(i => i.created_time))
            {
                lstPosts.Items.Add(new ListItem($"[{ item.created_time }] {item.message}", item.id));
            }

            lblPostsGroup.Text = lstPosts.Items.Count.ToString();

            btnExtract.Enabled = true;
            btnExtract.CssClass = "btn btn-success btn-sm active";

        }

        protected void lstPosts_SelectedIndexChanged(object sender, EventArgs e)
        {
            var idPost = lstPosts.SelectedValue.ToString();

            var fb = new FacebookClient();
            fb.AccessToken = Session["AccessToken"].ToString();

            dynamic likes = fb.Get($"{idPost}/reactions?offset=0&limit=5000");

            lstLikes.Items.Clear();


            if (likes != null && likes.data != null)
            {
                lblLike.Text = likes.data.Count.ToString();
                foreach (var item in likes.data)
                {
                    lstLikes.Items.Add(item.name);
                }
            }
            else
                lblLike.Text = "No likes";

            dynamic comments = fb.Get($"{idPost}/comments?fields=created_time,message,from&offset=0&limit=5000");
            lstComments.Items.Clear();

            if (comments != null && comments.data != null)
            {
                lblComments.Text = comments.data.Count.ToString();

                foreach (var item in comments.data)
                {
                    lstComments.Items.Add($"({item.from.name}) - {item.message}");
                }
            }
            else
                lblComments.Text = "No comments";
        }

        protected void btnExtract_Click(object sender, EventArgs e)
        {
            var fileName = Server.MapPath("~/file.xlsx");

            CreateExcelDoc(fileName);
            CreateExcelDocLikes(fileName);

            new DownloadFile(fileName).ProcessRequest(HttpContext.Current);
        }

        public void CreateExcelDoc(string fileName)
        {
            if (File.Exists(fileName))
                File.Delete(fileName);

            likesConsolidate = new Dictionary<string, List<string>>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = GenerateStyleSheet();
                stylesPart.Stylesheet.Save();

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Analytic" };

                sheets.Append(sheet);

                workbookPart.Workbook.Save();

                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                // Constructing header
                Row row = new Row();
                row.RowIndex = 1;
                Cell cell = ConstructCell("Post", CellValues.String);
                cell.StyleIndex = 1;
                row.Append(cell);

                foreach (ListItem item in lstPosts.Items)
                    row.Append(ConstructCell(item.Text.PadRight(200, ' ').Substring(28), CellValues.String));

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);

                // Constructing body
                row = new Row();
                row.RowIndex = 2;
                cell = ConstructCell("Date", CellValues.String);
                cell.StyleIndex = 1;
                row.Append(cell);

                foreach (ListItem item in lstPosts.Items)
                    row.Append(ConstructCell(item.Text.PadLeft(30, ' ').Substring(1, 19), CellValues.String));

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);

                var fb = new FacebookClient();
                fb.AccessToken = Session["AccessToken"].ToString();

                row = new Row();
                row.RowIndex = 3;
                cell = ConstructCell("Likes", CellValues.String);
                cell.StyleIndex = 1;
                row.Append(cell);

                foreach (ListItem item in lstPosts.Items)
                {
                    List<string> likesUsr = new List<string>();

                    dynamic likes = fb.Get($"{item.Value}/reactions?offset=0&limit=5000");
                    row.Append(ConstructCell(likes.data.Count.ToString(), CellValues.String));


                    var repeated = ((IEnumerable)likes.data).Cast<dynamic>()
                                  .Select(p => (string)p.name)
                                  .GroupBy(x => x)
                                  .Where(g => g.Count() > 1)
                                  .Select(y => y.Key)
                                  .ToList();

                    likesUsr.AddRange(repeated);
                    likesAll.AddRange(((IEnumerable)likes.data).Cast<dynamic>().Select(p => (string)p.name));

                    likesConsolidate.Add(item.Value, likesUsr);
                }

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);

                row = new Row();
                row.RowIndex = 4;
                cell = ConstructCell("Comments", CellValues.String);
                cell.StyleIndex = 1;
                row.Append(cell);

                foreach (ListItem item in lstPosts.Items)
                {
                    dynamic comments = fb.Get($"{item.Value}/comments?fields=created_time,message,from&offset=0&limit=5000");
                    row.Append(ConstructCell(comments.data.Count.ToString(), CellValues.String));
                }

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);

                //Blank
                row = new Row();
                row.RowIndex = 5;
                row.Append(ConstructCell("", CellValues.String));

                row = new Row();
                row.RowIndex = 6;
                cell = ConstructCell("Who repeated Like?", CellValues.String);
                cell.StyleIndex = 1;
                row.Append(cell);

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);

                int indexColumn = 1;
                foreach (ListItem item in lstPosts.Items)
                {
                    uint indexRow = 6;
                    //dynamic likes = fb.Get($"{item.Value}/reactions?offset=0&limit=5000");
                    List<string> likes = likesConsolidate[item.Value];

                    foreach (var like in likes)
                    {
                        cell = InsertCellInWorksheet(headerColumns[indexColumn], indexRow, worksheetPart);
                        cell.DataType = CellValues.InlineString;
                        cell.InlineString = new InlineString() { Text = new Text(RemoveTroublesomeCharacters(like.ToString())) };

                        // Save the worksheet.
                        worksheetPart.Worksheet.Save();

                        indexRow++;
                    }
                    indexColumn++;
                }

                // Save the worksheet.
                worksheetPart.Worksheet.Save();

                workbookPart.Workbook.Save();
            }
        }

        /// <summary>
        /// Removes control characters and other non-UTF-8 characters
        /// </summary>
        /// <param name="inString">The string to process</param>
        /// <returns>A string with no control characters or entities above 0x00FD</returns>
        public static string RemoveTroublesomeCharacters(string inString)
        {
            if (inString == null) return null;

            StringBuilder newString = new StringBuilder();
            char ch;

            for (int i = 0; i < inString.Length; i++)
            {

                ch = inString[i];

                if (XmlConvert.IsXmlChar(ch)) //this method is new in .NET 4
                {
                    newString.Append(ch);
                }
            }
            return newString.ToString();

        }

        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.
            if (row.Elements<Cell>().Where(c => c.CellReference != null && c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                Cell refCell = null;

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);
                worksheet.Save();
                return newCell;
            }
        }

        private Stylesheet GenerateStyleSheet()
        {
            return new Stylesheet(
                new Fonts(
                    new Font(
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" },
                        new Alignment() { WrapText = true }),
                    new Font(
                        new Bold(),
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" },
                        new Alignment() { WrapText = true }
                        ),
                    new Font(
                        new Italic(),

                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font(

                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Times New Roman" })
                ),
                new Fills(
                    new Fill(
                        new PatternFill() { PatternType = PatternValues.None }),
                    new Fill(
                        new PatternFill() { PatternType = PatternValues.Gray125 }),
                    new Fill(
                        new PatternFill(
                            new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } }
                        )
                        { PatternType = PatternValues.Solid })
                ),
                new Borders(
                    new Border(
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder()),
                    new Border(
                        new LeftBorder(
                            new Color() { Auto = true }
                        )
                        { Style = BorderStyleValues.Thin },
                        new RightBorder(
                            new Color() { Auto = true }
                        )
                        { Style = BorderStyleValues.Thin },
                        new TopBorder(
                            new Color() { Auto = true }
                        )
                        { Style = BorderStyleValues.Thin },
                        new BottomBorder(
                            new Color() { Auto = true }
                        )
                        { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                ),
                new CellFormats(
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 },
                    new CellFormat() { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true },
                    new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true },
                    new CellFormat() { FontId = 3, FillId = 0, BorderId = 0, ApplyFont = true },
                    new CellFormat() { FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true },
                    new CellFormat(
                        new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
                    )
                    { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }
                )
            );
        }


        public void CreateExcelDocLikes(string fileName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                string rId = "rId6";

                Sheet sheet = new Sheet() { Name = "Consolidate Likes", SheetId = 4U, Id = rId };
                workbookPart.Workbook.Sheets.Append(sheet);

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>(rId);

                worksheetPart.Worksheet = new Worksheet();

                workbookPart.Workbook.Save();

                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                // Constructing header
                Row row = new Row();
                Cell cell = ConstructCell("Consolidate User who liked", CellValues.String);
                cell.StyleIndex = 1;
                row.Append(cell);

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);

                foreach (var item in likesAll.OrderBy(q => q).Distinct().ToList())
                {
                    row = new Row();
                    row.Append(ConstructCell(item, CellValues.String));

                    // Insert the header row to the Sheet Data
                    sheetData.AppendChild(row);
                }

                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();

            }
        }

        private Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(RemoveTroublesomeCharacters(value)),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }

        protected void btnSearch_Click(object sender, EventArgs e)
        {
            var idGroup = lstGroups.SelectedValue;

            if (String.IsNullOrEmpty(idGroup))
            {
                foreach (ListItem item in this.lstGroups.Items)
                    idGroup = item.Value;
            }

            loadGroupFeeds(idGroup);
        }
    }
}