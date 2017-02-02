using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace WebSiteWithFacebookIntegration
{
    public class DownloadFile : IHttpHandler
    {
        private string fileName { get; set; }

        public DownloadFile(string fileName)
        {
            this.fileName = fileName;
        }

        public void ProcessRequest(HttpContext context)
        {
            var extension = Path.GetExtension(fileName);

            System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;
            response.ClearContent();
            response.Clear();
            response.ContentType = "text/plain";
            response.AddHeader("Content-Disposition",
                               "attachment; filename=file"+ extension + ";");
            response.TransmitFile(fileName);
            response.Flush();
            response.End();
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}