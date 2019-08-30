using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using SharePointPnP.Modernization.Framework.Publishing;
using SharePointPnP.Modernization.Framework.Transform;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            string siteUrl = "https://fisdemo.sharepoint.com/sites/TestWikiSite/";
            string userName = "gsheth@fisdemo.net";
            AuthenticationManager am = new AuthenticationManager();
            using (var cc = am.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, userName, "p@ssw0rD1)#RCIndia"))
            {
                var targetContext = cc.Clone("https://fisdemo.sharepoint.com/sites/TestCommSiteForSiteDesign");
                //var pageTransformator = new PublishingPageTransformator(cc, targetContext);
                var pageFile = cc.Web.GetFileByServerRelativeUrl("/sites/TestWikiSite/SitePages/testwppage101.aspx");
                cc.Load(pageFile, p => p.ListItemAllFields);
                cc.ExecuteQuery();

                var page = pageFile.ListItemAllFields;


                //PublishingPageTransformator publishingPageTransformator = new PublishingPageTransformator(cc, targetContext, "C:\\Users\\GautamSheth\\Documents\\GitHub\\sp-dev-modernization\\Tools\\SharePoint.Modernization\\ConsoleApp1\\bin\\Debug\\webpartmapping.xml", "C:\\FIS\\Test\\custompagelayoutmapping.xml");
                //PublishingPageTransformator publishingPageTransformator = new PublishingPageTransformator(cc, targetContext);
                //PublishingPageTransformationInformation pti = new PublishingPageTransformationInformation(page)
                //{
                //    Overwrite = true,
                //    RemoveEmptySectionsAndColumns = false,
                //    SkipTelemetry = true,
                //    HandleWikiImagesAndVideos = true,
                //    DisablePageComments = true,
                //};


                PageTransformator pageTransformator = new PageTransformator(cc, targetContext);

                PageTransformationInformation pti = new PageTransformationInformation(page)
                {
                    HandleWikiImagesAndVideos = true,
                    Overwrite = true,
                    SkipTelemetry = true,
                    AddTableListImageAsImageWebPart = true,
                    CopyPageMetadata = true,
                };

                try
                {
                    Console.WriteLine($"Transforming page {page.FieldValues["FileLeafRef"]}");

                    pageTransformator.Transform(pti);

                    //publishingPageTransformator.Transform(pti);

                    //var pageTemp = targetContext.Web.LoadClientSidePage("TestPageforTransform.aspx");
                    //pageTemp.SaveAsTemplate("TestPageforTransform.aspx");
                }
                catch (ArgumentException ex)
                {
                    Console.WriteLine($"Page {page.FieldValues["FileLeafRef"]} could not be transformed: {ex.Message}");
                }
            }
        }
    }
}
