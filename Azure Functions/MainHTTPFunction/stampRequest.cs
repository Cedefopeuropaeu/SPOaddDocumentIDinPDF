using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;
using PnP.Core.Services;
using System.Net.Http;
using PnP.Core.Auth;
using System.Collections.Generic;
using PnP.Core.Model;
using System.Security;
using PnP.Core.Model.SharePoint;
using System.Linq;
using PnP.Core.QueryModel;
using System.Web;
using PnP.Core;
using System.Linq.Expressions;
using static Microsoft.AspNetCore.Hosting.Internal.HostingApplication;
using PnP.Core.Model.Teams;
using iText.IO.Source;
using iText.Kernel.Pdf;
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Annot;
using iText.Layout;
using iText.Kernel.Colors;
using iText.Kernel.Pdf.Canvas;

[assembly: FunctionsStartup(typeof(MainHTTPFunction.Startup))]

namespace MainHTTPFunction
{
    public class stampRequest
    {

        private IPnPContextFactory _pnpContextFactory;

        public stampRequest(IPnPContextFactory pnpContextFactory)
        {
            _pnpContextFactory = pnpContextFactory;
        }

        [FunctionName("stampRequest")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequestMessage request,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            AppInfo _appInfo = new AppInfo();
            try
            {
                _appInfo.ClientId = System.Environment.GetEnvironmentVariable("ClientId", EnvironmentVariableTarget.Process);
                _appInfo.ClientSecret = System.Environment.GetEnvironmentVariable("ClientSecret", EnvironmentVariableTarget.Process);
            }
            catch(Exception ex)
            {
                return new BadRequestObjectResult(ex.Message);
            }
            string body;
            try
            {
                 body = request.Content.ReadAsStringAsync().Result;
            }
            catch(Exception ex)
            {
                return new BadRequestObjectResult(ex.Message);
            }

            BodyPDFInfo bodyParams = JsonConvert.DeserializeObject<BodyPDFInfo>(body);
            
            var clientSecret = new SecureString();
            foreach (char c in _appInfo.ClientSecret) clientSecret.AppendChar(c);

            var onBehalfAuthProvider = new OnBehalfOfAuthenticationProvider(_appInfo.ClientId, bodyParams.tenantId, clientSecret, () => request.Headers.Authorization.Parameter);
            ////var results = new List<ListData>();
            try
            {
                using (var pnpContext = await _pnpContextFactory.CreateAsync(new System.Uri(bodyParams.siteUrl), onBehalfAuthProvider))
                {
                    try
                    {
                        IFile fileToDownload = await pnpContext.Web.GetFileByServerRelativeUrlAsync(bodyParams.fileRef, w => w.ListItemAllFields, w => w.CheckOutType, w => w.CheckedOutByUser);

                        if(fileToDownload.CheckOutType != CheckOutType.None)
                        {
                            string message = "The file is checked-out by " + fileToDownload.CheckedOutByUser.UserPrincipalName + ". Please check-in the file before proceeding.";
                            return new BadRequestObjectResult(message);
                        }

                        IFolder parentFolder = await fileToDownload.ListItemAllFields.GetParentFolderAsync();

                        Stream downloadedContentStream = await fileToDownload.GetContentAsync(true);
                        var bufferSize = 2 * 1024 * 1024;  // 2 MB buffer
                                                           //string tempfile = System.IO.Path.GetTempFileName();
                        using (var contentStream = new MemoryStream())
                        {
                            var buffer = new byte[bufferSize];
                            int read;
                            while ((read = await downloadedContentStream.ReadAsync(buffer, 0, buffer.Length)) != 0)
                            {
                                contentStream.Write(buffer, 0, read);
                            }
                            byte[] content = contentStream.ToArray();
                            StampingProperties stampingProperties = new StampingProperties();
                            stampingProperties.UseAppendMode().PreserveEncryption();

                            RandomAccessSourceFactory randomAccessSourceFactory = new RandomAccessSourceFactory();
                            IRandomAccessSource randomAccessSource = randomAccessSourceFactory.CreateSource(content);
                            PdfReader reader = new PdfReader(randomAccessSource, new ReaderProperties());
                            var outStream = new MemoryStream();
                            PdfWriter writer = new PdfWriter(outStream);


                            PdfDocument pdfDoc = new PdfDocument(reader, writer, stampingProperties);



                            Color gray = new DeviceRgb(126, 126, 126);

                            PdfPage page = pdfDoc.GetPage(1);

                            Rectangle pageSize = page.GetPageSize();
                            Rectangle rectOuter = new Rectangle(pageSize.GetRight() - 143, pageSize.GetTop() - 17, 133, 13);
                            Rectangle rectText = new Rectangle(pageSize.GetRight() - 143, pageSize.GetTop() - 17, 133, 13);

                            PdfSquareAnnotation squareAnnotation = new PdfSquareAnnotation(rectOuter);
                            squareAnnotation.SetNonStrokingOpacity(0.25f);
                            squareAnnotation.SetOpacity(new PdfNumber(0.30f));
                            squareAnnotation.SetInteriorColor(new float[] { (float)0.64, (float)0.64, (float)0.64 });
                            squareAnnotation.SetFlags(PdfAnnotation.PRINT);
                            squareAnnotation.SetColor(ColorConstants.BLACK);
                            squareAnnotation.SetText(new PdfString(bodyParams.docID));
                            page.AddAnnotation(squareAnnotation);

                            PdfAnnotation freeText = new PdfFreeTextAnnotation(rectText, new PdfString(bodyParams.docID))
                                .SetDefaultAppearance(new PdfString("/Helv 10 Tf 0 g"))
                                .SetFlags(PdfAnnotation.PRINT)
                                .SetColor(ColorConstants.LIGHT_GRAY);
                            page.AddAnnotation(freeText);

                            pdfDoc.SetCloseWriter(false);
                            pdfDoc.Close();
                            
                            await fileToDownload.CheckoutAsync();
                            IFile addedFile;
                            try
                            {
                                
                                outStream.Seek(0, SeekOrigin.Begin);
                                addedFile = await parentFolder.Files.AddAsync(bodyParams.FileLeafRef, outStream, true);
                                //addedFile = await parentFolder.Files.AddAsync("test.pdf", outStream, true);
                            }
                            catch (Exception ex)
                            {
                                await fileToDownload.UndoCheckoutAsync();
                                if (ex is PnPException)
                                {
                                    PnPException x = (PnPException)ex;
                                    return new BadRequestObjectResult(x.Error.ToString());
                                }
                                else
                                {
                                    return new BadRequestObjectResult(ex.Message);
                                }
                            }
                            await addedFile.ListItemAllFields.LoadAsync();
                            addedFile.ListItemAllFields["hasStamp"] = true;
                            try
                            {
                                await addedFile.ListItemAllFields.UpdateAsync();
                                await fileToDownload.CheckinAsync();
                            }
                            catch (Exception ex)
                            {
                                await fileToDownload.UndoCheckoutAsync();
                                if (ex is PnPException)
                                {
                                    PnPException x = (PnPException)ex;
                                    return new BadRequestObjectResult(x.Error.ToString());
                                }
                                else
                                {
                                    return new BadRequestObjectResult(ex.Message);
                                }
                            }




                            //PdfDocument pdf = PdfReader.Open(contentStream);
                            //PdfPage firstpage = pdf.Pages[0];
                            ////var gfx = XGraphics.FromPdfPage(firstpage);

                            ////string watermark = bodyParams.docID;
                            //var gfx = XGraphics.FromPdfPage(firstpage, XGraphicsPdfPageOptions.Prepend);

                            //XFont font = new XFont("Times New Roman", 12, XFontStyle.BoldItalic);
                            //XSize size = gfx.MeasureString(bodyParams.docID, font);
                            //double width = Convert.ToDouble(firstpage.Width);
                            //double height = Convert.ToDouble(firstpage.Height);
                            ////var brush = new XSolidBrush(XColor.FromKnownColor(XKnownColor.LightGray));
                            ////gfx.DrawRectangle(brush, (width - size.Width - 11), 5, 170, 16);
                            ////XStringFormat format = new XStringFormat();
                            ////format.Alignment = XStringAlignment.Near;
                            ////format.LineAlignment = XLineAlignment.Near;
                            //var textAnnot = new PdfTextAnnotation();
                            //textAnnot.Color = XColor.FromKnownColor(XKnownColor.Black);
                            //textAnnot.Title = bodyParams.docID;
                            //textAnnot.Subject = "This is the subject";
                            //textAnnot.Contents = bodyParams.docID;
                            //textAnnot.Icon = PdfTextAnnotationIcon.Note;
                            ////textAnnot.Icon = PdfTextAnnotationIcon.NoIcon;
                            ////gfx.DrawString(bodyParams.docID, font, XBrushes.Black, new XPoint((width - size.Width - 5), 5), format);
                            ////var rect = gfx.Transformer.WorldToDefaultPage(new XRect(new XPoint((width - size.Width - 35), 2), new XSize(1, 1)));
                            //textAnnot.Rectangle = new PdfRectangle(gfx.Transformer.WorldToDefaultPage(new XRect(new XPoint(30, 60), new XSize(30, 30))));
                            //textAnnot.Opacity = 0.5;
                            //textAnnot.Open = true;
                            //firstpage.Annotations.Add(textAnnot);

                            //PdfRubberStampAnnotation rsAnnot = new PdfRubberStampAnnotation();
                            //rsAnnot.Icon = PdfRubberStampAnnotationIcon.TopSecret;
                            //rsAnnot.Flags = PdfAnnotationFlags.ReadOnly;

                            //rsAnnot.Color = XColor.FromKnownColor(XKnownColor.Black);
                            //rsAnnot.Contents = bodyParams.docID;
                            //XRect rect = gfx.Transformer.WorldToDefaultPage(new XRect(new XPoint(30, 70), new XSize(30, 30)));
                            //rsAnnot.Rectangle = new PdfRectangle(rect);

                            // Add the rubber stamp annotation to the page
                            //firstpage.Annotations.Add(rsAnnot);



                            //MemoryStream outstream = new MemoryStream();
                            //pdf.Save(outstream, false);



                        }
                    }
                    catch (Exception ex)
                    {
                        if (ex is PnPException)
                        {
                            PnPException x = (PnPException) ex;
                            return new BadRequestObjectResult(x.Error.ToString());
                        }
                        else
                        {
                            return new BadRequestObjectResult(ex.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (ex is PnPException)
                {
                    PnPException x = (PnPException)ex;
                    return new BadRequestObjectResult(x.Error.ToString());
                }
                else
                {
                    return new BadRequestObjectResult(ex.Message);
                }
            }
            return new OkObjectResult("ΟΚ");
        }

    }
    public class Startup : FunctionsStartup
    {
        public override void Configure(IFunctionsHostBuilder builder)
        {
            builder.Services.AddPnPCore();
            builder.Services.AddPnPCoreAuthentication();
        }
    }
    public class AppInfo
    {
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
    }
}
