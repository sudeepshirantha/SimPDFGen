using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using SelectPdf;
using System.Web.Routing;
using Biz.ImageLib;

namespace SimPDFGen
{
    public class ReportGenerator
    {
        public static void GenereateTextFile(string outPath, string docBase, string outFileName, string text)
        {
            //path for the PDF file to be generated
            string docDir = string.Format("{0}/{1}", System.Web.Hosting.HostingEnvironment.MapPath(docBase), outPath);
            string outFileFullPath = string.Format("{0}/{1}", docDir, outFileName);
            if (System.IO.Directory.Exists(docDir))
            {
                try
                {
                    if (System.IO.File.Exists(outFileFullPath))
                    {
                        System.IO.File.Delete(outFileFullPath);
                    }
                }
                catch { }
            }
            else
            {
                System.IO.Directory.CreateDirectory(docDir);
            }

            // Creates the file on server
            System.IO.File.WriteAllText(outFileFullPath, text, System.Text.ASCIIEncoding.Unicode);
        }



        public static string GenereateWordFile(string outPath, string docBase, string outFileName, string doc)
        {
            //path for the PDF file to be generated
            string docDir = string.Format("{0}/{1}", System.Web.Hosting.HostingEnvironment.MapPath(docBase), outPath);
            string outFileFullPath = string.Format("{0}/{1}", docDir, outFileName);
            if (System.IO.Directory.Exists(docDir))
            {
                try
                {
                    if (System.IO.File.Exists(outFileFullPath))
                    {
                        System.IO.File.Delete(outFileFullPath);
                    }
                }
                catch { }
            }
            else
            {
                System.IO.Directory.CreateDirectory(docDir);
            }

            // Creates the file on server
            System.IO.File.WriteAllText(outFileFullPath, doc, System.Text.ASCIIEncoding.Unicode);

            return outFileFullPath;
        }


        public static void GeneratePDF(string outPath, string docBase, string outFileName, string html, string logoPrifix, string title, int? totalPages)
        {
            GeneratePDF(outPath, docBase, outFileName, html, logoPrifix, title, null, totalPages);
        }

        public static void GeneratePDF(string outPath, string docBase, string outFileName, string html, string logoPrifix, string title, string footerWarning, int? totalPages)
        {
            ImageUtil.ValidateInit();

            //path for the PDF file to be generated
            string docDir = string.Format("{0}/{1}", System.Web.Hosting.HostingEnvironment.MapPath(docBase), outPath);
            string outFileFullPath = string.Format("{0}/{1}", docDir, outFileName);
            if (System.IO.Directory.Exists(docDir))
            {
                try
                {
                    if (System.IO.File.Exists(outFileFullPath))
                    {
                        System.IO.File.Delete(outFileFullPath);
                    }
                }
                catch { }
            }
            else
            {
                System.IO.Directory.CreateDirectory(docDir);
            }

            try
            {
                HtmlToPdf converter = new HtmlToPdf();

                converter.Options.PdfPageSize = PdfPageSize.A4;
                converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
                converter.Options.MarginLeft = 30;
                converter.Options.MarginRight = 30;
                converter.Options.MarginTop = 5;
                converter.Options.MarginBottom = 10;

                converter.Options.WebPageWidth = 720;
                if (string.IsNullOrEmpty(footerWarning))
                {
                    converter.Options.WebPageHeight = 800;
                }
                else
                {
                    converter.Options.WebPageHeight = 600;
                }
                // header settings
                converter.Options.DisplayHeader = true;
                converter.Header.DisplayOnFirstPage = true;
                converter.Header.DisplayOnOddPages = true;
                converter.Header.DisplayOnEvenPages = true;
                converter.Header.Height = 80;

                // footer settings
                converter.Options.DisplayFooter = true;
                converter.Footer.DisplayOnFirstPage = true;
                converter.Footer.DisplayOnOddPages = true;
                converter.Footer.DisplayOnEvenPages = true;
                if (!string.IsNullOrEmpty(footerWarning))
                {
                    converter.Footer.Height = 80;
                }
                else
                {
                    converter.Footer.Height = 60;
                }

                int y = 0;
                if (!string.IsNullOrEmpty(footerWarning))
                {
                    PdfTextSection text2 = new PdfTextSection(0, y,
                        footerWarning,
                        new System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold));
                    text2.HorizontalAlign = PdfTextHorizontalAlign.Center;
                    text2.ForeColor = System.Drawing.Color.Red;
                    converter.Footer.Add(text2);
                    y += 25;
                }

                PdfHtmlSection footerHtml = new PdfHtmlSection(0, y, footerContent, null);
                footerHtml.AutoFitWidth = HtmlToPdfPageFitMode.AutoFit;
                converter.Footer.Add(footerHtml);
                y += 40;

                // add page numbering element to the footer
                // page numbers can be added using a PdfTextSection object
                PdfTextSection text = new PdfTextSection(0, y,
                    title + " Page: {page_number} of {total_pages}  ",
                    new System.Drawing.Font("Arial", 8));
                text.HorizontalAlign = PdfTextHorizontalAlign.Center;
                converter.Footer.Add(text);

                // create a new pdf document converting an url
                PdfDocument doc = converter.ConvertHtmlString(html);


                // create memory stream to save PDF
                MemoryStream pdfStream = new MemoryStream();

                // save pdf document into a MemoryStream
                doc.Save(pdfStream);

                // reset stream position
                pdfStream.Position = 0;

                // close pdf document
                doc.Close();

                System.IO.File.WriteAllBytes(outFileFullPath, pdfStream.ToArray());

            }
            finally
            {

            }
        }




        /// <summary>
        /// Convert relative link to absolute link
        /// </summary>
        /// <param name="input">The input.</param>
        /// <returns></returns>
        private static string FormatImageLinks(string input)
        {
            if (input == null)
                return string.Empty;
            string tempInput = input;
            const string pattern = @"<img(.|\n)+?>";

            //Change the relative URL's to absolute URL's for an image, if any in the HTML code.
            foreach (Match m in Regex.Matches(input, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.RightToLeft))
            {
                if (m.Success)
                {
                    string tempM = m.Value;
                    string pattern1 = "src=[\'|\"](.+?)[\'|\"]";
                    Regex reImg = new Regex(pattern1, RegexOptions.IgnoreCase | RegexOptions.Multiline);
                    Match mImg = reImg.Match(m.Value);

                    if (mImg.Success)
                    {
                        string src = mImg.Value.ToLower().Replace("src=", "").Replace("\"", "").Replace("\'", "");

                        if (!src.StartsWith("http://") && !src.StartsWith("https://"))
                        {
                            //Insert new URL in img tag
                            src = "src=\"" + System.Web.Hosting.HostingEnvironment.MapPath(src) + "\"";
                            try
                            {
                                tempM = tempM.Remove(mImg.Index, mImg.Length);
                                tempM = tempM.Insert(mImg.Index, src);

                                //insert new url img tag in whole html code
                                tempInput = tempInput.Remove(m.Index, m.Length);
                                tempInput = tempInput.Insert(m.Index, tempM);
                            }
                            catch (Exception)
                            {

                            }
                        }
                    }
                }
            }
            return tempInput;
        }

    }
}
