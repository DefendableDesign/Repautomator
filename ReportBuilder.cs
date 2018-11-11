using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Linq;
using static Repautomator.OpenXmlTools;
using drc = DocumentFormat.OpenXml.Drawing.Charts;
using wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Repautomator
{
    public static class ReportBuilder
    {
        public static MemoryStream BuildReport(IConfigurationRoot config)
        {
            //Read the template from disk
            byte[] byteArray = File.ReadAllBytes(config["ReportConfiguration:TemplateFile"]);
            //Load it into memory
            MemoryStream mem = new MemoryStream();
            mem.Write(byteArray, 0, byteArray.Length);
            //Open the word document from memory
            using (WordprocessingDocument output = WordprocessingDocument.Open(mem, true))
            {
                var doc = output.MainDocumentPart.Document;
                var contentControls = output.ContentControls().ToList();

                foreach (var cc in contentControls)
                {
                    wp.SdtProperties props = cc.Elements<wp.SdtProperties>().FirstOrDefault();
                    string controlTitle = props.Elements<wp.SdtAlias>().FirstOrDefault().Val;
                    string controlTag = props.Elements<wp.Tag>().FirstOrDefault().Val;
                    string configKeyResult = String.Format("{0}:Result", controlTitle);
                    string queryResult = config[configKeyResult];
                    
                    Type sdtType = ((wp.SdtElement)cc).GetType();
                    OpenXmlElement targetContent = null;

                    //Determine if the content control is actually a docpart (Table of Contents, etc.), so we can ignore it.
                    bool isDocPart = false;
                    //If the content control document tag is a block tag and has docpart object children, it's a docpart.
                    if (sdtType == typeof(wp.SdtBlock) && (((wp.SdtBlock)cc).Descendants<wp.SdtContentDocPartObject>().Count() > 0)) { isDocPart = true; }

                    if (!isDocPart)
                    {
                        //Replace our ReportParameter placeholders.
                        if (controlTag.Equals("ReportParameter"))
                        {
                            if (config[controlTitle] != null)
                            {
                                targetContent = new wp.Text(config[controlTitle]);
                            }
                            else
                            {
                                targetContent = new wp.Text("No Result");
                            }
                        }

                        //Replace our SingleValue placeholders with tables containing the query results.
                        if (controlTag.Equals("SingleValue"))
                        {
                            if (queryResult != null) {
                                dynamic svData = JsonConvert.DeserializeObject<dynamic>(queryResult);
                                if (svData.rows.Count == 1)
                                {
                                    targetContent = new wp.Text(svData.rows[0][0].ToString());
                                }
                                else
                                {
                                    throw new IndexOutOfRangeException(String.Format("The QueryResult for {0} contained {1} row/s but a single row was expected.", controlTitle, svData.rows.Count));
                                }
                                
                            } else
                            {
                                targetContent = new wp.Text("No Result");
                            }
                        }

                        //Special handling is needed if a SingleValue or ReportParameter content control is actually configured as a SdtBlock instead of SdtRun.
                        if (controlTag.Equals("ReportParameter") || controlTag.Equals("SingleValue")) {                            
                            if (sdtType == typeof(wp.SdtBlock)) {
                                var targetContentParent = cc.GetFirstChild<wp.SdtContentBlock>().GetFirstChild<wp.Paragraph>().CloneNode(true);
                                targetContentParent.GetFirstChild<wp.Run>().InsertAfterSelf(new wp.Run(targetContent));
                                targetContentParent.GetFirstChild<wp.Run>().Remove();
                                targetContent = targetContentParent;
                            }
                        }

                        //Replace our Table placeholders with tables containing the query results.
                        if (controlTag.Equals("Table"))
                        {
                            targetContent = BuildTable(queryResult, config["ReportConfiguration:TableStyle"]);
                        }

                        //Update the charts in our Chart placeholders with the data from the query results.
                        if (controlTag.Equals("Chart"))
                        {
                            UpdateChart(
                                cc.Descendants<drc.ChartReference>().FirstOrDefault(), 
                                output, 
                                controlTitle.Split(':').Last(), 
                                queryResult);
                            targetContent = cc.GetFirstChild<wp.SdtContentBlock>().GetFirstChild<wp.Paragraph>().CloneNode(true);
                        }

                        //Remove PlaceholderText styles from all Run elements
                        foreach (var item in cc.Descendants<wp.Run>())
                        {
                            var rPr = item.GetFirstChild<wp.RunProperties>();
                            if (rPr != null && rPr.RunStyle != null) { if (rPr.RunStyle.Val == "PlaceholderText") { rPr.RunStyle.Remove(); } }
                        }

                        if (targetContent != null)
                        {
                            //The content control is inline
                            if (sdtType == typeof(wp.SdtRun))
                            {
                                cc.InsertAfterSelf(new wp.Run(targetContent));
                                cc.Remove();
                            }
                            //The content control is a table cell
                            else if (sdtType == typeof(wp.SdtCell))
                            {

                                wp.TableCell tc = (wp.TableCell)((cc.Descendants<wp.TableCell>().FirstOrDefault()).CloneNode(true));

                                if (targetContent.GetType() == typeof(wp.Text))
                                {
                                    wp.Run targetRun = tc.Descendants<wp.Run>().FirstOrDefault();
                                    targetRun.RemoveAllChildren();
                                    targetRun.Append(targetContent);
                                }
                                else
                                { tc.Append(targetContent); }

                                cc.InsertAfterSelf(tc);
                                cc.Remove();
                            }
                            //The content control is multi-line
                            else if (sdtType == typeof(wp.SdtBlock))
                            {
                                cc.InsertAfterSelf(targetContent);
                                cc.Remove();
                            }
                        }
                    }
                }

                foreach (var footer in output.MainDocumentPart.FooterParts)
                {
                    footer.Footer.Save();
                }

                doc.Save();
                output.Close();

                return mem;
            }
        }
    }
}
