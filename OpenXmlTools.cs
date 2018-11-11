using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using dr = DocumentFormat.OpenXml.Drawing;
using drc = DocumentFormat.OpenXml.Drawing.Charts;
using ss = DocumentFormat.OpenXml.Spreadsheet;
using wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Repautomator
{
    public static class OpenXmlTools
    {
        /// <summary>
        /// Creates a new Table from the supplied input data.
        /// </summary>
        /// <param name="json">
        /// String in JSON format representing the tabular data for updating the Chart's cached data points.
        /// The JSON object must contain a "fields" attribute as an array containing the field/column names. 
        /// The JSON object must contain a "rows" attribute as an array of arrays representing the rows and their values, with values matching the same order and cardinality of the field names.
        /// This is the same data as the underlying Excel spreadsheet contents.</param>
        /// <param name="tableStyle">
        /// String containing the name of the Wordprocessing.TableStyle to apply to the table.</param>
        /// <returns>
        /// Returns a new Wordprocessing.Table containing the tabular data from the JSON input, formatted with the specified TableStyle.</returns>
        public static wp.Table BuildTable(string json, string tableStyle)
        {
            json = ((json == String.Empty) || (json == null)) ? "{\"fields\": [ \"No Results\" ], \"rows\": [[ \"No Results\" ]]}" : json;

            //Splunk JSON data is a series of objects consisting of multiple key(column)/value(row) pairs in the result attribute.
            dynamic input = JsonConvert.DeserializeObject<dynamic>(json);

            if (input["rows"].Count == 0)
            {
                json = "{\"fields\": [ \"No Results\" ], \"rows\": [[ \"No Results\" ]]}";
                input = JsonConvert.DeserializeObject<dynamic>(json);
            }

            wp.Table result = new wp.Table();

            wp.TableProperties tableProperties1 = new wp.TableProperties();
            wp.TableStyle tableStyle1 = new wp.TableStyle() { Val = tableStyle };
            wp.TableWidth tableWidth1 = new wp.TableWidth() { Width = "5000", Type = wp.TableWidthUnitValues.Pct };

            tableProperties1.Append(tableStyle1);
            tableProperties1.Append(tableWidth1);
            result.Append(tableProperties1);


            wp.TableGrid tableGrid = new wp.TableGrid();

            //Build table header row
            wp.TableRow headerRow = new wp.TableRow();
            foreach (var columnName in input["fields"])
            {
                headerRow.Append(new wp.TableCell(new wp.Paragraph(new wp.Run(new wp.Text(columnName.ToString())))));
                tableGrid.Append(new wp.GridColumn());
            }
            result.Append(tableGrid);
            result.Append(headerRow);

            //Build table data rows
            foreach (var row in input["rows"])
            {
                wp.TableRow tr = new wp.TableRow();
                foreach (var cell in row)
                {
                    tr.Append(new wp.TableCell(new wp.Paragraph(new wp.Run(new wp.Text(cell.ToString())))));
                }
                result.Append(tr);
            }

            return result;
        }

        /// <summary>
        /// Updates a SpreadsheetDocument with new tabular data.
        /// </summary>
        /// <param name="ssdoc">
        /// The excel document to update.</param>
        /// <param name="json">
        /// String in JSON format representing the tabular data for updating the Chart's cached data points.
        /// The JSON object must contain a "fields" attribute as an array containing the field/column names. 
        /// The JSON object must contain a "rows" attribute as an array of arrays representing the rows and their values, with values matching the same order and cardinality of the field names.
        /// This is the same data as the underlying Excel spreadsheet contents.</param>
        /// <param name="sheetName">
        /// The name of the Excel worksheet where the chart data originates from.
        /// Used for updating the chart's cell references.</param>
        /// <returns>
        /// Returns the updated SpreadsheetDocument</returns>
        public static SpreadsheetDocument Update(this SpreadsheetDocument ssdoc, string json, string sheetName)
        {
            if ((json == null) || (json == String.Empty))
            {
                json = "{\"fields\": [ \"No Results\" ], \"rows\": [[ \"No Results\" ]]}";
            }
            
            //Splunk JSON data is a series of objects consisting of multiple key(column)/value(row) pairs in the result attribute.
            dynamic input = JsonConvert.DeserializeObject<dynamic>(json);

            if (input["rows"].Count == 0)
            {
                json = "{\"fields\": [ \"No Results\" ], \"rows\": [[ \"No Results\" ]]}";
                input = JsonConvert.DeserializeObject<dynamic>(json);
            }

            ss.Sheet sheet = ssdoc.WorkbookPart.Workbook.Descendants<ss.Sheet>().Where(s => s.Name.ToString() == sheetName).FirstOrDefault();
            if (sheet == null)
            {
                sheet = ssdoc.WorkbookPart.Workbook.Descendants<ss.Sheet>().FirstOrDefault();
            }
            WorksheetPart worksheet = (WorksheetPart)ssdoc.WorkbookPart.GetPartById(sheet.Id);
            ss.SheetData data = worksheet.Worksheet.GetFirstChild<ss.SheetData>();



            //Remove all the rows after our column headers row. We'll replace them with new rows as the table is populated from the Splunk search results.
            ss.Row firstRow = data.Elements<ss.Row>().First();
            while (firstRow.NextSibling<ss.Row>() != null) { firstRow.NextSibling<ss.Row>().Remove(); }

            ss.Row newHeader = new ss.Row();
            newHeader.DyDescent = 0.25;
            newHeader.RowIndex = 1;
            var columnNames = input["fields"];
            char startingHeaderColumn = 'A';
            foreach (var column in columnNames)
            {
                string cellRef = startingHeaderColumn.ToString() + 1;
                ss.Cell newCell = new ss.Cell();
                ss.CellValue cv = new ss.CellValue(column.ToString());

                newCell.CellReference = cellRef;
                newCell.DataType = ss.CellValues.String;
                newCell.Append(cv);
                newHeader.Append(newCell);

                startingHeaderColumn++;
            }

            data.InsertAfter(newHeader, data.Elements<ss.Row>().Last());
            data.RemoveChild(firstRow);

            for (int i = 0; i < input["rows"].Count; i++)
            {

                char startingColumn = 'A';
                int startingColumnVal = 1;
                int endingColumnVal = 0;
                //Set the Excel row index (Excel index starts at 1, not zero and row 1 has headers so add 2 to the for index)
                uint rowIndex = Convert.ToUInt32(i) + 2;

                ss.Row newRow = new ss.Row();
                newRow.DyDescent = 0.25;
                newRow.RowIndex = rowIndex;


                Debug.WriteLine(String.Format("Writing Excel Row {0}", i + 1));
                var row = input["rows"][i];
                foreach (var cell in row)
                {
                    string cellRef = startingColumn.ToString() + rowIndex;
                    ss.Cell newCell = new ss.Cell();
                    ss.CellValue cv = new ss.CellValue(cell.ToString());

                    newCell.CellReference = cellRef;
                    if (startingColumn == 'A') { newCell.DataType = ss.CellValues.String; }
                    newCell.Append(cv);
                    newRow.Append(newCell);

                    startingColumn++;
                    endingColumnVal++;
                }


                int numberOfCellsInRow = newRow.Descendants<ss.Cell>().Count();
                ListValue<StringValue> spans = new ListValue<StringValue>();
                spans.Items.Add(string.Format("{0}:{1}", startingColumnVal, endingColumnVal));
                newRow.Spans = spans;
                data.InsertAfter(newRow, data.Elements<ss.Row>().Last());
            }


            // Update Table Reference                        
            Debug.WriteLine("Updating Table");
            var table = worksheet.TableDefinitionParts.First().Table;
            table.Reference = string.Format("A1:{0}{1}", GetExcelColumnName(input["fields"].Count), input["rows"].Count + 1);
            table.TableColumns.RemoveAllChildren();
            for (int i = 0; i < columnNames.Count; i++)
            {
                var newColumn = new ss.TableColumn();
                newColumn.Id = new UInt32Value((uint)i + 1);
                newColumn.Name = new StringValue(columnNames[i].ToString());
                table.TableColumns.Append(newColumn);
            }



            return ssdoc;
        }

        /// <summary>
        /// Updates a ChartPart's caches and excel reference formilas with new tabular data.
        /// Use in conjunction with SpreadsheetDocument.Update to update the underlying spreadsheet values.
        /// </summary>
        /// <param name="chart">
        /// The OpenXml Chart's ChartPart</param>
        /// <param name="json">
        /// String in JSON format representing the tabular data for updating the Chart's cached data points.
        /// The JSON object must contain a "fields" attribute as an array containing the field/column names. 
        /// The JSON object must contain a "rows" attribute as an array of arrays representing the rows and their values, with values matching the same order and cardinality of the field names.
        /// This is the same data as the underlying Excel spreadsheet contents.</param>
        /// <param name="sheetName">
        /// The name of the Excel worksheet where the chart data originates from.
        /// Used for updating the chart's cell references.</param>
        /// <returns>
        /// Returns the updated ChartPart.</returns>
        public static ChartPart Update(this ChartPart chart, string json, string sheetName)
        {
            //Splunk JSON data is a series of objects consisting of multiple key(column)/value(row) pairs in the result attribute.
            dynamic input = JsonConvert.DeserializeObject<dynamic>(json);

            if (input["rows"].Count == 0)
            {
                json = "{\"fields\": [ \"No Results\" ], \"rows\": [[ \"No Results\" ]]}";
                input = JsonConvert.DeserializeObject<dynamic>(json);
            }
            
            var pointCount = new drc.PointCount();
            pointCount.Val = Convert.ToUInt32(input["rows"].Count);

            // The number of columns in the data table is equal to the number of data series + 1 (for the x-axis column).
            // Subtract 1 from the column count to get the number of data series.
            int dataColumnCount = input["fields"].Count - 1;
            int seriesCount = chart.ChartSpace.Descendants().Where(e => e.GetType().ToString().EndsWith("ChartSeries")).Count();

            if (dataColumnCount < seriesCount)
            {
                // Fewer columns in the data source than there are series in the chart template.
                // Delete excess series from the chart template so we don't end up with old/dummy data in the chart.
                int removeCount = seriesCount - dataColumnCount;

                for (int i = 0; i < removeCount; i++)
                {
                    chart.ChartSpace.Descendants().Where(e => e.GetType().ToString().EndsWith("ChartSeries")).Last().Remove();
                }

                seriesCount = chart.ChartSpace.Descendants().Where(e => e.GetType().ToString().EndsWith("ChartSeries")).Count();
            }
            else if (dataColumnCount > seriesCount)
            {
                // More columns in the data source than there are series in the chart template.
                // Add some template series to the chart template so we don't drop any data off the chart.
                int insertCount = dataColumnCount - seriesCount;
                OpenXmlElement firstSeries = chart.ChartSpace.Descendants().Where(e => e.GetType().ToString().EndsWith("ChartSeries")).First();

                for (int i = 0; i < insertCount; i++)
                {
                    OpenXmlElement newSeries = (OpenXmlElement)firstSeries.Clone();
                    newSeries.Descendants<drc.Index>().FirstOrDefault().Val.Value = Convert.ToUInt32(seriesCount + i);
                    newSeries.Descendants<drc.Order>().FirstOrDefault().Val.Value = Convert.ToUInt32(seriesCount + i);

                    var csp = newSeries.Descendants<drc.ChartShapeProperties>().FirstOrDefault();
                    var newFill = CreateAccentSolidFill(seriesCount + i + 1);
                    csp.ChildElements.First<dr.SolidFill>().InsertAfterSelf(newFill);
                    csp.ChildElements.First<dr.SolidFill>().Remove();

                    firstSeries.Parent.AppendChild(newSeries);
                }
                seriesCount = chart.ChartSpace.Descendants().Where(e => e.GetType().ToString().EndsWith("ChartSeries")).Count();
            }

            // Use LINQ to get all elements of type *ChartSeries
            var seriesQuery =
                from elements in chart.ChartSpace.Descendants()
                where elements.GetType().ToString().EndsWith("ChartSeries")
                select elements;

            if (dataColumnCount == seriesCount)
            {
                //We're syncing two separate IEnumerables by index, so we should confirm that the sizes match.
                for (int seriesIndex = 0; seriesIndex < seriesQuery.Count(); seriesIndex++)
                {
                    var series = seriesQuery.ElementAt(seriesIndex);
                    Debug.WriteLine(string.Format("Updating series {0} of {1}", seriesIndex + 1, seriesQuery.Count()));

                    //Set the name of the chart series to the column name from from the data table.
                    //The first column contains axis data, so add 1 to get the series column.
                    drc.SeriesText seriesTitle = series.Descendants<drc.SeriesText>().FirstOrDefault();
                    seriesTitle.StringReference.Descendants<drc.NumericValue>().FirstOrDefault().Text = input["fields"][seriesIndex + 1].ToString();

                    var oldAxisData = series.Descendants<drc.CategoryAxisData>().FirstOrDefault();
                    var newAxisData = new drc.CategoryAxisData();

                    var sc = new drc.StringCache();
                    //Cloning the element saves extra lines of instantiation code.
                    sc.Append(pointCount.CloneNode(true));
                    var stringRef = new drc.StringReference(sc);

                    //var stringPoints = new List<OpenXmlElement>();
                    var oldValues = series.Descendants<drc.Values>().FirstOrDefault();
                    var newValues = new drc.Values();

                    // Don't think of a drawing chart like the underlying spreadsheet.
                    // In a drawing chart the axis (column A) is included in EVERY series (column B+).

                    //Prepare Axis Data
                    var axisColumn = GetExcelColumnName(1);
                    var axisFormula = string.Format("{0}!${1}${2}:${1}${3}", sheetName, axisColumn, 2, input["rows"].Count + 1);

                    //Axis data in a chart consists of a cache of string values and the underlying Excel formula to generate these data.
                    //We manually populate the cache as word won't automatically update 'external' documents (even ones which are internal to the document package), because security.
                    stringRef.Append(new drc.Formula(axisFormula));
                    newAxisData.Append(stringRef);

                    //Update the cache with the string values from the data source
                    for (int r = 0; r < input["rows"].Count; r++)
                    {
                        var row = input["rows"][r];
                        Debug.WriteLine("Writing Chart Axis");
                        //Axis data is always the first column (0)
                        var cell = input["rows"][r][0].ToString();
                        var point = new drc.StringPoint(new drc.NumericValue(cell));
                        point.Index = Convert.ToUInt32(r);
                        stringRef.StringCache.Append(point);
                    }

                    //Prepare Value Data
                    //Add 1 to make the index 1-based because excel columns start at A (== 1), add 1 again because the first column contains axis details, so our value columns will be one extra column over.
                    var dataColumn = GetExcelColumnName(seriesIndex + 2);
                    //Standard Excel data reference formula eg: "MyQuery!$B$2:$B$25" - The 'MyQuery' sheet from cell B2 to B25
                    var dataFormula = string.Format("{0}!${1}${2}:${1}${3}", sheetName, dataColumn, 2, input["rows"].Count + 1);

                    //Value data in a chart consists of a cache of point values and the underlying Excel formula to generate these data.
                    //We manually populate the cache as word won't automatically update 'external' documents (even ones which are internal to the document package), because security.
                    var nc = new drc.NumberingCache();
                    nc.Append(new drc.FormatCode("General"));
                    nc.Append(pointCount.CloneNode(true));
                    var numRef = new drc.NumberReference(nc);
                    numRef.Append(new drc.Formula(dataFormula));
                    newValues.Append(numRef);

                    //Update the cache with the point values from the data source
                    for (int r = 0; r < input["rows"].Count; r++)
                    {
                        var row = input["rows"][r];
                        Debug.WriteLine("Writing Chart Data");
                        var cell = input["rows"][r][seriesIndex + 1].ToString();
                        var point = new drc.NumericPoint(new drc.NumericValue(cell));
                        point.Index = Convert.ToUInt32(r);
                        numRef.NumberingCache.Append(point);
                    }

                    //Swap out the old elements for our newly created replacements.
                    oldValues.Parent.ReplaceChild(newValues, oldValues);
                    oldAxisData.Parent.ReplaceChild(newAxisData, oldAxisData);
                }
            }
            else
            {
                //We should never end up here, something has gone very wrong.
                throw new IndexOutOfRangeException("The number of series in the chart does not match the number of data columns in the data table.");
            }

            return chart;
        }

        /// <summary>
        /// Updates a Chart and its underlying embedded SpreadsheetDocument with new data.
        /// </summary>
        /// <param name="chartRef">
        /// The reference to the chart to be updated</param>
        /// <param name="doc">
        /// The WordprocessingDocument containing the chart to be updated.</param>
        /// <param name="sheetName">
        /// The name of the Excel worksheet where the chart data originates from.
        /// Used for updating the chart's cell references.</param>
        /// <param name="json">
        /// String in JSON format representing the tabular data for updating the Chart's cached data points.
        /// The JSON object must contain a "fields" attribute as an array containing the field/column names. 
        /// The JSON object must contain a "rows" attribute as an array of arrays representing the rows and their values, with values matching the same order and cardinality of the field names.
        /// This is the same data as the underlying Excel spreadsheet contents.</param>
        /// <returns>Returns the updated WordprocessingDocument</returns>
        public static void UpdateChart(drc.ChartReference chartRef, WordprocessingDocument doc, string sheetName, string json)
        {
            if ((json == null) || (json == String.Empty)) { json = "{\"fields\": [], \"rows\": []}"; }

            //Splunk JSON data is a series of objects consisting of multiple key(column)/value(row) pairs in the result attribute.
            dynamic input = JsonConvert.DeserializeObject<dynamic>(json);

            //No results for the chart so we will remove and replace it with an error.
            if (input["rows"].Count == 0)
            {
                //Build a replacement for the Content Control's contents with an error message.
                OpenXmlElement targetContent = new wp.Paragraph(new wp.Run(new wp.Text("Chart unavailable. No results found.")));
                //Remove the chart from the document
                doc.MainDocumentPart.DeletePart(chartRef.Id);
                //Insert our replacement Content Control after 
                var cc = chartRef.Ancestors().Where(e => e.LocalName.StartsWith("sdt")).FirstOrDefault();
                cc.RemoveAllChildren();
                cc.Append(targetContent);
            }
            else
            {
                ChartPart chart = (ChartPart)doc.MainDocumentPart.GetPartById(chartRef.Id);
                chart.ChartSpace.GetFirstChild<drc.ExternalData>().AutoUpdate.Val = true;

                // Update Chart
                chart = chart.Update(json, sheetName);

                SpreadsheetDocument ssdoc = SpreadsheetDocument.Open(chart.EmbeddedPackagePart.GetStream(FileMode.Open, FileAccess.ReadWrite), true);
                ssdoc.Update(json, sheetName);
                ssdoc.WorkbookPart.Workbook.Save();
                ssdoc.Close();
                Debug.WriteLine(String.Format("{0}: Chart and Spreadsheet updated successfully.", sheetName));
            }
        }

        /// <summary>
        /// Gets the Content Controls in the OpenXmlPart.
        /// </summary>
        /// <param name="part">The OpenXMLPart to get the Content Controls from.</param>
        /// <returns>
        /// Returns an IEnumerable of Content Controls from the part.</returns>
        public static IEnumerable<OpenXmlElement> ContentControls(this OpenXmlPart part)
        {
            return part.RootElement
                .Descendants()
                .Where(e => e is wp.SdtBlock || e is wp.SdtRun || e is wp.SdtCell);
        }

        /// <summary>
        /// Gets the Content Controls in the WordprocessingDocument.
        /// </summary>
        /// <param name="doc">The WordprocessingDocument to get the Content Controls from.</param>
        /// <returns>
        /// Returns a an IEnumerable of Content Controls from the document.</returns>
        public static IEnumerable<OpenXmlElement> ContentControls(this WordprocessingDocument doc)
        {
            foreach (var cc in doc.MainDocumentPart.ContentControls())
                yield return cc;
            foreach (var header in doc.MainDocumentPart.HeaderParts)
                foreach (var cc in header.ContentControls())
                    yield return cc;
            foreach (var footer in doc.MainDocumentPart.FooterParts)
                foreach (var cc in footer.ContentControls())
                    yield return cc;
            if (doc.MainDocumentPart.FootnotesPart != null)
                foreach (var cc in doc.MainDocumentPart.FootnotesPart.ContentControls())
                    yield return cc;
            if (doc.MainDocumentPart.EndnotesPart != null)
                foreach (var cc in doc.MainDocumentPart.EndnotesPart.ContentControls())
                    yield return cc;
        }

        /// <summary>
        /// Converts an integer into the corresponding Excel column name.
        /// </summary>
        /// <param name="columnNumber">The integer corresponding to the column name. This is not zero-based: Column A == 1.</param>
        /// <returns>
        /// Returns the column name as a string.</returns>
        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        /// <summary>
        /// Office document themes contain six accent colours. This method will convert an integer into the corresponding Drawing.SolidFill accent value.
        /// For values which exceed the six accent colours, the SolidFill will be progressively tinted 20% lighter.
        /// </summary>
        /// <param name="value">The integer representing the colour value. This is zero-based: 0 = Drawing.SchemeColorValues.Accent1</param>
        /// <returns></returns>
        private static dr.SolidFill CreateAccentSolidFill(int value)
        {
            int tintNumber = value / 6;
            int accentNumber = value % 6;

            var fill = new dr.SolidFill();
            var color = new dr.SchemeColor();
            var lumMod = new dr.LuminanceModulation();
            var lumOff = new dr.LuminanceOffset();

            switch (accentNumber)
            {
                case 0:
                    color.Val = dr.SchemeColorValues.Accent1;
                    break;
                case 1:
                    color.Val = dr.SchemeColorValues.Accent2;
                    break;
                case 2:
                    color.Val = dr.SchemeColorValues.Accent3;
                    break;
                case 3:
                    color.Val = dr.SchemeColorValues.Accent4;
                    break;
                case 4:
                    color.Val = dr.SchemeColorValues.Accent5;
                    break;
                case 5:
                    color.Val = dr.SchemeColorValues.Accent6;
                    break;
            }

            if (tintNumber > 4) tintNumber = 4;
            int modifier = tintNumber * 20 * 1000;
            lumMod.Val = 100000 - modifier;
            lumOff.Val = modifier;

            color.AppendChild(lumMod);
            color.AppendChild(lumOff);
            fill.AppendChild(color);

            return fill;
        }
    }
}
