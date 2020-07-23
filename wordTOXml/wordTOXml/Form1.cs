using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;


namespace wordTOXml
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string pathWord = "C:\\Users\\mishka\\Desktop\\UpWork\\wordTOXml\\wordTOXml\\bin\\Debug\\newword\\Before.docx";
            string savePath = "C:\\Users\\mishka\\Desktop\\UpWork\\wordTOXml\\wordTOXml\\bin\\Debug\\newword\\After.docx";

            Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = appWord.Documents.Open(pathWord);

            for (int i = 1; i <= doc.Shapes.Count; i++)
            {
                Shape objShape = doc.Shapes[i];

                if (objShape.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                {
                    ContentControls objContentControls = objShape.TextFrame.TextRange.ContentControls;

                    foreach (ContentControl objContentControl in objContentControls)
                    {

                        objShape.TextFrame.MarginTop = 0;
                        objShape.TextFrame.MarginBottom = 0;
                        objShape.TextFrame.MarginLeft = 0;
                        objShape.TextFrame.MarginRight = 0;
                        objShape.TextFrame.AutoSize = 0;

                        Microsoft.Office.Interop.Word.Table objTable = doc.Tables.Add(objShape.TextFrame.TextRange, 1, 2, WdDefaultTableBehavior.wdWord9TableBehavior, WdAutoFitBehavior.wdAutoFitContent);

                        objTable.TopPadding = 0;
                        objTable.BottomPadding = 0;
                        objTable.LeftPadding = 0;
                        objTable.RightPadding = 0;

                        objTable.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
                        objTable.Rows.SetHeight(objShape.Height, WdRowHeightRule.wdRowHeightExactly);
                        objTable.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
                        objTable.PreferredWidth = 100;

                        objTable.Cell(1, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        objTable.Cell(1, 1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop;

                        objTable.Cell(1, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                        objTable.Cell(1, 2).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom;

                    }
                }
            }

            doc.SaveAs2(savePath, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatXPS);

            doc.Close();
            appWord.Quit(false);
        }
        

        private void button2_Click(object sender, EventArgs e)
        {
            string pathWord = "C:\\Users\\mishka\\Desktop\\UpWork\\wordTOXml\\wordTOXml\\bin\\Debug\\newword\\Before.docx";
            string savePath = "C:\\Users\\mishka\\Desktop\\UpWork\\wordTOXml\\wordTOXml\\bin\\Debug\\newword\\After.docx";

            using (var mainDoc = WordprocessingDocument.Open(pathWord, false))
            using (var resultDoc = WordprocessingDocument.Create(savePath,
              WordprocessingDocumentType.Document))
            {
                // copy parts from source document to new document
                foreach (var part in mainDoc.Parts)
                    resultDoc.AddPart(part.OpenXmlPart, part.RelationshipId);

                IEnumerable<DocumentFormat.OpenXml.Vml.Shape> shapes =
                    resultDoc.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Vml.Shape>();

                UInt32 height = 0;

                foreach (var shape in shapes)
                {
                    // Getting shape height for adjusting table height
                    var val = shape.Style.Value;
                    string height_str = val.Split(new string[] { "height:" }, StringSplitOptions.None)[1]
                      .Split(new string[] { "pt;" }, StringSplitOptions.None)[0]
                      .Trim();
                    double h1 = double.Parse(height_str, System.Globalization.CultureInfo.InvariantCulture);
                    int h = Convert.ToInt32(h1);
                    height = Convert.ToUInt32(h);

                    string attributes = shape.Style.Value;
                    var map = attributes
                      .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Distinct()
                      .Select(x => x.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries))
                      .ToDictionary(p => p[0], p => p[1]);

                    map["margin-left"] = "0";
                    map["margin-right"] = "0";
                    map["margin-top"] = "0";
                    map["margin-bottom"] = "0";
                    //map["auto-size"] = "0";

                    //map["left-padding"] = "0";
                    //map["right-padding"] = "0";
                    //map["top-padding"] = "0";
                    //map["bottom-padding"] = "0";

                    shape.Style.Value =
                        string.Join(";", from pair in map
                                         select pair.Key + ":" + pair.Value);

                    shape.InsetMode = DocumentFormat.OpenXml.Vml.Office.InsetMarginValues.Custom;
                    OpenXmlElementList lst =  shape.ChildElements;
                    OpenXmlElementList textboxchilds = lst.First().ChildElements;

                    foreach (var cc in resultDoc.ContentControls())
                    {

                        DocumentFormat.OpenXml.Wordprocessing.Table table1 = new DocumentFormat.OpenXml.Wordprocessing.Table();

                        TableProperties tableProperties1 = new TableProperties();
                        DocumentFormat.OpenXml.Wordprocessing.TableStyle tableStyle1 = new DocumentFormat.OpenXml.Wordprocessing.TableStyle() { Val = "ad" };
                        BiDiVisual biDiVisual1 = new BiDiVisual();
                        TableWidth tableWidth1 = new TableWidth() { Width = "5200", Type = TableWidthUnitValues.Pct };
                        TableJustification tableJustification1 = new TableJustification() { Val = TableRowAlignmentValues.Center };

                        TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
                        TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 0, Type = TableWidthValues.Dxa };
                        TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 0, Type = TableWidthValues.Dxa };


                        tableCellMarginDefault1.Append(tableCellLeftMargin1);
                        tableCellMarginDefault1.Append(tableCellRightMargin1);
                        TableLook tableLook1 = new TableLook() { Val = "04A0" };


                        tableProperties1.Append(tableStyle1);
                        tableProperties1.Append(biDiVisual1);
                        tableProperties1.Append(tableWidth1);
                        tableProperties1.Append(tableJustification1);
                        tableProperties1.Append(tableCellMarginDefault1);
                        tableProperties1.Append(tableLook1);


                        TableGrid tableGrid1 = new TableGrid();
                        GridColumn gridColumn1 = new GridColumn() { Width = "555" };
                        GridColumn gridColumn2 = new GridColumn() { Width = "556" };


                        tableGrid1.Append(gridColumn1);
                        tableGrid1.Append(gridColumn2);

                        TableRow tableRow1 = new TableRow() { RsidTableRowAddition = "00AE49FD", RsidTableRowProperties = "00AE49FD" };

                        TableRowProperties tableRowProperties1 = new TableRowProperties();
                        TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = 200, HeightType = HeightRuleValues.Exact };

                        TableJustification tableJustification2 = new TableJustification() { Val = TableRowAlignmentValues.Center };

                        tableRowProperties1.Append(tableRowHeight1);
                        tableRowProperties1.Append(tableJustification2);

                        TableCell tableCell1 = new TableCell();

                        TableCellProperties tableCellProperties1 = new TableCellProperties();
                        TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "100", Type = TableWidthUnitValues.Pct };
                        //tableCell1.TableCellProperties.TableCellMargin.BottomMargin = new BottomMargin() { Width = "0" };
                        //tableCell1.TableCellProperties.TableCellMargin.TopMargin = new TopMargin() { Width = "0" };
                        TableCellVerticalAlignment tableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Top };
                        tableCellProperties1.Append(tableCellWidth1, tableCellVerticalAlignment);

                        DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph2 = new DocumentFormat.OpenXml.Wordprocessing.Paragraph() { RsidParagraphAddition = "00AE49FD", RsidParagraphProperties = "00AE49FD", RsidRunAdditionDefault = "00AE49FD" };

                        ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                        Justification justification1 = new Justification() { Val = JustificationValues.Right };

                        ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                        RightToLeftText rightToLeftText2 = new RightToLeftText();

                        paragraphMarkRunProperties2.Append(rightToLeftText2);

                        paragraphProperties2.Append(justification1);
                        paragraphProperties2.Append(paragraphMarkRunProperties2);

                        paragraph2.Append(paragraphProperties2);

                        tableCell1.Append(tableCellProperties1);
                        tableCell1.Append(paragraph2);

                        TableCell tableCell2 = new TableCell();

                        TableCellProperties tableCellProperties2 = new TableCellProperties();
                        TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "100", Type = TableWidthUnitValues.Pct };
                        //tableCell2.TableCellProperties.TableCellMargin.BottomMargin = new BottomMargin() { Width = "0" };
                        //tableCell2.TableCellProperties.TableCellMargin.TopMargin = new TopMargin() { Width = "0" };
                        TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Bottom };

                        tableCellProperties2.Append(tableCellWidth2);
                        tableCellProperties2.Append(tableCellVerticalAlignment1);

                        DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph3 = new DocumentFormat.OpenXml.Wordprocessing.Paragraph() { RsidParagraphAddition = "00AE49FD", RsidParagraphProperties = "00AE49FD", RsidRunAdditionDefault = "00AE49FD" };

                        ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                        Justification justification2 = new Justification() { Val = JustificationValues.Left };

                        ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
                        RightToLeftText rightToLeftText3 = new RightToLeftText();

                        paragraphMarkRunProperties3.Append(rightToLeftText3);

                        paragraphProperties3.Append(justification2);
                        paragraphProperties3.Append(paragraphMarkRunProperties3);

                        paragraph3.Append(paragraphProperties3);

                        tableCell2.Append(tableCellProperties2);
                        tableCell2.Append(paragraph3);

                        tableRow1.Append(tableRowProperties1);
                        tableRow1.Append(tableCell1);
                        tableRow1.Append(tableCell2);

                        table1.Append(tableProperties1);
                        table1.Append(tableGrid1);
                        table1.Append(tableRow1);

                        cc.InnerXml = "<w:tbl>" + table1.InnerXml + "</w:tbl>";
                        //cc.PrependChild(table1);                        //cc.Append(table1);
                    }
                }


            }


        }
    }
}
