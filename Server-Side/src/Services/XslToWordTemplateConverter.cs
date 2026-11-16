using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;

namespace EJ2APIServices.Services
{
    /// <summary>
    /// Converts an XSL/XSLT stylesheet (stored as a .txt file) into a Word document
    /// whose merge fields are real DocIO MERGEFIELDs, so the result can be used as a
    /// mail-merge template.
    ///
    /// The mapping is:
    ///   <xsl:for-each select="X">  ->  «BeginGroup:X» ... «EndGroup:X»
    ///   <xsl:value-of select="Y"/> ->  MERGEFIELD Y
    ///
    /// The HTML/XHTML vocabulary of the stylesheet drives the Word formatting:
    /// h1..h6 become heading styles, p becomes a paragraph, ul/ol become lists,
    /// table/tr/td become a Word table, and the style="text-align:center"
    /// attribute is honoured.
    /// </summary>
    public class XslToWordTemplateConverter
    {
        private static readonly XNamespace XslNs = "http://www.w3.org/1999/XSL/Transform";

        /// <summary>
        /// Builds a WordDocument from the XSL stylesheet at <paramref name="xslFilePath"/>.
        /// The CALLER owns the returned document and must Close() it.
        /// </summary>
        public WordDocument ConvertToWordDocument(string xslFilePath)
        {
            string xslContent = File.ReadAllText(xslFilePath);
            XDocument xDoc = XDocument.Parse(xslContent);

            WordDocument document = new WordDocument();
            IWSection section = document.AddSection();
            section.PageSetup.Margins.All = 72f;

            // Find the xsl:template elements; fall back to the document root when the
            // stylesheet has none.
            var templates = xDoc.Descendants(XslNs + "template").ToList();
            if (templates.Count == 0)
            {
                if (xDoc.Root != null)
                {
                    ProcessNodes(xDoc.Root.Nodes(), section, null);
                }
            }
            else
            {
                foreach (var template in templates)
                {
                    ProcessNodes(template.Nodes(), section, null);
                }
            }

            return document;
        }

        private void ProcessNodes(IEnumerable<XNode> nodes, IWSection section, IWTable? currentTable)
        {
            foreach (var node in nodes)
            {
                if (node is XElement element)
                {
                    if (element.Name.Namespace == XslNs)
                    {
                        ProcessXslElement(element, section, currentTable);
                        continue;
                    }

                    string localName = element.Name.LocalName.ToLowerInvariant();

                    switch (localName)
                    {
                        case "html":
                        case "body":
                        case "div":
                        case "section":
                            ProcessNodes(element.Nodes(), section, currentTable);
                            break;

                        case "h1":
                        case "h2":
                        case "h3":
                        case "h4":
                        case "h5":
                        case "h6":
                            var headingPara = section.AddParagraph();
                            ApplyHeadingStyle(headingPara, localName);
                            ApplyParagraphStyles(element, headingPara);
                            RenderInlineContent(element, headingPara);
                            break;

                        case "p":
                            var para = section.AddParagraph();
                            ApplyParagraphStyles(element, para);
                            RenderInlineContent(element, para);
                            break;

                        case "ul":
                            foreach (var li in element.Elements().Where(e => e.Name.LocalName.Equals("li", StringComparison.OrdinalIgnoreCase)))
                            {
                                var bulletPara = section.AddParagraph();
                                bulletPara.ListFormat.ApplyDefBulletStyle();
                                ApplyParagraphStyles(li, bulletPara);
                                RenderInlineContent(li, bulletPara);
                            }
                            break;

                        case "ol":
                            foreach (var li in element.Elements().Where(e => e.Name.LocalName.Equals("li", StringComparison.OrdinalIgnoreCase)))
                            {
                                var numPara = section.AddParagraph();
                                numPara.ListFormat.ApplyDefNumberedStyle();
                                ApplyParagraphStyles(li, numPara);
                                RenderInlineContent(li, numPara);
                            }
                            break;

                        case "table":
                            ConvertTable(element, section);
                            break;

                        default:
                            if (element.HasElements)
                            {
                                ProcessNodes(element.Nodes(), section, currentTable);
                            }
                            else if (!string.IsNullOrWhiteSpace(element.Value))
                            {
                                var textPara = section.AddParagraph();
                                ApplyParagraphStyles(element, textPara);
                                RenderInlineContent(element, textPara);
                            }
                            break;
                    }
                }
            }
        }

        private void ProcessXslElement(XElement element, IWSection section, IWTable? currentTable)
        {
            string localName = element.Name.LocalName;

            switch (localName)
            {
                case "for-each":
                    string select = element.Attribute("select")?.Value ?? "Group";
                    string groupName = NormalizeGroupName(select);
                    var beginPara = section.AddParagraph();
                    beginPara.AppendField($"BeginGroup:{groupName}", FieldType.FieldMergeField);

                    ProcessNodes(element.Nodes(), section, currentTable);

                    var endPara = section.AddParagraph();
                    endPara.AppendField($"EndGroup:{groupName}", FieldType.FieldMergeField);
                    break;

                case "if":
                case "when":
                case "otherwise":
                case "choose":
                    ProcessNodes(element.Nodes(), section, currentTable);
                    break;

                case "value-of":
                    var valPara = section.AddParagraph();
                    string fieldSelect = element.Attribute("select")?.Value ?? "Field";
                    valPara.AppendField(NormalizeFieldName(fieldSelect), FieldType.FieldMergeField);
                    break;

                case "apply-templates":
                    break;

                default:
                    ProcessNodes(element.Nodes(), section, currentTable);
                    break;
            }
        }

        private void ConvertTable(XElement tableElement, IWSection section)
        {
            IWTable table = section.AddTable();
            table.TableFormat.Borders.BorderType = BorderStyle.Single;
            table.TableFormat.Borders.LineWidth = 1f;
            table.TableFormat.Borders.Color = Color.LightGray;

            // Parse borders attribute if present
            var borderAttr = tableElement.Attribute("border");
            if (borderAttr != null && int.TryParse(borderAttr.Value, out int b) && b > 0)
            {
                table.TableFormat.Borders.BorderType = BorderStyle.Single;
                table.TableFormat.Borders.LineWidth = 1f;
                table.TableFormat.Borders.Color = Color.Black;
            }

            // Collect rows and repeating regions
            ProcessTableChildren(tableElement, table);
        }

        private void ProcessTableChildren(XElement container, IWTable table)
        {
            foreach (var child in container.Elements())
            {
                if (child.Name.Namespace == XslNs)
                {
                    if (child.Name.LocalName == "for-each")
                    {
                        string select = child.Attribute("select")?.Value ?? "Items";
                        string groupName = NormalizeGroupName(select);

                        // Collect row templates inside for-each (might be nested in xsl:if, xsl:choose, etc.)
                        var rowsInForEach = FindElementsDeep(child, "tr");
                        bool isFirstRow = true;
                        int rowCount = rowsInForEach.Count;

                        for (int i = 0; i < rowCount; i++)
                        {
                            var trElement = rowsInForEach[i];
                            bool isLastRow = (i == rowCount - 1);
                            AddTableRow(table, trElement, isHeader: false, groupName: groupName, isGroupStart: isFirstRow, isGroupEnd: isLastRow);
                            isFirstRow = false;
                        }
                    }
                    else if (child.Name.LocalName is "if" or "choose" or "when" or "otherwise")
                    {
                        ProcessTableChildren(child, table);
                    }
                }
                else if (child.Name.LocalName.Equals("tr", StringComparison.OrdinalIgnoreCase))
                {
                    bool isHeader = child.Elements().Any(e => e.Name.LocalName.Equals("th", StringComparison.OrdinalIgnoreCase));
                    AddTableRow(table, child, isHeader: isHeader);
                }
                else if (child.Name.LocalName is "tbody" or "thead" or "tfoot")
                {
                    ProcessTableChildren(child, table);
                }
            }
        }

        private List<XElement> FindElementsDeep(XElement root, string localName)
        {
            var result = new List<XElement>();
            foreach (var el in root.Elements())
            {
                if (el.Name.LocalName.Equals(localName, StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(el);
                }
                else
                {
                    result.AddRange(FindElementsDeep(el, localName));
                }
            }
            return result;
        }

        private void AddTableRow(IWTable table, XElement trElement, bool isHeader, string? groupName = null, bool isGroupStart = false, bool isGroupEnd = false)
        {
            var row = table.Rows.Count == 0
                ? table.AddRow(isHeader)
                : table.AddRow(isCopyFormat: false, autoPopulateCells: false);
            row.IsHeader = isHeader;

            // Background color for row (if specified on tr)
            Color? bgColor = ParseColor(trElement.Attribute("bgcolor")?.Value ?? trElement.Attribute("style")?.Value);
            if (bgColor.HasValue)
            {
                row.RowFormat.BackColor = bgColor.Value;
            }
            else
            {
                row.RowFormat.BackColor = Color.Empty;
            }

            var cellElements = trElement.Elements().Where(e =>
                e.Name.LocalName.Equals("th", StringComparison.OrdinalIgnoreCase) ||
                e.Name.LocalName.Equals("td", StringComparison.OrdinalIgnoreCase)).ToList();

            for (int i = 0; i < cellElements.Count; i++)
            {
                var cellElement = cellElements[i];
                var cell = row.AddCell();

                Color? cellBgColor = ParseColor(cellElement.Attribute("bgcolor")?.Value ?? cellElement.Attribute("style")?.Value);
                if (cellBgColor.HasValue)
                {
                    cell.CellFormat.BackColor = cellBgColor.Value;
                }
                else if (bgColor.HasValue)
                {
                    cell.CellFormat.BackColor = bgColor.Value;
                }
                else
                {
                    cell.CellFormat.BackColor = Color.Empty;
                }

                var para = cell.AddParagraph();
                if (isHeader || cellElement.Name.LocalName.Equals("th", StringComparison.OrdinalIgnoreCase))
                {
                    para.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                }
                ApplyParagraphStyles(cellElement, para);

                // Insert BeginGroup in first cell of first repeating row
                if (isGroupStart && i == 0 && !string.IsNullOrEmpty(groupName))
                {
                    para.AppendField($"BeginGroup:{groupName}", FieldType.FieldMergeField);
                }

                RenderCellContent(cellElement, para, isHeader || cellElement.Name.LocalName.Equals("th", StringComparison.OrdinalIgnoreCase));

                // Insert EndGroup in last cell of last repeating row
                if (isGroupEnd && i == cellElements.Count - 1 && !string.IsNullOrEmpty(groupName))
                {
                    para.AppendField($"EndGroup:{groupName}", FieldType.FieldMergeField);
                }
            }
        }

        private void RenderCellContent(XElement cellElement, IWParagraph para, bool isBold)
        {
            foreach (var node in cellElement.Nodes())
            {
                if (node is XText text)
                {
                    if (!string.IsNullOrWhiteSpace(text.Value))
                    {
                        var textRange = para.AppendText(text.Value.Trim());
                        if (isBold) textRange.CharacterFormat.Bold = true;
                    }
                }
                else if (node is XElement child)
                {
                    if (child.Name.Namespace == XslNs)
                    {
                        if (child.Name.LocalName == "value-of")
                        {
                            string select = child.Attribute("select")?.Value ?? "Field";
                            string fieldName = NormalizeFieldName(select);
                            para.AppendField(fieldName, FieldType.FieldMergeField);
                        }
                        else
                        {
                            RenderCellContent(child, para, isBold);
                        }
                    }
                    else
                    {
                        bool childBold = isBold || child.Name.LocalName.Equals("b", StringComparison.OrdinalIgnoreCase) || child.Name.LocalName.Equals("strong", StringComparison.OrdinalIgnoreCase);
                        RenderCellContent(child, para, childBold);
                    }
                }
            }
        }

        private void RenderInlineContent(XElement container, IWParagraph para)
        {
            foreach (var node in container.Nodes())
            {
                if (node is XText text)
                {
                    string val = text.Value;
                    if (!string.IsNullOrEmpty(val))
                    {
                        para.AppendText(val.Trim());
                    }
                }
                else if (node is XElement child)
                {
                    if (child.Name.Namespace == XslNs)
                    {
                        if (child.Name.LocalName == "value-of")
                        {
                            string select = child.Attribute("select")?.Value ?? "Field";
                            para.AppendField(NormalizeFieldName(select), FieldType.FieldMergeField);
                        }
                        else
                        {
                            RenderInlineContent(child, para);
                        }
                    }
                    else
                    {
                        bool isBold = child.Name.LocalName.Equals("b", StringComparison.OrdinalIgnoreCase) || child.Name.LocalName.Equals("strong", StringComparison.OrdinalIgnoreCase);
                        bool isItalic = child.Name.LocalName.Equals("i", StringComparison.OrdinalIgnoreCase) || child.Name.LocalName.Equals("em", StringComparison.OrdinalIgnoreCase);

                        string childText = child.Value;
                        if (!string.IsNullOrEmpty(childText))
                        {
                            var tr = para.AppendText(childText.Trim());
                            if (isBold) tr.CharacterFormat.Bold = true;
                            if (isItalic) tr.CharacterFormat.Italic = true;
                        }
                    }
                }
            }
        }

        private void ApplyHeadingStyle(IWParagraph para, string headingTag)
        {
            switch (headingTag.ToLowerInvariant())
            {
                case "h1":
                    para.ApplyStyle(BuiltinStyle.Heading1);
                    break;
                case "h2":
                    para.ApplyStyle(BuiltinStyle.Heading2);
                    var h2Style = para.Document.Styles.FindByName("Heading 2") as IWParagraphStyle;
                    if (h2Style != null)
                    {
                        h2Style.CharacterFormat.Bold = true;
                        h2Style.CharacterFormat.Italic = false;
                        h2Style.CharacterFormat.ItalicBidi = false;
                    }
                    break;
                case "h3":
                    para.ApplyStyle(BuiltinStyle.Heading3);
                    break;
                case "h4":
                    para.ApplyStyle(BuiltinStyle.Heading4);
                    break;
                default:
                    para.ApplyStyle(BuiltinStyle.Heading1);
                    break;
            }
            para.ParagraphFormat.AfterSpacing = 6f;
            para.ParagraphFormat.BeforeSpacing = 12f;
        }

        private void ApplyParagraphStyles(XElement element, IWParagraph para)
        {
            // Check align attribute
            string? alignAttr = element.Attribute("align")?.Value;
            if (!string.IsNullOrWhiteSpace(alignAttr))
            {
                ApplyAlignment(alignAttr.Trim(), para);
            }

            // Check style attribute (e.g. style="text-align:center;")
            string? styleAttr = element.Attribute("style")?.Value;
            if (!string.IsNullOrWhiteSpace(styleAttr))
            {
                var declarations = styleAttr.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                foreach (var decl in declarations)
                {
                    var parts = decl.Split(':', 2);
                    if (parts.Length == 2)
                    {
                        string prop = parts[0].Trim().ToLowerInvariant();
                        string val = parts[1].Trim().ToLowerInvariant();

                        if (prop == "text-align")
                        {
                            ApplyAlignment(val, para);
                        }
                    }
                }
            }
        }

        private void ApplyAlignment(string alignValue, IWParagraph para)
        {
            switch (alignValue.ToLowerInvariant())
            {
                case "center":
                    para.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                    break;
                case "right":
                    para.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                    break;
                case "left":
                    para.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                    break;
                case "justify":
                    para.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Justify;
                    break;
            }
        }

        private string NormalizeGroupName(string xpath)
        {
            string cleaned = xpath.Trim();
            cleaned = Regex.Replace(cleaned, @"[^a-zA-Z0-9_:]+", "_").Trim('_');
            return string.IsNullOrEmpty(cleaned) ? "Group" : cleaned;
        }

        private string NormalizeFieldName(string xpath)
        {
            string cleaned = xpath.Trim();
            if (cleaned.StartsWith("@")) cleaned = cleaned.Substring(1);
            int lastSlash = cleaned.LastIndexOfAny(new[] { '/', '\\' });
            if (lastSlash >= 0)
            {
                cleaned = cleaned.Substring(lastSlash + 1);
            }
            cleaned = Regex.Replace(cleaned, @"[^a-zA-Z0-9_:]+", "_").Trim('_');
            return string.IsNullOrEmpty(cleaned) ? "Field" : cleaned;
        }

        private Color? ParseColor(string? colorStr)
        {
            if (string.IsNullOrWhiteSpace(colorStr)) return null;

            colorStr = colorStr.Trim();
            var hexMatch = Regex.Match(colorStr, @"#([0-9a-fA-F]{6}|[0-9a-fA-F]{3})");
            if (hexMatch.Success)
            {
                string hex = hexMatch.Groups[1].Value;
                if (hex.Length == 3)
                {
                    hex = $"{hex[0]}{hex[0]}{hex[1]}{hex[1]}{hex[2]}{hex[2]}";
                }
                int r = System.Convert.ToInt32(hex.Substring(0, 2), 16);
                int g = System.Convert.ToInt32(hex.Substring(2, 2), 16);
                int b = System.Convert.ToInt32(hex.Substring(4, 2), 16);
                return Color.FromArgb(r, g, b);
            }

            try
            {
                var prop = typeof(Color).GetProperty(colorStr, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.IgnoreCase);
                if (prop != null)
                {
                    return (Color)prop.GetValue(null)!;
                }
            }
            catch { }

            return null;
        }
    }
}
