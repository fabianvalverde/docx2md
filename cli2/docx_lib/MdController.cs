using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Text;
using draw = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;

namespace docx_lib
{
    public class MdController
    {
        private static int linksCount = 0;
        public static void ProcessParagraph(Paragraph block, StringBuilder textBuilder, MainDocumentPart mainPart, Dictionary<string, Stream> images, IEnumerable<HyperlinkRelationship> hyperlinks)
        {
            string constructorBase = "";
            bool isEsc = false; //is an escape url

            //iterate along every element in the Paragraphs and childrens
            foreach (var run in block.Descendants<Run>())
            {
                string prefix = "";
                var links = block.Descendants<Hyperlink>();
                if (run.InnerText != "")
                {
                    string[] escapeCharacters = new string[2];

                    foreach (var text in run)
                    {
                        if (text is Text)
                        {
                            escapeCharacters = ContainsEscape(text.InnerText);
                            if (isBlockQuote(block?.ParagraphProperties))
                            {
                                constructorBase += "\n";
                                constructorBase += ">" + text.InnerText;
                                constructorBase += "\n";
                                continue;
                            }
                            else
                            {
                                if (escapeCharacters[0] is not "")
                                {
                                    constructorBase += "" + text.InnerText.Replace(escapeCharacters[0], escapeCharacters[1]);
                                    isEsc = true;
                                }
                                else
                                {
                                    constructorBase += text.InnerText;
                                }
                                if (text.InnerText == "/")
                                {
                                    continue;
                                }
                            }
                        }

                        if (text is Break) { constructorBase += "\n"; continue; }
                        //checkbox
                        if (text.InnerText == "☐") { constructorBase = " [ ]"; continue; }
                        if (text.InnerText == "☒") { constructorBase = " [X]"; continue; }

                        /// Hyperlink
                        if (links.Count() > 0 && links.Count() > linksCount)
                        {
                            var LId = links.ElementAt(linksCount).Id;
                            var result = buildHyperLink(text, hyperlinks, LId, isEsc);
                            //is hyperlink
                            if (result != "")
                            {
                                constructorBase += result;
                                linksCount++;
                                break; //this break prevents duplication of hyperlink description
                            }
                        }
                        //code block
                        if (isCodeBlock(block?.ParagraphProperties))
                        {
                            constructorBase = "~~~~\n" + constructorBase + "\n~~~~\n";
                            continue;
                        }
                    }
                }

                //Images
                if (run.Descendants<Drawing>().Count() > 0)
                {
                    string description = run.Descendants<DocProperties>().First().Description ?? "";
                    string rId = run.Descendants<draw.Blip>().First().Embed.Value;
                    var imagePart = mainPart.GetPartById(rId);
                    string imageName = Path.GetFileName(imagePart.Uri.OriginalString);
                    if (Path.GetExtension(imageName).Equals(".bin"))
                    {
                        imageName = run.Descendants<DocProperties>().First().Name;
                    }

                    //This will help me to avoid new lines in the description.
                    if (description.Contains("/n"))
                    {
                        string[] substrings = description.Split('\n');
                        description = substrings[0];
                    }

                    constructorBase = "![" + description + "](" + "../images/" + imageName + ")";

                    MemoryStream imageStream = new MemoryStream();
                    imagePart.GetStream().CopyTo(imageStream);

                    if (!images.ContainsKey(imageName))
                    {
                        images.Add(imageName, imageStream);
                    }
                }
                //fonts, size letter, links
                if (run.RunProperties != null)
                {
                    prefix = ProcessBoldItalic(run);
                    constructorBase = prefix + constructorBase + prefix;
                }

                //general style, lists, aligment, spacing
                if (block.ParagraphProperties != null)
                {
                    prefix = ProcessParagraphFormats(block);

                    if (prefix == null) prefix = "";

                    if (prefix.Contains("#") || prefix.Contains("-"))
                    {
                        constructorBase = prefix + " " + constructorBase;
                    }

                    if (prefix.Contains(">"))
                    {
                        constructorBase = ProcessBlockQuote(run);
                    }

                }
                //linksCount = 0;
                textBuilder.Append(constructorBase);
                constructorBase = "";
            }
            linksCount = 0;
            constructorBase = textBuilder.ToString();
            textBuilder.Clear();

            textBuilder.Append(constructorBase);
            //if code block

            textBuilder.Append("\n");
        }

        public static void ProcessTable(Table node, StringBuilder textBuilder)
        {
            List<string> headerDivision = new List<string>();
            int rowNumber = 0;

            foreach (var row in node.Descendants<TableRow>())
            {
                rowNumber++;

                if (rowNumber == 2)
                {
                    headerDivider(headerDivision, textBuilder);
                }

                textBuilder.Append("| ");
                foreach (var cell in row.Descendants<TableCell>())
                {
                    foreach (var para in cell.Descendants<Paragraph>())
                    {
                        if (para.ParagraphProperties != null)
                        {
                            headerDivision.Add(para.ParagraphProperties.Justification.Val);
                        }
                        else
                        {
                            headerDivision.Add("normal");
                        }
                        textBuilder.Append(para.InnerText);
                    }
                    textBuilder.Append(" | ");
                }
                textBuilder.AppendLine("");
            }
        }

        private static string ProcessParagraphFormats(Paragraph block)
        {
            string style = block.ParagraphProperties?.ParagraphStyleId?.Val;

            if (style == null)
            {
                style = "single";
                block.ParagraphProperties.AppendChild(new ParagraphStyleId() { Val = "single" });
            }

            int num;
            string prefix = "";
            if ("top" == block.ParagraphProperties?.ParagraphBorders?.TopBorder?.LocalName
                && null == block.ParagraphProperties?.ParagraphBorders?.BottomBorder
                && null == block.ParagraphProperties?.ParagraphBorders?.LeftBorder)
            {
                prefix += "---\n";
                return prefix;
            }

            //to find Heading Paragraphs
            if (style.Contains("Heading"))
            {
                num = int.Parse(style.Substring(style.Length - 1));

                for (int i = 0; i < num; i++)
                {
                    prefix += "#";
                }
                return prefix;
            }

            //to find List Paragraphs
            if (style == "ListParagraph")
            {
                return prefix = "-";
            }

            //to find quotes Paragraphs
            if (style == "IntenseQuote")
            {
                return prefix = ">";
            }

            return prefix;
        }

        private static string ProcessBoldItalic(Run run)
        {
            //extract the child element of the text (Bold or Italic)
            OpenXmlElement expression = run.RunProperties.ChildElements.ElementAtOrDefault(0);

            string prefix = "";

            //to know if the propertie is Bold, Italic or both
            switch (expression)
            {
                case Bold:
                    if (run.RunProperties.ChildElements.Count == 2)
                    {
                        prefix = "***";
                        break;
                    }
                    prefix = "**";
                    break;
                case Italic:
                    prefix = "*";
                    break;
            }
            return prefix;
        }

        private static string ProcessBlockQuote(Run block)
        {
            string text = block.InnerText;
            string[] textSliced = text.Split("\n");
            string textBack = "";

            foreach (string n in textSliced)
            {
                textBack += "> " + n + "\n";
            }

            return textBack;
        }

        private static string[] ContainsEscape(string InnerText)
        {
            string[] result = new string[2];
            if (InnerText.Contains("#"))
            {
                result[0] = "#";
                result[1] = "\\#";
                return result;
            }
            else if (InnerText.Contains("-"))
            {
                result[0] = "#";
                result[1] = "\\#";
                return result;
            }
            else if (InnerText.Contains(">"))
            {
                result[0] = ">";
                result[1] = "\\>";
                return result;
            }
            else if (InnerText.Contains("["))
            {
                result[0] = "[";
                result[1] = "\\[";
                return result;
            }
            else if (InnerText.Contains("!["))
            {
                result[0] = "![";
                result[1] = "\\!\\[";
                return result;
            }
            else if (InnerText.Contains("*"))
            {
                result[0] = "![";
                result[1] = "\\!\\[";
                return result;
            }
            else
            {
                result[0] = "";
                result[1] = "";
                return result;
            }
        }

        private static string buildHyperLink(OpenXmlElement text, IEnumerable<HyperlinkRelationship> hyperlinks, string id = "", bool isEsc = false) //STRING LITERAL OR OPTIONAL
        {
            string cbt = "";
            if (text is RunProperties)
            {   //get to runStyles

                //var asd = text.Descendants<RunStyle>();
                foreach (RunStyle runStyle in text.Descendants<RunStyle>())
                {
                    //RunStyle runStyle = (RunStyle)text.FirstChild;
                    if (runStyle.Val == "Hyperlink")
                    {
                        if (isEsc)
                        {
                            cbt = hyperlinks.First(leenk => leenk.Id == id).Uri + "";

                        }
                        else
                        {
                            cbt = "[" + text.Parent.InnerText + "](" + hyperlinks.First(leenk => leenk.Id == id).Uri + ")";
                        }
                        return cbt;
                    }
                }
            }
            return "";
        }

        private static bool isBlockQuote(ParagraphProperties? Properties)
        {
            if (Properties == null) return false;
            // have 4 borderlines
            bool isLines = false;
            //shade
            bool isShading = false;
            //  indentation
            bool isIndentation = false;
            foreach (var style in Properties)
            {
                if (style is Shading) isShading = true;

                if (style is ParagraphBorders) isLines = true;

                if (style is Indentation) isIndentation = true;
            }
            return (isLines && isShading == false && isIndentation);
        }

        private static bool isCodeBlock(ParagraphProperties Properties)
        {
            if (Properties == null) return false;
            // have 4 borderlines
            bool isLines = false;
            //shade
            bool isShading = false;
            //  indentation
            bool isIndentation = false;

            foreach (var style in Properties)
            {
                if (style is Shading) isShading = true;

                if (style is ParagraphBorders) isLines = true;

                if (style is Indentation) isIndentation = true;
            }

            return (isLines && isShading && isIndentation);
        }

        private static void headerDivider(List<String> align, StringBuilder textBuilder)
        {
            textBuilder.Append("|");
            foreach (var column in align)
            {
                switch (column)
                {
                    case "left":
                        textBuilder.Append(":---|");
                        break;

                    case "center":
                        textBuilder.Append(":---:|");
                        break;

                    case "right":
                        textBuilder.Append("---:|");
                        break;

                    case "normal":
                        textBuilder.Append("---|");
                        break;
                }
            }
            textBuilder.AppendLine("");
        }
    }
}
