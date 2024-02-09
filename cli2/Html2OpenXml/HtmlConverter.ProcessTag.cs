/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml - All Rights Reserved
 * 
 * This source is subject to the Microsoft Permissive License.
 * Please see the License.txt file for more information.
 * All other rights reserved.
 * 
 * THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY 
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
 * PARTICULAR PURPOSE.
 */
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml.IO;
using w14 = DocumentFormat.OpenXml.Office2010.Word;

namespace HtmlToOpenXml
{
    using a = DocumentFormat.OpenXml.Drawing;
    using pic = DocumentFormat.OpenXml.Drawing.Pictures;
    using wBorder = DocumentFormat.OpenXml.Wordprocessing.Border;

    partial class HtmlConverter
    {
        //____________________________________________________________________
        //
        // Processing known tags

        #region ProcessAcronym

        private void ProcessAcronym(HtmlEnumerator en)
        {
            // Transform the inline acronym/abbreviation to a reference to a foot note.

            string title = en.Attributes["title"];
            if (title == null) return;

            AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);

            if (elements.Count > 0 && elements[0] is Run)
            {
                string runStyle;
                FootnoteEndnoteReferenceType reference;

                if (this.AcronymPosition == AcronymPosition.PageEnd)
                {
                    reference = new FootnoteReference() { Id = AddFootnoteReference(title) };
                    runStyle = htmlStyles.DefaultStyles.FootnoteReferenceStyle;
                }
                else
                {
                    reference = new EndnoteReference() { Id = AddEndnoteReference(title) };
                    runStyle = htmlStyles.DefaultStyles.EndnoteReferenceStyle;
                }

                Run run;
                elements.Add(
                    run = new Run(
                        new RunProperties
                        {
                            RunStyle = new RunStyle() { Val = htmlStyles.GetStyle(runStyle, StyleValues.Character) }
                        },
                        reference));
            }
        }

        #endregion

        #region ProcessBlockQuote

        private void ProcessBlockQuote(HtmlEnumerator en)
        {
            //CompleteCurrentParagraph(true);

            //currentParagraph = htmlStyles.Paragraph.NewParagraph();

            // Save the new paragraph reference to support nested numbering list.

            AlternateProcessHtmlChunks(en, "</blockquote>");

            //CompleteCurrentParagraph(true);

        }

        #endregion

        #region ProcessBody

        private void ProcessBody(HtmlEnumerator en)
        {
            List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
            htmlStyles.Paragraph.ProcessCommonAttributes(en, styleAttributes);

            if (styleAttributes.Count > 0)
                htmlStyles.Runs.BeginTag(en.CurrentTag, styleAttributes.ToArray());

            // Unsupported W3C attribute but claimed by users. Specified at <body> level, the page
            // orientation is applied on the whole document
            string attr = en.StyleAttributes["page-orientation"];
            if (attr != null)
            {
                PageOrientationValues orientation = Converter.ToPageOrientation(attr);

                SectionProperties sectionProperties = mainPart.Document.Body.GetFirstChild<SectionProperties>();
                if (sectionProperties == null || sectionProperties.GetFirstChild<PageSize>() == null)
                {
                    mainPart.Document.Body.Append(HtmlConverter.ChangePageOrientation(orientation));
                }
                else
                {
                    PageSize pageSize = sectionProperties.GetFirstChild<PageSize>();
                    if (!pageSize.Compare(orientation))
                    {
                        SectionProperties validSectionProp = ChangePageOrientation(orientation);
                        if (pageSize != null) pageSize.Remove();
                        sectionProperties.PrependChild(validSectionProp.GetFirstChild<PageSize>().CloneNode(true));
                    }
                }
            }
        }

        #endregion

        #region ProcessBr

        private void ProcessBr(HtmlEnumerator en)
        {
            elements.Add(new Run(new Break()));
        }

        #endregion

        #region ProcessCite

        private void ProcessCite(HtmlEnumerator en)
        {
            ProcessHtmlElement<RunStyle>(en, new RunStyle() { Val = htmlStyles.GetStyle(htmlStyles.DefaultStyles.QuoteStyle, StyleValues.Character) });
        }

        #endregion

        #region ProcessDefinitionList

        private void ProcessDefinitionList(HtmlEnumerator en)
        {
            ProcessParagraph(en);
            currentParagraph.InsertInProperties(prop => prop.SpacingBetweenLines = new SpacingBetweenLines() { After = "0" });
        }

        #endregion

        #region ProcessDefinitionListItem

        private void ProcessDefinitionListItem(HtmlEnumerator en)
        {
            AlternateProcessHtmlChunks(en, "</dd>");

            currentParagraph = htmlStyles.Paragraph.NewParagraph();
            currentParagraph.Append(elements);
            currentParagraph.InsertInProperties(prop =>
            {
                prop.Indentation = new Indentation() { FirstLine = "708" };
                prop.SpacingBetweenLines = new SpacingBetweenLines() { After = "0" };
            });

            // Restore the original elements list
            AddParagraph(currentParagraph);
            this.elements.Clear();
        }

        #endregion

        #region ProcessDiv

        private void ProcessDiv(HtmlEnumerator en)
        {
            // The way the browser consider <div> is like a simple Break. But in case of any attributes that targets
            // the paragraph, we don't want to apply the style on the old paragraph but on a new one.
            if (en.Attributes.Count == 0 || (en.StyleAttributes["text-align"] == null && en.Attributes["align"] == null && en.StyleAttributes.GetAsBorder("border").IsEmpty))
            {
                List<OpenXmlElement> runStyleAttributes = new List<OpenXmlElement>();
                bool newParagraph = ProcessContainerAttributes(en, runStyleAttributes);
                CompleteCurrentParagraph(newParagraph);

                if (runStyleAttributes.Count > 0)
                    htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes);

                // Any changes that requires a new paragraph?
                if (newParagraph)
                {
                    // Insert before the break, complete this paragraph and start a new one
                    this.paragraphs.Insert(this.paragraphs.Count - 1, currentParagraph);
                    AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);
                    CompleteCurrentParagraph();
                }
            }
            else
            {
                // treat div as a paragraph
                ProcessParagraph(en);
            }
        }

        #endregion

        #region ProcessFont

        private void ProcessFont(HtmlEnumerator en)
        {
            List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
            ProcessContainerAttributes(en, styleAttributes);

            string attrValue = en.Attributes["size"];
            if (attrValue != null)
            {
                Unit fontSize = Converter.ToFontSize(attrValue);
                if (fontSize.IsFixed)
                    styleAttributes.Add(new FontSize { Val = (fontSize.ValueInPoint * 2).ToString(CultureInfo.InvariantCulture) });
            }

            attrValue = en.Attributes["face"];
            if (attrValue != null)
            {
                // Set HightAnsi. Bug fixed by xjpmauricio on github.com/onizet/html2openxml/discussions/285439
                // where characters with accents were always using fallback font
                styleAttributes.Add(new RunFonts { Ascii = attrValue, HighAnsi = attrValue });
            }

            if (styleAttributes.Count > 0)
                htmlStyles.Runs.MergeTag(en.CurrentTag, styleAttributes);
        }

        #endregion

        #region ProcessHeading

        private void ProcessHeading(HtmlEnumerator en)
        {
            char level = en.Current[2];

            // support also style attributes for heading (in case of css override)
            List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
            htmlStyles.Paragraph.ProcessCommonAttributes(en, styleAttributes);

            AlternateProcessHtmlChunks(en, "</h" + level + ">");

            Paragraph p = new Paragraph(elements);
            p.InsertInProperties(prop =>
                prop.ParagraphStyleId = new ParagraphStyleId() { Val = htmlStyles.GetStyle(htmlStyles.DefaultStyles.HeadingStyle + level, StyleValues.Paragraph) });

            // Check if the line starts with a number format (1., 1.1., 1.1.1.)
            // If it does, make sure we make the heading a numbered item
            OpenXmlElement firstElement = elements.First();
            Match regexMatch = Regex.Match(firstElement.InnerText, @"(?m)^(\d+\.)*\s");

            // Make sure we only grab the heading if it starts with a number
            if (regexMatch.Groups.Count > 1 && regexMatch.Groups[1].Captures.Count > 0)
            {
                int indentLevel = regexMatch.Groups[1].Captures.Count;

                // Strip numbers from text
                firstElement.InnerXml = firstElement.InnerXml.Replace(firstElement.InnerText, firstElement.InnerText.Substring(indentLevel * 2 + 1)); // number, dot and whitespace

                htmlStyles.NumberingList.ApplyNumberingToHeadingParagraph(p, indentLevel);
            }

            htmlStyles.Paragraph.ApplyTags(p);
            htmlStyles.Paragraph.EndTag("<h" + level + ">");

            this.elements.Clear();
            AddParagraph(p);
            AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
        }

        #endregion

        #region ProcessHorizontalLine

        private void ProcessHorizontalLine(HtmlEnumerator en)
        {
            // Insert an horizontal line as it stands in many emails.
            CompleteCurrentParagraph(true);

            // If the previous paragraph contains a bottom border or is a Table, we add some spacing between the <hr>
            // and the previous element or Word will display only the last border.
            // (see Remarks: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.bottomborder%28office.14%29.aspx)
            if (paragraphs.Count >= 2)
            {
                OpenXmlCompositeElement previousElement = paragraphs[paragraphs.Count - 2];
                bool addSpacing = false;
                ParagraphProperties prop = previousElement.GetFirstChild<ParagraphProperties>();
                if (prop != null)
                {
                    if (prop.ParagraphBorders != null && prop.ParagraphBorders.BottomBorder != null
                        && prop.ParagraphBorders.BottomBorder.Size > 0U)
                        addSpacing = true;
                }
                else
                {
                    if (previousElement is Table)
                        addSpacing = true;
                }

                if (addSpacing)
                {
                    currentParagraph.InsertInProperties(p => p.SpacingBetweenLines = new SpacingBetweenLines() { Before = "240" });
                }
            }

            // if this paragraph has no children, it will be deleted in RemoveEmptyParagraphs()
            // in order to kept the <hr>, we force an empty run
            currentParagraph.Append(new Run());

            // Get style from border (only top) or use Default style 
            TopBorder hrBorderStyle = null;

            var border = en.StyleAttributes.GetAsBorder("border");
            if (!border.IsEmpty && border.Top.IsValid)
                hrBorderStyle = new TopBorder { Val = border.Top.Style, Color = StringValue.FromString(border.Top.Color.ToHexString()), Size = (uint)border.Top.Width.Value };
            else
                hrBorderStyle = new TopBorder() { Val = BorderValues.Single, Size = 4U };

            currentParagraph.InsertInProperties(prop =>
            prop.ParagraphBorders = new ParagraphBorders
            {
                TopBorder = hrBorderStyle
            });
        }

        #endregion

        #region ProcessHtml

        private void ProcessHtml(HtmlEnumerator en)
        {
            List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
            htmlStyles.Paragraph.ProcessCommonAttributes(en, styleAttributes);

            if (styleAttributes.Count > 0)
                htmlStyles.Runs.BeginTag(en.CurrentTag, styleAttributes.ToArray());
        }

        #endregion

        #region ProcessHtmlElement

        private void ProcessHtmlElement<T>(HtmlEnumerator en) where T : OpenXmlLeafElement, new()
        {
            ProcessHtmlElement<T>(en, new T());
        }

        /// <summary>
        /// Generic handler for processing style on any Html element.
        /// </summary>
        private void ProcessHtmlElement<T>(HtmlEnumerator en, OpenXmlLeafElement style) where T : OpenXmlLeafElement
        {
            List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>() { style };
            ProcessContainerAttributes(en, styleAttributes);
            htmlStyles.Runs.MergeTag(en.CurrentTag, styleAttributes);
        }

        #endregion

        #region ProcessFigureCaption

        private void ProcessFigureCaption(HtmlEnumerator en)
        {
            this.CompleteCurrentParagraph(true);

            currentParagraph.Append(
                    new ParagraphProperties
                    {
                        ParagraphStyleId = new ParagraphStyleId() { Val = htmlStyles.GetStyle(htmlStyles.DefaultStyles.CaptionStyle, StyleValues.Paragraph) },
                        KeepNext = new KeepNext()
                    },
                    new Run(
                        new Text("Figure ") { Space = SpaceProcessingModeValues.Preserve }
                    ),
                    new SimpleField(
                        new Run(
                            new Text(AddFigureCaption().ToString(CultureInfo.InvariantCulture)))
                    )
                    { Instruction = " SEQ Figure \\* ARABIC " }
                );

            ProcessHtmlChunks(en, "</figcaption>");

            if (elements.Count > 0) // any caption?
            {
                Text t = (elements[0] as Run).GetFirstChild<Text>();
                t.Text = " " + t.InnerText; // append a space after the numero of the picture
            }

            this.CompleteCurrentParagraph(true);
        }

        #endregion

        #region ProcessImage

        private void ProcessImage(HtmlEnumerator en)
        {

            String defaultImage = "ffd8ffe000104a46494600010101012c012c0000ffdb004300080606070605080707070909080a0c140d0c0b0b0c1912130f141d1a1f1e1d1a1c1c20242e2720222c231c1c2837292c30313434341f27393d38323c2e333432ffc2000b080320032001012200ffc4001b00010002030101000000000000000000000006070304050201ffda00080101000000019f8000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000003cc50000000000000129c8000000000031d3bbf90000000000001e39f6f6c00000000000c74eda3d500000000000073aaab7b60000000000063a76d1ea80000000000039d555bdb0000000000031d3b68f54000000000001ceaaaded8000000000018e9db47aa000000000000e75556f6c00000000000c74eda3d500000000000073aaab7b60000000000063a76d1ea80000000000039d555bdb0000000000031d3b68f54000000000001ceaaaded8000000000018e9db47aa000d18f69ed77fa4000000073aaab7b60000000000063a76d1ea80042e1fbfd0d2e64a277e800000039d555bdb0000000000031d3b68f54004421b60c80e3577239d000000073aaab7b60000000000063a76d1ea80186a49d49c3875b5afbc00000039d555bdb0000000000031d3b68f54008fd7d6ffa054f2d958000000e75556f6c00000000000c74eda3d50022f0eb640ac7b932001afb000073aaab7b60000000000063a76d1ea801c4adadecc1e6a49c49400d3acad1d800073aaab7b60000000000063a76d1ea801e6a6924dc2270ab6b38015c7024b6080039d555bdb0000000000031d3b68f5400706b995cc3670c4e1b3d938011faee7703b23ba001ceaaaded8000000000018e9db47aa001c281e9e7c3b138920018aaa91cde0bc0b532800e75556f6c00000000000c74eda3d50007ce569ed75bd00084472d5cb8aabefce80073aaab7b60000000000063a76d1ea800000039b56d89203855bd9dd70039d555bdb0000000000031d3b68f54000000158e7b1c2bee5da7e801ceaaaded8000000000018e9db47aa0f38f30000045e0d6aee06bd552a998039d555bdb0000000000031d3b68f5410de5d8c000035ea997cbc08e57f69744039d555bdb0000000000031d3b68f5439b577c9ccac00008072ad2f402b8d7b3be80e75556f6c00000000000c74eda3d51f2afdbef412d0e9800071ab3b37b001a9554d258039d555bdb0000000000031d3b68f544421f6aed579ccb4728001e6aceb4fc008ac22d4dd039d555bdb0000000000031d3b68f54e7d5b3b93b0d5bd8b0000087c46d6d8003e56596c9039d555bdb0000000000031d3b68f55f2b2cb648e55633d93000695553b9380073eac9f4941ceaaaded8000000000018e9db47aa89c2ed4dc088c32d2df0015c60b380008744ad5d90e75556f6c00000000000c74eda3d5d2aae6f2a02b7d5b3fd8023f5dda5d20003cd5dd0b0c39d555bdb0000000000031d3b68f56b4f96600d7ab2413b00c55548e6e0001caac2c5ef8e75556f6c00000000000c74eda3cb835a9bc01c5ad6c4900083c7ad5c8000083c72d5cc73aaab7b60000000000063a76c5804ce5a0042e2969ee03995758bdf000031d57dc9e1ceaaaded8000000000018e9ddddab3be801f2b3f966fd1f2b1d9b1800003895ad99d973aaab7b60000000000063a77dda7d0000d4ab25335117835abb803e79fbe80080f1ed3f7ceaaaded8000000000018e9d9acc4000e057564f6daf544c25e018cf7f40182aa93cd79d555bdb000000000003153d6875807cf3ec1058eda7b35ff2ed2f401e3cbdfa0023d5e5a3eaaab7b60000000000062a7ad0eb01f3e78cbf3ebe7df1586e4b2b4b33b07d3e797a79fbf47d7c7d0aeb4e7d565bdb000000000003153d68f541f3c7c7bf3ebc9ef9f5afceecec7bfb8c7cf7e7d797df9e9e5efe8d5aabbf19b7b60000000000061a7ed1ea83c3d63f4f5e1e9e23f0eb332fbf3f5ebc7bf2faf1ebcfbf1e9f1f7d822f02f96f6c00000000000c34fda3d5079f2fa7bc6fbf7cfde7747cfbf3f7e7df993cfcf9ebc7af3931fa3e3d63cc15a712ded80000000000182a0b47aa079f9efe3efcf2f4f9f3d7cfbf0f9ebebcbd79fbf3efc3d79f9ebc640e4d5f6eec80000000000c1505a3d500000000000073aaab776400000000006bd4360f48000000000001a15ddbbb2000000000035ea100000000000005bbb200000000001a80000000000000dafa000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000c11fcf2039fc8ec6f8007075a4be8000000000000000000e7555f6cceca2d029fc9c002b0e45c394000000000000000000e7554e8da7ea2d029fc9c0038d83bff004000000000000000000e75544da6116814fe4e8f4474b6e5925040f9d64e940e45c9e2f427712e0ee4f3a6884675f624f2e6b41391bf278bcaa478a11c1f7219a7a0000000000e7555dde479b5787029fc9e375eedf7789a53c950561c8b879b587a92e08ffbedecc63af67c76bbecf6637a36875ab0e448b66318e712eae38126f91a95cec00000000039d5549fa904eff76073f93d53a36b6f6b54d96dcfa2b0e45c3cdac3b765eb545b56dfda7fd5bd8799d3f9088bd89d2aa7af67a130f9c486a6edd96acb956ffb0000000001ceaaa4f3eabf95dde14fe4b4e6e5b255dcab7764561c8b879b5849e7f869feb5a0a8bcdbf8a19c0d1c98ac4cd5aca27c8dd7b38e9d6993330e3b5f7c00000000073aaa93cff9757f94fe4f507cb7bd2a4d6b87d8ac39170f36b0944fb0d3fd7b3d51f8b7e09149ac9e3309b1772aeee5948742a71d8abfb9337cfbd3c800000000039d5549e7e834493f93c02312794c6e2724b082b0e45c3cdac2513ec34ff5ecf547e2dfadb856574eb9e458921a9f4e63b10bc33895d4be6c8c95aef59df400000000073aaa93cfd8aa9d49fc9f0d77c33b762e60abf9370f36b0944fb0d3fd7b3d5179b7f815d79f5dae1cd265c8aeb5767bd189c4bb875e6065b0bbe0000000001e74b36c9ad876731cfd1dce901a7e37bce9e6d9f9a3936da5f379a7cde86ee8e5d88f6ded6dc62073d94b1727e75738000000000000000000003e54ba927dd8c69dabd000000000000000000000000e54279be77a5f220000000000000000000000000000000000000000000c5ab9b6716ae5da00c5ab9b641835f3ec0000000000000000008dd7b25b06395e492c2008d57d26b00112834ba700000000000000000038d11ed4b63b5dc92c2008d57d26b00112834ba70000000000001a9c2c19fb7bbabcbd8eb1cad6ea6d6b70b5b376b7da9ccdbe9c76bb92584c1c1d5cbd8e92355f49a63c2cb21cc8941a5d3878e069eef7bd8000000000722b4c67bb2f7aa5c96e64f351e1b6b056380f56377a375ec96c18ed7724b0b4eafd53ed81258d57dbfa7e1b9676e44a0d2e9c6bd67cd3a5666c00000000008546271dc86c4e5b39aea3d614923d5d482c688c4a61248a42e4b60c6ebd92d831daee496145e1b2895c6e09ddb26355f7bb13ab038fc96c18941a5d388046259318ac32553c00000000006872635c3944fb875af76c9ae38164f74d1e54763ddfb1e375ec96c18ed7724b08d3e547e35d8b3a355f77ec7736abdab722506974e2a2d6b2b67c561b76d0000000000e240b476b6b97299efcaa34ed1abb6ed8fbcaaff9db3b9ca9058d1baf64b60c76bb92585ceafb979fa1c8ec59d1aafa4f3f6bd4396e18941a5d36a77cfd1f6e3fa00000000054ba966766355f4aa7887c27abca9a4c956f2ec590f0ab690d8b1baf64b60c76bb925855af0e7b28e5d5dd9b36355f75ecf72aaedeb622506974e2a2d6b676bc61d8c80000000007ca7f1d99bd5ef1a513e6b54be3d5b3b6a974ec5ec40781dfb1e375ec96c18ed7724b0aade5cf643068d75ecf8d57c9a76a13c7964ea2506974e20516924d2330d90d8a0000000002150e3b9c3ebd9e5771defd8e44a0c76f8bbd6bc6ebd92d831daee496146601f1d8e5e4b76375f77b9181d5b2b3c4a0d2e9c6bd6dca376cadf0000000000e4e8743abcbf9d535b4f6f687339bbdd7e679ea61d1cfbd83473ef39dcbdbecf3f1f4b1696dfbe1e6edfa6b69ecee3e71f4b67b5ec00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000003ffc4003110000104010204060202020203010000000400020305011534101114331213203035503140214124603290232570a0ffda0008010100010502ff00bfe767c2dd7635aec6b5d8d6bb1ad7635aec6b5d8d6bb1ad7635aec6b5d8d6bb1ad7635aec6b5d8d6bb1ad7635aec6b5d8d6bb1ad7635aec6b5d8d6bb1ad7635aec6b5d8d6bb1ad7635aec6b5d8d6bb1ad7635aec6b5d8d6bb1ad7635aec6b5d8d6bb1ad7635aec6b5d8d6bb1ad7635aec6b5d8d6bb1ad7635aec6b5d8d6bb1ad7635aec698ef1c7f592769302264669e5ad3cb5a796b4f2d69e5ad3cb5a796b4f2d69e5ad3cb5a796b4f2d69e5ad3cb5a796b4f2d69e5ad3cb5a796b4f2d69e5ad3cb5a796b4f2d69e5ad3cb5a796b4f2d69e5ad3cb5a796b4f2d69e5ad3cb5a796b4f2d69e5ad3cb5a796b4f2d69e5ad3cb5a796b4f2d69e5ad3cb5a796b4f2d69e5ad3cb4e0496350fb6fac93b4ab7e3bec8fd821f6df59276956fc77d91fb043edbeb24ed2adf8efb23f6087db7d649da55bf1df647ec10fb6fac93b4ab7e3bec8fd821f6df59276956fc77d91fb043edbeb24ed2adf8efb23f6087db7d649da55bf1def4e6403292f5ab5d9532f70873c727e80fd821f6df59276956fc77ba7dbe79e739ce620499968c5a96bca857e1036ce8f38ce1d8fdd3f6087db7d649da55bf1dee5c1de151c6e95e155c63e38995d09589a178f2d41be07fee9fb043edbeb24ed2adf8ef6e69310c323f32494e27970fa6cc4ea46c6796439fa817f70fd821f6df59276956fc77b770ff0811b3cc95b8c35bea323f28ca27f31ff0070fd821f6df59276956fc77b779b40f79ebb4f91a1f77134799bdf3f6087db7d649da55bf1deddac7e657b73e1745262687d39ce1b89e4f3a7a48fc227b6510d147615230a8666cf0fbc7ec10fb6fac93b4ab7e3bdb7370e6930646229cdc33d56e6e1ac8e374b243162187dbb52fa8215419e4cbef1fb043edbeb24ed2adf8ef72cc1eaa2ce32dc8571966222229b0a49638b065ce396739766a80f25bedda99d3c18c65ce755634efe719ac33a983dd3f6087db7d649da55bf1dee9d5b194a71261b3f85e7cab39ce5423ca4381aa68fee4923628c921c4914c1f89cae03f03c521c2911c8d963f70fd821f6df59276956fc77bd9c633892b049168a2265588c4d6b5b8f72e4cf13861dc4cf1c6d8a35246d96328770b3d399e07fb87ec10fb6fac93b4ab7e3bf78e2b020d9ce5ceaa0fa7838d987d4c1fce335e5f563fb67ec10fb6fac93b4ab7e3bf77f0ac4bea89aa13a823d36e1f93284564421aec3dbed1fb043edbeb24ed2adf8ef539d86b629593c7fab7067951318e92414768a3fa6685b3c33c2e1e6a633db3f6087db7d649da55bf1deab931559dd34bfa93ccd1e19a574f35307cb1ebb60fcf85aecb1c11582c7f64fd821f6df59276956fc77a4d2b020ee765eece32dcd41de637f4edccf366085c964b5b86b7d8b40fa69ebcbe908c679e3d83f6087db7d649da55bf1de8ce70dc1e5e4b22ac3ea67b607cd635d963812f05c1fa3645f4a37e735c274a37b248ed26096374325399e367b07ec10fb6fac93b4ab7e3bd172672c4513a694781a340ad41e9e410a7093c723658fdf73b0d69a564b269c3f365f6ee03f3238e474520a434a1fd67ec10fb6fac93b4ab7e3b8984e041defcc8fa80fca8b84b13668ca19c2cf54779127bf72672c430ba79a085b043edfe55889d21158674a47acfd821f6df59276956fc770cff18b12fab22b03eaa7f41e1e0b81cdcb5d52779f1fba510d1477bdd23e9c3f2a2f74c1b058ee6e58fa833ce8bd47ec10fb6fac93b4ab7e3b85c19e06471ba59051da2c1e9b707c588e4745206534b83dcb42fa922b84ea89f7ee43e78865741341334887d27ec10fb6fac93b4ab7e39164b451e47ba5929c3f2e3f5d985d2cc09791276bb0f67b56a674f06319738117020defbb18734e17221152679137a4fd821f6df59276956fc72b233aa22b43ea88f627858442440f1a6a83bcb77b3248d8a32677133d307e277e89e260b1f38cb7356675307a0fd821f6df59276956fc75c19e5c6c63a478833451fd9b10bab87f9c66aceea62f62e4cf13851dc4911b1b147fa5721f85c34ee1a78a46cd1713f6087db7d649da4310d1a9e591d3494e1f85beddc02a195d04a310d2a0f51c56041b39cbb3541f4e3fe9bd8d91858ce148a733cb9389fb043edbeb24ed29487490d789d5918c72c7b79c61d8b00f224f5e66449f19c3b1e8fc2b12faa26a84ea08fd5b20faa1ff0019ae2faa1f81fb043edbeb24ed26b72f7862e041fdd24769304d0ba0969cee5e9b833cb8a363a59051da30ff00ad6e1f952864e45218ec3d88fd821f6df5927695307fc7bf680f5317f38cd69bd5c3c2799a3c334ae9e5a60f937d9febf1c31edcd136688881c3cf4c672ca3f6087db7d649da559f1decf3f5db83e070f3bc69a0998442adccf3a6045c964e3186b7d9feb9afeb1ee5a87d4418ce5ae04ac1639fb043edbeb24ed2adf8ef5f3f473e3cf839b87b4e0dc2115a5b8496c4cc0c37e735c2e041b9ae6b9fa39acf0e6bf3c3970e6b9ae7c79ae7eab50fa79c02f22106e70eaf43edbeb25ed2adf8ef567d19e18f4943e0981f13e291ef74cda813cc97d18f5e387f5e8c2cf1c7a4881a4c12c4e86514cf15721f6df592f6556fc77ab2b0b3c32b0b3f8e39e16a0f9f1c31ba62a189b040b3c3fbc2cfe3fbc7e32b1f9cac70fe9678656387f7ebb70fcd8b80fb6fac9bb2ab7e3bd59f46563d88448e123fa59585fde3d19e38e1fd2cf0cac7a339fe3f9f55907d29087db7d64dd855bf1debe4b97ab92e5c392e4b92feb970e5c392e4b971e4b970e4bfae5c392cac70e4b970f0faaca36c8021f6df59376156fc77d91fb043edbeb27dba00d1a30b51116a222d4445a888b51116a222d4445a888b51116a222d4445a888b51116a222d4445a888b51116a222d4445a888b51116a222d4445a888b51116a222d4445a888b51116a222d4445a888b51116a222d4445a888b51116a222d4445a888b51116a222d4445a888b51116a222d4445a888b51116a22230e1a40d0fb6fac9f6ff6c3edbeb3f2ba58174b02e9605d2c0ba58174b02e9605d2c0ba58174b02e9605d2c0ba58174b02e9605d2c0ba58174b02e9605d2c0ba58174b02e9605d2c0ba58174b02e9605d2c0ba58174b02e9605d2c0ba58174b02e9605d2c0ba58174b02e9605d2c0ba58174b02e9605d2c0ba58174b02e9605d2c0ba58174b02e9605d2c0ba58174b02fc63fe87a69710c3ae0ea1b71e6978965b0366b83a14b6171fbb3db423cdae0e9b9f137fd2cfd870ad2faa1b85eed951ed3ddb3f915176bfd2cfd870089c884b5d87355eed951ed3dd9aac69e5d18458c72c7fa59fb0e34c5f3c2bddb2a3da7032d631f32d9972aea2751d917120edd93e7d47581309baa98999e719f23a10b5431571eecc04dc4d267a99f3982d0987311b0ca3157323f392a7ce61b22a1c846b0c8f81973e073cf2a45d44c98794c41dc78ddc279e31a222e27973d4cf9cc1684c3996c712d76a862a9265258a49191309ba91f971643b2c3496216eb3cdaec3dbf527ec15640d2259637432b1ee8de290d287bddb2a3da2b63723c6a01e525fa1cfc8816615caa0eccb8f4d9fc8a8bb569f1cb9e790f4b2cac32aa41635e2cf844ab98a6934d2431aad97309cae09cc2320eba53319a1fe0a16412455b2e65015817928943544d3b25a391adcf8999543db56866489d0f4d2cac229a5858a9cbcb25fa93f60a8f79721f8e35545f4e45eed951ed11b379e6631cf220cd14753c0d2217b331c904b9867c7f38f459fc8a8bb569f1caaa2c4a7a21be2190f1f9c46318c614b8f0cb07f0427c51ca9f5423d45136189f34512b72e027852ec0e93ca0956c189cde1790630e543db2e4f2445550e253b859c5884e63b2c7b1de367d41fb0547bcce39e2c04e9095397d4572a3da3bfe39fc898e6671b1c72b043ff237a2cfe4545dab4f8e547bb52f6556fc8f023730f7d4923228c8bb7e7321a4cabf2b31bdad54bb0b6f8e547bbe177b2543dbb6f8e4d7b999ea675d4ce9cf73f282cf307ea0fd82a3de23c5c1636719c67851ed14ccf2e689fe5cad761ed5f84549e714dc789cc6f819e8b3f915176ad3e3951eed4bd9559f23c08dcc3df576465d3206a6374518d044afbfe0a9b6164cf1d7aa997111fc2f66c72543dbb36f8ebd55c514c56981ad3035a685cf4c0d318d8d9f507ec151ef385c87e07f0a3da2b91b3190abed3a76ea61f23edbcd62a8173291e9b3f915176ad3e3951eed4bd9559f23c08dcc3df56edce2c1571b1ce3cb3c50b6c0ceb2754b9ff0338f160a83231085bae4c9af23c36591f348a87b6e6e1ed9e170f3c52ba1946b11c864e78e3b2499f24c2905be6c7e3ea0fd82a3de709636cb19303869d51ed14b13278ccac986cf10eb652b3144c863f4d9fc8a8bb569f1ca8f76a5ecaacf91e046e61efab40725479c65b95fce73a64d80d514bc0f05a64730f28ee50c124ef3006875ea87b4ac6bf05b6589f0bf8471be57d757f48dfa97370f6f442a8c7862cf1907865cf442a8e26458e328434cb470d4400b167d6e1607bba2157e13d8d91bd10aa31e18b2bf38e8854d16063b86431b39c063633c261602168e1f38431e059c61d8e8854c1a08dd9cf85adba172b510a4c7482e535ad662f5dfe32a36f21783e364b875486e4da90f0a3863871feadf9c130e47210b72f86275f7f04952152631cf2141d309feba683198d9aa8a8b391e66af265ca8eb8a950354c1b3ff00cb64919137ad15464432b94923226f5a2a8c986577b324ac89bd68aa322195dea9268e15d68aa39e29bfd26ef64a937aaeb60a977decddec9526f3d57dda543ddff48b31e4286d1cc55804e292aeb60a977decddec9526f3d57dda543ddfb1989847c3ef20c2d7989b79065405c24e11043068f580d436434f2f19ec471e5d6035010c263534f140d7dd8ed5af3136f21ca80d809e241118d1eb01a1cf809915d6c152efb84b3470b5f763356bcc4dbd854078e4f1bbd92a4de106403275ec7ce3bb81d98e564cce17dda543dde0e7358d96e466675e628ae46932d735edfa9b033a381ef748fe2d7398eae2fab1c987040f9c65ae8a4cc52b1f89234e7618d9e5ccd3e319738687100e69581079657cf271c672dcd59b92a1577b2549bd575b054bbe4510d1609e790993d1526b8862bbd9218970ae73b2f7701c990590525a540afbb4a87baa79983c2599296fe221b288f8266110fd459865904bc79a355f5b1470e608b2adc18e16aa277fe656f079462a69fcc115c4fe506aa20f34c56a2944caf1a78d5656c7e4f9112b6062642a91dc8d577b2549bd575b054bbe56c3124e5e2111aabae63a3f222e568045d32a87786c15dec935b97383ab86063c681edb2afe932a9a7f2ca57dda543dd572478c85595ad99bd3c3e1b3ac636354a47827faa90d1a2525d8ed46d93cc62a3ddab583ce0954cfe49aade7f34d553079217190c1e2525d0cd465abca8d53fc82bbd92a4deabad82a5df7190a1e2525c8ac455bbc88955fc8abbd92a88fc67f03e3f301433bc052beed2a1eea9dfe610873448c7d4044e3c3733fb15fe595f507d83436cc5ce4654634d2a9809c78551ee97e705c1d3958cf2ce0ac69eece5ce161ea09c639611c73038e7348232a382595495c44502a6f90577b2549bd575b054bbe4616c0e220f208ca643248b35a4b21557f22aef64a8f79c27dba8fb8afbb4a87bb9fc67f3e9c7fcb1f8fa72a5cce4a0eba11e35779ff0009516e785e41cdababff00d6aa3838d84d998d55f5d1450ab7cf2ae54dbf577b2549bd575b054bbe56b3665395657c588318e4acb3cabd557c8abbd92ac97109dc2c65f28141b3cc315f76950f754ecf2c84378251bc0d4ec46d6b241a45e06fd4c8dcb24411f11114a5430b6c0eeb255479ff002f81116271dcdcb5cb18e791a1e9c6e05372c2d571f1cb048443136cac3ab72a7cf2b0577b2549bd575b054bbe562dcb2c1559f1e607cf146db3b1c13c2b33cac55decb80570dc31d681b5a79ee31ea907f14aafbb4a87baae86f04eab6cba658346762cec98f895536694bfa8b601de3e23d64d342ab65f24fe37107945aa883ce338db00e93d0256ca52ce3964397c9315dec9526f55d6c152ef95a81921bf8e21d7ca5e5eccc724127953e338762ef64810b06412c4f81fc030642df144c82257dda543dd5342c9e22c09447711c594a7882b0487ea66ad167ce883286b85832b4d1339d343f44c3c44634d0d423c43e3d13d78c467441943582c394eaf11eed34358c78712c31cecd3435107040e52c31cecd34351063c0fe1382390b441945542459fc27802c8fd3434c6618c9618e76e9a1a84684752c31ced7528b9cc7502479c370dc70987888c69a1a845847cf0ce39e24ab1255a20c9950231358d637ffc62ff00ffc4003b100001030103080805040202030000000001000203101133721221223151617191132023303241509214404281a10452628260b190c170a0a2ffda0008010100063f02ff009fe2762b9773572ee6ae5dcd5cbb9ab9773572ee6ae5dcd5cbb9ab9773572ee6ae5dcd5cbb9ab9773572ee6ae5dcd5cbb9ab9773572ee6ae5dcd5cbb9ab9773572ee6ae5dcd5cbb9ab9773572ee6ae5dcd5cbb9ab9773572ee6ae5dcd5cbb9ab9773572ee6ae5dcd5cbb9ab9773572ee6ae5dcd5cbb9ab9773572ee6ae5dcd5cbb9ab9773572ee6ae5dcd5cbb9ab9773572ee6ae5dcd5cbb9ab9773572ee69aeda2df4d770a07b217169f3570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e570e45ce85c1a359a45847a6bb8521e1ea7361a45847a6bb8521e1ea7361a45847a6bb8521e1ea7361a45847a6bb8521e1ea7361a45847a6bb8521e1ea7361a45847a6bb8521e1ea7361a45847a6bb8521e1ea7361a45847a6bb8521e1dff692007679aece12711572c5da43ed2ac63f4bf69f409b0d22c23d35dc290f0ef8c5fa63c5ead26d2ad6446cda732d4cf72b5d11b3767a08ff00507299fbbcc2041b41f9e9b0d22c23d35dc290f0ef7e1a339feb283182d7141d200f93fd752db3224fdc118e41610be1e43a27c3bbe7a6c348b08f4d770a43c3bc7c875345a9cf76b71b574ee1a6fd5b8758b80ed199c2b42649e6467f9d9b0d22c23d35dc290f0ef08fdc404d67ee3620d1a875e5679072919b1df3b361a45847a6bb8521e1de331a8718ee2553fdbbd308769816d9f21361a45847a6bb8521e1de3ff8e920e1e49b20d4e16f5893a827c9fb8da9cffdeeef1d23bec36af88b74edb4a6c8cd47bf9b0d22c23d35dc290f0ef0b4ea29d11f2d4be1a4398f84f5be198748f8b726c6dd6e362646353459de6434f66cfcd3a179d07eadc7bf9b0d22c23d35dc290f0ef72997add5bd587310847fa8b48fdcad8e46ba96c8f6b789459fa6f79569ce4ae9e51da1d43677990d3da3ff0001003392ba317c34addeb7ac971ed19afbe9b0d22c23d35dc290f0efb2dba12eddaac91846ff002a5ebfdcb39b559130b90925d393f03bc73dc6c684e95de7a97c4bc661e1a7c43068bbc5c53646fdc6d4d7b4dad3decd8691611e9aee1487877f6116859e2030e65f5f35756e22ac68006eef7e198730f126c4df3d7b9358d163451cc78b5a53a377d8ed5f0ef3a2ef0f1ef66c348b08f4d770a43c3e7cbfea399a8939c9596e1da3ff0003a994d1da3356f5bd5a7c6dcceef26c348b08f4d770a43c3e7c91e06e66acb70ecd9f9eb74cc1a0ff00c141ff004ea7041cd3683ddcd8691611e9aee1487875cb89b004248cdad3f2dd030e93f5f0418d16b8a6c6dfb9dbd6746fd453a37eb0be19e70777361a45847a6bb8521e1d7f8661c6b21e7b277e3e55d23b504e91facaf8978ce7c3dc74ac1da33f210734d8420ffab538775361a45847a6bb8521e1d62ffab53422e71b495611615f0f21d21e13b7e53a161d066bde5067d3adc7720d1980ee72da3b37eadcad3e076672b4773361a45847a6bb8521e1d5b4ea08bbe8199ab2dc3b367e574f18d36ebde10734d846a595f58f10f92cd78eccd5bd67bc7677774e89de7ab72746f1a417c3bce76f878773361a45847a6bb8521e1d5f8661ce7c69b1b3594d89be54e9583b377e0a120d5e636a0f61b5a7e40b89b0045ff4ea68dcba778d066ade7bcf8860d26f8b826bd86c704d91bf71b3b89b0d22c23d35dc290f0ea190ebfa46f45ee3693ad74ef1a6fd5b855d1bc5ad28c6efb1daba190f66ed5b8fc87c330ebf1a6c6cd6536366a1dee6bb7676ab1c7b37ebee26c348b08f4d770a43c2b6959aedb99aad70ecd9af7f56cfac784a2d70b085d0c87b46eade3be7487ec117b8da4e72ba778d37eae1df18cebfa4ef45ae1611ad742f3a6cd5bc75e6c348b08f4d770a43c2bf0ec3a4ef1704d8d82d714d8dbf73b7adf1318ce3c6107b0d8e083c6bfa86cef725a7b366adeb3ddb73bbe43e25831a6c8cd6136466a3d69b0d22c23d35dc290f0a3a43f61b517b8dae2be21e349de1e1dc6534764ed5b9657d07c410734da0f77d1b4f68ff00c041a3392833ead6e3f205a45a0a2cfa4e7695d13cf66ffc1eb4d8691611e9aee148785346edba967bb6e77772e8dfa8a746ff002fcaf8790e89f0eeee9cf71b1a13a5779ea5f12f198787e48b7eb19daac3ac2c871ed19f9eacd8691611e9aee148782e818749daf70418d16b8ea4231afcced3dd5adbc6ea5b085d1bcf6adfcf73f0cc3987893636f9ebdc9ac68b1a3e4fe258331f126cadf2fca6c8c3a27a9361a45847a6bb852391de4336f4e91e748af8978ce7c3de7c4c6318ff00b4d9187384246fdc6ceb97fd47335127392b2dc3b47fe3e50b1c2d694e8cfd8ed0be1de745de1e3d49b0d22c23d35dc29145f4c61587c0dcee560ef2c3a8acd76ef0acf76ef12046707ac48f03733565b8766cfc9f96b5b78dce16f59ef1b99d59b0d22c23d35dc281ad1693a9060d7f51dfdf18ddf63b13a378ce17c348701ffaeaf40c3a4fd7c1358d16b8a6c6df2d7bfbad6b6f79d3b0683f5ee2848357d43720e69b41a4d8691611e9aee14f897ff4f90cb60ed5bf95b0ab1d7add75748ed413a47eb72f8978ce7c3de5bde3a37ea29d13bc97c33ce0a4d8691611e9aee148787c8fc446344f886c4d919ac26c8cd469d0b0e8335ef2833e9d6e41a3301de71ef7a460ed19f9081198841df50cce0a6c348b08f4d7f0a43c3e44b5c2d0564eb69f09563ede89daf72d13a6ff000ade85b78eceef9be91a3b37fe0a0efa0e672988d45948b08f4d7f0a43c3e48b0e63e451617690f2b50ca36960fc2e9de3459ab79f9c744ef34e8dfac29ff4cf39c30e4d22c23d35f84d21e1f2229d3463b46fe42631badc991b750f901def4ec1a6cd7bc562c23d35f84d21e1f28e95a33bbf08757eddd8efb46edd9c522c23d364c2690f0f9bf2f9893287845a291611e9b261348787a9cd8691611e9b26134898f99a1c06a57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed57ed52b5b334b8b730a45847a6c984fabc5847a75cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb55cc7ed5731fb5583fe07dd23b5345abc1226c7639b95e67a81cf04826cccbc1222f602003667ef9d139afb5bb1782440edff0c9b0d73de3733ab1e3a3f1f7d371a3387f864d86a1ff004ea70dc839a6d0691e3a3f1f7ce91e1d9477ad4ff720367f864d87a9f0cf3ab3b691e3a3f1d4b23d393f015ee48d8dccafa4f72cd293b9d9d064a321ff0083d7918c92c68dcafbf01349d8a47b0d8e0afbf014d27ea64ccdb1590f66dfcab7a693dcb3bfa46ec72e9f2ac68d76f92c98341bb7cd5a6693dcaf0b86c767568cce1adb52cfd3807f9159e77fdb32be7fb9669dff007ce847fa80013f50a99243605d9f66dddad5bd349ee5e3cb6ec727cd03b2646d968d8afbf014865765587350bde6c685640321bb4eb59e67f35a33bf9ac9fd48cdfb820e69b41f3f4a9b0d2589de6c4e8dfada83da6c704d907dc28f1d1f8e9d1467b47fe0532626daaf23b5592b6cdf4e8243a43c276f5a6e3467052d2c595238476f979ae903b2d835eea16db98f92cb27219b4f9a2f63f2ecd62ca47b1c724d3a369d293fd532810d66d2b34fff00cac893ec479d232758cd439f41b99b40f79e8da76eb56c7207eed49cd3683a88a4dc450b1a7b367e699523ba3b7cbcd17c6ee900f2f3a7c3b8e8bbc3b8fa54d868fc0be218349be2e14c871ecdff00851e2a3f1d247efb02b026c635f99da68e8dfa8a730eb69b13241f4956f566e3467052d196ea6e95251fc4d238ff0073ac560d428f6ec254788534d8d7710aeb27094d8d9e16ad391ade25462276516eb34fec54aefe34607784691ab271e798d26e214afd8da372b5374aaf0351d24d78d60da83b68b7d266c347e056145bf41ceda46c71ed18ee628fc68d211fcc7526e348b08eacdc68ce0a5a3f051f84d21e35971151e2142f79b1a15903001b5cb4a6773a0739a403aad14fec53fed47e0a8c749b88527da96b1c5a772be93dcafa4f72b5ce2e3be909fe03d266c347e0a16fd633b55875d5f8e8f66c36263ff69b5070d46b249b5c8346b29add82ceacdc68ce0a5a3f051f84d21e35971951e2146c03c2d169e346cb3da4bb386ad089a3eca11bcd3fb1528dd6d1b6ea768d63846bf11a4dc429785b42c99b940b732b91cd5c8e6ae9bcd5c8e6831a2c68d5e93361a3f057e218333bc5c6afc74e980d17ebe34e8a6b4c7e4762b7a608c505b9275ba9d291a0cff7d69b8d19c14b47e0a3f09a43c6b2e32a3c428fde0518cca02468b0859523c00ad1998dccda7f62883a8a7467cb570a06fea1a491f505d8b1c5dfc917bcdae349b8845a751cc9d1bbc8a6c8cd6d434c31de6d7224c809f200a74a4d8e71b7326471ccfce76fa54d868fc15746ff000b93a277951f8e86378b5a512d1971ed1d4048c88f6942360b1a3ad371a3382968fc147e134878d65c654788503d978dfcab08b0d36a7cefd1b3386d2487fb0a6c906a2b2646114c98d85c536dcf217e7349b88a653734a356f593234b4efae4b1a5c772cb7e794fe3d28b5c2d07c95c47c95b1c6d69dc3a96c91b5c7785711f25646c0d1bba96be16dbb5785dee56b616dbbf3f705ce85849f3215c47ca992f0083e455c47c95b1c6d69dc2b711f2594d85808f302b698196f0568819cabda461dbd781dee5d9c401daac39c15711f2594c89ad3b4044ec59f2dbf656195bfd82cae863f6ab1a001b946df32ea3ddb5d5b1ec0e1bd5dd9c0abb27895646c0de1fe30f8cf91a08e48f2c0d46d5a3067dee59721e0362b06b4c8fcfcffc7b3e678d4e59999636b56789e3faabb7f259a2237bb32e92439727e07fe2eca91c1a37abf8f9ac98e56b8ee34ca91c1a36957f1f35931cad71dc7bab647068deafe3e6ac8e46b8ee3d71d23c36ddaafe3e68f47235d66c3fe1231d1d829fdc50e03dd371d1d83af0e2a4dc07f84864765b956af0b7dc8be402cc9b35d3fb8a1c07ba6e3a3b075e1c549b80f52b65786ad163dcae1dcd6946f6aec9e0eea65c96e4ee0bc4ef6a11b1c728eeea747238e56e0bc4ef6acb8edc9de296caf0d5a2d7b95c3b9ad28de176720b7679d72e439adb1789ded5911936d96eaa7f714380d72a47868deb443deae1dcd6944f0ac649a5b0d5b8e8ec0bb47e7d9e6b46171fbad36bd9f95951b8386eac38a937015ca71006d2ac6e53f82b8773563b299c5653482368f4ab45e3bc28bde4971f3ea0734d8479856bbc6dcce4f88f98441d613641ada6d4d78d445b42e3a8674f90fd45003594c8c79045fadda9a117c8eb5c7a96836145afbc67e6831d1d829fdc50e034748efb0da8be4369ff005d531486d7b7cf68a371d1ee6788b6c1b91738da4f9d72e33c46d4246fdc6ca438a937014323f5056b8d8df26f52d69b5be6d42466a3e9396d665300b058569c4f1f64d7cac0e90e7cfe4ae99ed4d9a26e4826c2292b7f8db4cb1e1933d3a33ae33f8a640d7266a659f0c79e8d31b3298d1b569c2f1f642699b94e76a07c95d33daba789b9241ce0508dada0c74760a7f714380d18226dac6eff0035a70bc7d974f336db7c2d2ae99ed4e9a3606b9baecf3a3378228dc740d68b4940c8d0f937ab1d1308e083e3ba3f8a745f4c9fee90e2a4dc050423c2cff74e9e61a3f4b76ab3a26598519e01659e26d0c0753f38e3e97a533568073d646406b2db68fc14247899a428d0753f4685a35479a809f13f4ba9a73307dd68e53fec8c4230d61a0c26831d1d829fdc50e03d4d39583eeb4729fc02746d8c35aedb48a8dc7404fd22daca375b489db1c290e2a4dc05247ed71a46ce9d999a15fb11699d99c5227ec70f49c91a521d41769213bbca9a1138fd974b280d16d965b493052c4f8fc81cdc15a35af89fe16fdd171d65323da73ab0536bcea6ad390d9b06aa6846e77008cd23435a37d0613418e8ec14fee28701a653b393a86d5a4f21bfb5b4d08dcee013a57b325ad16e7348bef46e3a3b0564c268de34871526e011eb8f499243e668db5a1d27993468dafa4982ac9c79689a7c37f3b7ed47ce708ac87c81b051af7b43a422dcfe547ef229fd4d063a3b053fb8a1c068f1e4cd1146cd23439eecf9fc9665370a45f7a371d184ea3a3594ed160a44dfe54871526e029233638d237868ced5e11c917383400b41d195e11cbd25cd3ac1a0b5c1b20d60ab5f234216668dbaa8f1fc2af8cfd41169d62960d65323d82b2b4fee346c6f706c8d1667f356be4681c50632edbf9a3781a0c74760a7f714380d26b76db46c323835edcc2df356ba4681c574515d8d676d22a371d447fa9b737d6ade96ddc16c8c6a1474e753730a438a937014138f0bf5f1a7452ddedd8ad13b39a30406db7c4ea0b1efc86e7767f493fa8885a0f88751d29d1005a37d232751d1ea7483c3267fbd038f863cfd4f88885aefa875328e83369f3561513fc81cf418e8ec14fee28701a74d10ed1bac6dea5be18ff72730eb69b1324fda6d408d45371d261a9edb324ac891b926b9b333cdc9b1b058d14871526e0286378cc56716b3c9dd4c98dbc4ec590dd7e676fa55a63b0ed6e65e2939ab5b1da76bb3d2e1aae075009581c02b86a3d13036dd7d5b5f1e96d19978a4e6ad11e51fe59e85c616da55c35003505932b7282b86aca8a30d34c891b94d570d5951c61aeadb2462dda178a4e6adc8ca3fca85ef84171d655c3506b45806a5932b7282b86a3d13036dd6b2646070deb365b7eead2d2fc4558d160a812b03ac570d44c5186db5b0ab7a3c93fc732f149cd784bb11592c6868ddff00a637ffc4002e1000010204050304030101010101000000010011102141513161a1f0f17181b12030509140c1d1e1609070a0ffda0008010100013f21ff00dfec813ae3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f2e3cb8f21890c04cebf1ba9f88612040afc9fff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00f8bb039290d8adf1ba9f886e73f93d4e1b15be3753f10dce7f27a9c362b7c6ea7e21b9cfe4f5386c56f8dd4fc43739fc9ea70d8adf1ba9f886e73f93d4e1b15be3753f10dce7f27a9c362b7c6ea7e21b9cfe4f5386c56f8dd4fc43739fbe2ba433fa20cb670c29f47a94fa68174b1063a13f01a9c362b7c6ea7e21b9cfdd258395200025fc511144c4943439b21aa94fd840c49857f929958847c180361d0cb00e08afe76a70d8adf1ba9f886e73f7481204245ac8974d6002bbd19c3a3d052014018f55e782b3470de7355bf3b5386c56f8dd4fc43739fb98a0842e4ef083868f501c821db8a84401188982b23c7557f375386c56f8dd4fc43739fb8711bc3fa4d151fb20a0c061eb089c4374452578771fe7e6ea70d8adf1ba9f886e73f704bb4fe0a200e70fd9ec182c5c78080b153fdfba170d8b27e06a70d8adf1ba9f886e73f70ec44c87d1171813858728f50e5301c959eca14e0b1d04bdca414ae10bcdee992364e37e9efea70d8adf1ba9f886e73f700cb818855a5cd714452408e6bdbd476eb1628b2132ed1eedd864c919aa6157162f7cd4e1b15be3753f10dce7eec9a93d8b221012260d10510e40711d6e98e8c8ce0cba7413212f81078088484899354794033a9fefb94bcb0d4286192180154de01f636220a041007e9656ace287ded4e1b15be3753f10dce7ef3a907d1ea4d5650665dd0249c120e4990d23a9117215c94e4a5c603ba2c480603dc580c5c955904961409e68113537857b0c22972a214ac2c980c5c1f775386c56f8dd4fc43739fbe44012851172972498b3ef3730a0b15501bdd69a84c54d9566135952980c58083490b10a8953b0baaf61ccd2cf775386c56f8dd4fc43739fe7e40a2e51c3243926aa93973a03d19dab22a1021104803f48612523f7ee6a70d8adf1ba9f886e73fce240124b00886148fdacd9273507aaae29b2017725d4042a013822bedea70d8adf1ba9f886e73f581104e49a2a6649fc6a04b0d3fd27d71601520a771ea164e36e88754dfb175ab88ea3dbd4e1b15be3753f10dce7ebd5c43c2a9c679aeb11f884be4bf5c91e273bf44f74009b5fd8a6466c825029c11442bb00e80fb5a9c362b7c6ea7e21b9cfd47a313ea0a25029c9354705028530ca73a8b7e256a1811386ce810088018014f66b2a7cca84709291fb40111c1983ecea70d8adf1ba9f886e73f49c9c0039250884e1b95d5734e735952c12841fe16e42887e4122c1fc2767e899a9d425aa189ba265ed526535d74c902b154cd399ad9ecea70d8adf1ba9f886e73f4e6e82141642e9cec161c2133737854578054ea705a26b61707f000882724d1106ece812a60c5b25ee530720aff84e042e0aae14ae3d8d4e1b15be3753f10dce7e89f261b88e4a6e4a9c090c4628162a77831b45580cc7f0269989885a810b29edd3340da5b75cfdc2001043828e50354b64b28965342b1f5ea70d8adf1ba9f886e73890211809928a42740fdacf759d6587a0dc584ffd08980a620d1502331f7ac70c4ae14fba6e2a0831d3fd7bd2ecc7611fd8d8151658bd66a70d8adf1ba9f886e738d330e629627981601512bb8bfa9ee4a915174e0c2e0a915e1b8f772cce654a3080ed19000061203dfaa61201e51f263bf547aa4e1636f56a70d8adf1ba9f886e7384c4c895e27471725511721a7faf59008621c145c4cf72c83e7322e043dc1b823dba4a586a1434490c00aa90377507f002880620d54d277805570d7d46a70d8adf1ba9f886e7344801ce08c6054f3e68826ea9920000000c07b010dc7f59a1953c06cba619ce7355bda69b1725576125850278a004deff008422018a6764420208c4595354c7250fa75386c56f8dd4fc43739aa36e71b269f745814e5f17b50264ae77c910582091f4a8219e5bfb2d75198a9b2a8926b2a53218b01f879224450dd5565316593b002e3d1a9c362b7c6ea7e2191dab9cc93b40ae553d6c2682fee304cb50466b1dfaaaf3f716f5e50a2e51cb243926ab32c9c941f88dba2c429898c54558731a7faf46a70d8adf1ba9f8860a1a0173747082b1fa4010180901ee188c103106aa5a9338ad929e84c80fda3960070457d248024960314430a47ed55b2e7407e33c0355c94eb003f48637d12f9c75386c56f8dd4fc41fd8d8029d71cee7bc16f8f7174d763fda70258e27a54a961a7fa4e862c02a0926baa7d9266c314f7faa738bb3dcaf0251b26a6f986e233b9b830d4e1b15be3753f106994cc83c9fc0a1c6592c81761c07e901d288df389ef93f66c8b5399fa27aa004dafed3b3ee98620eb097b85888f6c1d38d8ac4c4e46e2ea898cc8ded0d4e1b15be375bf10d867ec92c1757adca739555d14ab82e2c8e9387e8da1588e044c1b265821a20060053d78277c20ceda84c7738d9361a92a8fb94c0bfe8474c90e08a2965ed02b5386c56f8dd67c43719fb0c44b8ee15f7583102eb05d07e9020e087d03620d51034cdde09ab39a715d3e446533ca9d425aa080ed193330989912405d07e962ec8162eba0fd2b827397d265669c04ccfe931322cebd133d55e4bfe84425972ba3960904110d8adf1baaf886e73f5942dd55f7545454572779aaffab3c9d12c11f6d1cff14e95b143b1320747e9ca53077228cb6c151146bfc2e8861d91ec3ba965f660b205a1dc8add9511511516f154392ab2f4e14212363742e98ec56284066a2d0d8adf1bba5a1b9cfd788f6f2b10fb583b857dd5623d9620b191c4f7430c100fd3cac5d48603b2a44660202a499a186c303aaff3cac410a75fd23fa2fd562745b7d2c058fb2c0efe514d13ab765fbfed5fbf958ba90a76588f65887747faaac9320187aa980946c94762b7c6ed9686e73f585613aea809ac47b2c5d910e1a0dd539b9fa588cc41c8ff005498900da9a17f9e5620a8ebfa47f444d349994fb447f909f782dd97effb57dd562ea429d9623d1621dd1914cf7ea10247098913a02e1fd043863827a21a2e50d8adf1bb05a1b9cfd83b943a074403221d00d021d76ae90806453b17626c0874221d06221d0926b7320c89e8ee86e219193605d888743a3b20740c88747a0f55d88870c9dea9c2172b186c56f8ddead0dce7f27afc362b7c6ecd684d6082a4fe4ddddddddddddddddddddddddddddddddddddddddddddddddddcd78001586c56f8ddfadf2fb15be34800410e0d170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3170c5c3100000602400ffc1f06c4bc018ae187f504871300867f4337e45e187f7df9141253198900de570c3fa8207007ff008cd7e181709b1faa671da6461bcc87bda07810d17c7fc66bf108dddd4203004e08ac3699186f321ef542959b0c086e00c3fe335ff43ecc626d510da6461bcc84479358cfec440b9214e799d48c06740158d205f83d6f4f21832d051b1810253338863de1a7539d858047f909a917c3dd487dea93b546601df4b237eab137f13988f520923a5644e3fad45e1912385d9107e90bf45ce9137e90bf6586120c2ef16c1985c9b046a0edff00445f0e7520773ad3f558e420999a1a7ed842402036f8c251ae589cbf89c479eb445fb97a1c453678750810027015f8ad7e18489d8d8b89a080c6629f745c155c39582b699186f32100160099101d14ea683aa13e72134313ac0260c261c372516f5681e04345f0b463c886024b3bb6681073031d0847b43180040379cba4504887c013e80815c45e40b40e40e80300b6121ed540a0465af7444c0be7fe9324130f8008117703cf481844d31fb80d0cc06397646c582b0e282bd250dbad09cd2602ebc030509831d05026080683ae34f8a35f86a1e4262905995bbb432649cd42b6f9437990819ca5f50220e612582102247840ba7048d8dd630209633c13d910006067e9d03c0868be168c7910652e02ce984047703e283e9800ba21313018044386299cc02d51bbdcf018c37acaf666042c980c10c7ed429820322502fa2881cb1040779404d3c92ed10bac7bf3486dd6470b891bac1ae384b3a611192cd80ea8acb040800b017c4b5f86a1e420180e0c8844002a995a13fa667a1230de641487162a62ea84530fd9e8133cfa40c90713e2f4e81e04345f0b463c8804eda621ba5a02fd47831deaeb7abc06bf184a2aeb9c7e9621c5830688925324958c878c3c48dbabe506e6ff0042203b6f061b7591b6610d614b08e4cb902e40b3ba84f07ebe24d7e1a87910086e99a230301620c7799044386442d89f52c857e8884b81c1812009258041058188e88b8c060b292f4b40f021a2f85a31e443759886e9686a1e0c778badeaf0251802e506d946058019afb5f13221d51e2186a20ac4684e0103609fe22e3b17f086dd6450c62340c05d481026ab9e2e789d201e1477f573c4dba2c16f89d7e1a87911609c99b4ba3bcc84023c6d033082c0c7fca04fab62e8426120d22458406eb919fa9a07810d17c2d18f221bacc4374b4350f063bc5d6f578198e060fa81880224e2d508bc1d72a4390183e60f32ca1853818843be4535e840904105884194917fa844828a32008a3cd7261b75908f7092421a6d751428cd319c2281a8864d16f8b128ee5c0e4e47b53737da06004be7f13afc350f222161c0c5629252371430de6420d9b3345294c531d47a0f5aa8b1e89b76643d5a07810d17c2d18f221bacc4374b4358f063bc5d6f578019b62cb23f3818823084909914c5dc2e7122af00b18cff00c0fea00f31c5ca3f29bd0f780765860118080ecb919086c56801b22cc68b145c29a04420e340432360cda9b7c50200cc4aab8aa2b370676bd0060e0c0bcb8aa273317616f41470b018e8a63f69c829915fe9ec63f66992b8aa0000003008871ac0915c55119b8312d40804043838ae2a87f0cd322246448b924504012260b638a7649fdac4761c88bf720fda39180188355c5574daa0a787007582facf425c069fd1062077982c5920a06400e312dd043712c23977407461dfd68455fb85356f91bfe5c8040702831d8cc5200500b018590130c74e240c0110407290013e5843f59ff9e0bcb2459d5d7d1186ebc4810c0dde8d067e8100800c3ff971371ece4cb8aa9387767213612ce32e2aa4415d9cf6a66bd9c997154cab83b39eb1443f03d9d7154780ce29cdff0013a37830d63c88616cc61b865ed691e0c35ef23d7aa7886c57ff008909642131692e24831e6523a728616cc61b865ed691e0c35ef23d7aa7886c57f92cb44153d9136cfa41701473bd48a7215ae023b41cd8399dc5c91524e1dbe8260401c8717244f6c9ccec41956a3e251c6cf5994cfd147be9d8a95197e5f48851c0f003cd72445c7c711928616cc61b8651cc6b1230d9e00c355325f45107e80828d32eccc748f061af79082ca140ccbb20e42b9009999a6043869a94754f10d8af13a0de24608b9d6e2c35532874a042172cb44105bc08e3e287203203f6b1f0a17a0ede78111e5cf75335906591a2186c4621603a011817082063582e2b11321430dc8c02b0f47ad517333a828f12b0fa064e04c11444c5f10d97868de0c358f22185b3186e1942694d2b8b239726029d1e80482082c4553638b880d23c180f21351d48f0e7891c622ac7e8ea5237795c5a1aa7886c5781476d6c915b270f01e80374e3e0514771fd65f1262dd193c2d5d4d0532e0b8ca813120e840c0c0f0eb030a721abfd83f0247dab0783613bb08389277d2b06e048fbd21864100077aad744d0017340c83a274348e84cb220610830253cf0d1bc186b1e4430b6630dc328302612431d3ee4147ddf2002eb0522cc41cccf8202040b888691e0c1ce0980ba286789980e817655153944ad3aed03948c8c324354f10d8af0399499e6509cd3ff00a0a01672cc551ce0c1ae2062495d87c51602782930e6c0b9d1626dd182764366186eb3106e8ffea41c4577ea0e4a537ad60cd9a6ba53d12321698ac0e6c983558a4326e7d1cd1bc186b1e4430b6630dc32f4494adb12c506e879430e4c493986acf830d23c180871c1fbf66223e306ba89c0a27950d53c4362bc0c42b35831c1a10f92e7919401019a2c080c11467c4a10080f206651c105616fa43ecac48a4d218830de662040210707146bd8f0220c8c05c15376c61f68a8b91c9470f87d4aa0080c0061002246fc4a3afa8a02cb863c0c8cde8068de0c358f22185b3186e1942c990ea45dd94081b6ef4211a7862435fe061a4783029372f222d9f87ea868fe61aa7886c575a658bd7d4f04cd6074f882918add288072c88cd8709e792c10c5403c186d338b2b9fd1a4013a67d8fb8491f6ce2679d14200bea222fd88000300c10c0b0d61870b46f061ac7910c2d98c370ca0f1291e601b6acc40100180016082faadd61b4c8c348f06053da78f5888e19f7a30283a8c354f10d8af0318a4d604665a3841c1424e49182d11c8f89510e0858873056086e32639bb844c396799ec8cd270135ce0c5e7ec2382931d0a032c462204189c8c02183ceeb58e3b83cd097a4026ea08833f918890670f895e0d2b8a1a37830d63c88616cc61b86500854d7780fa0cc3005993e28143ce392bfe43ac491a18691e0c6730900077eaa40f2849406c0f3730281caeb986a9e21b15e0419a1902001984cb80c7fca956736a670f09835841a15871b74f897c5c05437f4055274244ce032069aefe896123fa4195487f5a7a0ec981878917f40e222cc81a110d2082c50ccd0f4194346f061ac7910c2d98c370ca028e00d628824410c4620c46586b18f0b18104b2394394e0705691e0c05e2e8bec923d2034311769199f8764d822c21aa7886c5783cf87f48d5ea68977b7a044e377041b9e53bef8a387de2449d4accfe2181439881284b87aa002e3ed2b011214470f45cc14181b42be9201016794ea5667f106306a698229319c99e2b9828614c0609b55d76375cc146e690ce2d068371d8ae60a9438670f127f418a263210b33f88024d5374000000000a046712f326b98298dcd82c87c50976375cc15d23b557462c4fa3ca1fda1022ec2171018001a20809387a2e60a240642d588080041c4147093ca9217bf53f88eb93d02c01a037ff008c6fffda0008010100000010ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00dfff00ff00ff00ff00ff00ff00bfff00ff00ff00ff00fd7fff00ff00ff00ff00ff00ebff00ff00ff00ff00ff00d7ff00ff00ff00ff00ff00febfff00ff00ff00ff00fd7fff00ff00ff00ff00ff00ebff00ff00ff00ff00ff00d7ff00ff00ff00ff00ff00febfff00ff00ff00ff00fd7fff00ff00ff00ff00ff00ebff00ff00ff00ff00ff00d7ff00ff00ff00ff00ff00febfff00ff00ff00ff00fd7fff00ff00ff00ff00ff00ebff00ff00ff00ff00ff00d7ff00e4ff00ff00ff00febfff00ff00ff00ff00fd7ff8afff00ff00ff00ebff00ff00ff00ff00ff00d7ff009f3fff00ff00febfff00ff00ff00ff00fd7ff3f7ff00ff00ff00ebff00ff00ff00ff00ff00d7ff00ff00bfff00ff00febfff00ff00ff00ff00fd7fff00fdff00ff00ff00ebff00ff00ff00ff00ff00d7ff00df9ffdff00febfff00ff00ff00ff00fd7ff5f5ff00ff00ff00ebff00ff00ff00ff00ff00d7ff00dd3ffeff00febfff00ff00ff00ff00fd7fff008fff00d3ff00ebff00ff00ff00ff00ff00d7ff00e4ff00f7bffebfff00ff00ff00ff00fd7fff00ff00feff00ff00ebff00ff00ff00ff00ff00d7ff00ff00ff00cff7febfff00ff00ff00ff00fd7e7fff00fe7eff00ebff00ff00ff00ff00ff00d7e3ff00ff008fe3febfff00ff00ff00ff00fd7e3fff00f9ff003febff00ff00ff00ff00ff00d7a5ff00ff00dff1febfff00ff00ff00ff00fd78dfff00efff009febff00ff00ff00ff00ff00d7bcff00ff00ff00fe7ebfff00ff00ff00ff00fd71cfff00f7ff00f7ebff00ff00ff00ff00ff00d63ebffd7fff007ebfff00ff00ff00ff00fd4fe3ff00bfff00ff00ebff00ff00ff00ff00ff00d07f3ffbff00ff00bebfff00ff00ff00ff00fd6ff9ff009fff00ff00ebff00ff00ff00ff00ff00d1ff00ff00fbff00ff00eebfff00ff00ff00ff00fd7ffcbf7fff00ff00ebff00ff00ff00ff00ff00dbff00e7c7fa7fe2bfff00ff00ff00ff00fdff00ff003bff00e7ff00abff00ff00ff00ff00ff00ff00ff00ff00bfcbff00fcbfff00ff00ff00ff00ff00ff00b066b238df9bff00ff00ff00ff00ff00d7ef3d2d999b7dbfff00ff00ff00ff00ff007e20720c6e07ebff00ff00ff00ff00ff00d7ec24b20ea1dcbfff00ff00ff00ff00ff007f474df65793efff00ff00ff00ff00ff00d7ff00ff00ff00ff00ff00feff00ff00ff00ff00ff00ff008000000000001fff00ff00ff00ff00ff00cfff00ff00ff00ff00ff00ff007fff00ff00ff00ff00fdff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00dfff00ff00ff00ff00ff00ff00ff00ff00ff004d7ffd7fff00ff00ff00ff00ff00ff00ff00ff00f657ff00c7ff00ff00ff00ff00ff00ff00ff00ff00ff007d6bfc82cbdacb5fff00ff00ff00ff00f5d777d262ab3ebdff00ff00ff00ff00ff005f65fd2a1c3fa79fff00ff00ff00ff00f5f3efd3efdbfa5bff00ff00ff00ff00ff005b05bd3edb7fa7bfff00ff00ff00ff00f5b737d3edea1acbff00ff00ff00ff00ff005b5f7f3efb2f29ff00ff00ff00ff00ff00f5ff00dffbedfacfdfff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff003fff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00e1ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00f5ff00f5feff00ff00ff00ff00ff00ff00ff00ff00ff007dff007fbfff00ff00ff00ff00ff00ff00ff00ff00f55ff7fbff00ff00ff00ff00ff00ff00b638d395f777be7fff00ff00ff00ff00f7eb63975394cb9cff00ff00ff00ff00ff003f7f0575c546beefff00ff00ff00ff00f9f5db775cd5bbbeff00ff00ff00ff00ff00b59b2f75b75aba17ff00ff00ff00ff00e05fb4d757f7eba5ff00ff00ff00ff00ff000d93cd75ff006fbb3fff00ff00ff00ff00f7df77c75fe7dbfdff00ff00ff00ff00ff009aef8433d88b9f9fff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ff00ffc4002e10010001020306070101010101010100000111002131415110617181a1f0205091b1c1d1f130e140609070a0ffda0008010100013f10ff00eff0ac28ac33827cdeebaebaebaebaebaebaebaebaebaebaebaebaebaebaebaebaebaebaebaebaebaebaebaebaebaebaeb906062c404c75f2deebab60daf929030d6bf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3eebf2cfbafcb3ee9bee802019e34e15d8f4f96f75d5b3ad7bbcd3ee15d8f4f96f75d5b3ad7bbcd3ee15d8f4f96f75d5b3ad7bbcd3ee15d8f4f96f75d5b3ad7bbcd3ee15d8f4f96f75d5b3ad7bbcd3ee15d8f4f96f75d5b3ad7bbcd3ee15d8f4f96f75d5b3ad7bbcd3ee15d8f4f96f75d5b3ad7bbfbe0a0c974e4f9a641f22cf913ef4e0b70adc0cf1223921ef4a223fc31c794f90fdc2bb1e9f2deebab675af77f504400255cab19b02c5733efe94ed26572aef69b970884e709a1660799ed57c9b2e03dd4324957046934784df8ba3ad1f622d20733feff00b85763d3e5bdd756ceb5eefea99f248b8b0e7c5dd5a8c80cfd54508ca32fa07dde958106d8561c125bb99d69df2b90c8398d4bb56e3f13efc7feff00b85763d3e5bdd756ceb5eefe989130d60c39d263715d568b11c54be4c71c784788a24e0170bf34c37d260580c44c1a0592c13236ea3feefb85763d3e5bdd756ceb5eefe8cf42af0c544c4dea60a134191901078924870a256103a2b9d1a44a422d0fb3fee7dc2bb1e9f2deebab675af77f4413009ab0194fe8fe1a0e178d5be00e7fd4e659b7177d4ff87ee15d8f4f96f75d5b3ad7bbfa324c6382bf45a5b20edbc64a61c4246525ce4dbc44d98760012b48fe59b85b748a88b2aef043acff00484653758b81de54414e996c5c56e4b55a1a633598ef1b7fc1f70aec7a7cb7baead9d6bddfd0b62eecc48686d734cfbaf4a8c3714b0b1e7c4dfe23563b8fd4fb71a73660ef73e19d6022c3aa62f365fe92d2198db7681feecc07dd15b279387a7fc1f70aec7a7cb7baead9d6bddfd408c2346b3f8ff6a0e340428f66898203c96e7bb1e3419ab2387131362c03ccbf55328781038d9ef6a4a3444a9f76b031315d7377ba1fd301c861bf601fe51e525f14d82a5ff3c59e5e7a45b90d465948d94519c8126c7b91dffdfee15d8f4f96f75d5b3ad7bbfb0a00301c83f38d3394601c80b51c18c150d10b1d167deb7e744bd68ccc970f530157aa51929ba7177ff42ec9bb42a650b12e1f671ac48a4ce2f2e06fe1b2260c00e1703efc6a6891dd5971ed9d1760cda3fdbee15d8f4f96f75d5b3ad7bbfb9a208464795346bc51d031d29695a69f9a2c4160f4158e95808c880e47f5c48a1ce0f2e2efe153785c870682ec19b4360544dda3f345528ecacb876cea662480e3703efc7fb7dc2bb1e9f2deebab675af77fdf3c46e3cdcf818d36525f14dd6b038098bf60bfe780de40b1e3d8a6fe35396509651453ad6a9cb81f79febf70aec7a7cb7baead9d6bddff00719604aad829d1dfcb267ccf48a8698c4e5bb45ff7c580eb28b66f271f5a742c4b3f1e66250ac1bd60dc7fa7dc2bb1e9f2deebab675af778d3d09e80175a2a869ccde264ff00cd13026257d3e38384d32a0cd9ad42329b2c5c5ef2f15e5a6735926f1bd40b234e5943725636a18e7f29cffa7dc2bb1e9f2deebab675af778e3312c31ccf9bca924d084e18670e4ff942022237133ff9230d6ecd6437adaa5586b41906e0b5602ca6619b9f0374ebfc3130840be2a713139d18c1e3c8b8d2b1635938f2713f9fdc2bb1e9f2deebab675af778996b8acdf83169c60f1e4dd6b16b41c272abe1dd0b8fb8f6e1ff002454cbeab66f230e3341205a5918f370282d98eb02c1fc124870a9684a88b76a9fe527cb5a232e23da682c1c0322383fcbee15d8f4f96f75d5b3ad7bbc28250a40062d1a24a9ec971f4ac650852d8a7c9ff6a532ff005cb3e27b51536b57450ed853fa86e713fe208392967ab97de2b7e5e2a68846c367a394eb3fc83402486e380a4a0d9fc9b9c6a6a6cc0e3711edc3f97dc2bb1e9f2deebab675af778727a087039f17771a8d5fe81aaee0bd13ba66ed8ae2d2490e15271efc2c989c1c4f4a9955c7db34e3994274b369f7ff00037e13d6012b4d30b8b230e6e2f1a91916516fa71718fe9053a002fadc7d9c29ba066d4f8a832137571c4ef2fe3f70aec7a7cb7baead9d6bdde086d4bccec0e19bc29282adc55ac19da4bfdae3c2369724c1f73798d1e68606cb83f7bea3e2de8b27c39ffc1853431862e6e2f2a98e0c390cd6e0bd40d072cd66b7adff00a1960422592944dcea9cbed14a0e1499b7723ba8402223713f87dc2bb1e9f2deebab675af76d1c0250c00674b72eea8cf8df68a416284d8647cdddc68000001603c02849bf5d5b9a47023a146254249b9570f93dbfb447a3758f81de54c797466b535121017fbf1708fed0c8179958727078d3e655b8898d636ee8af93ccc3d3f87dc2bb1e9f2deebab675af76dc1ac80be5713edc69f0066f77714158c662bae2bbc3c5a60c185ef33ddc29be966d7eaa15c31b7cd38667f4c2a76633136ed03fda55361f24cb9de9340c80401807f788b04231303970795591a6d0663b92d53ac98e6ce5bc7c7f70aec7a7cb7baead9d6bddb208457175c0ef2a5949bb35ac3ac005f5b8fb38f8d180108964a95294c70c65fc6ee145a614fea1bcc4a3cad7ac8ff3c2d2106f81f01fe51c335f94d828642b6b37e0c0ff00816118e90364a8fcaf2cdcb8985611000ab608f0707978fee15d8f4f96f75d5b3ad7ba91200255c02942ca9e4f3e7cb75233ce265a39fda81b02002c1fc2d2509cd64379575a5836c81c6adc774acbec7df8ff00208c9bb43e696d46096c7f05620ca6639b97037ce9ff0011b24ebcbe070f4a71ec4594625444e14adfb41ff7c5f70aec7a7cb7baead9d6bdd583d9ccbe8f1f671a67c19735a84c02db1317e0dc7f27a109d1737c72df51a6223651f345128425be09c593feff001c4411ce0f262efe153ac124783d9c6841066c83fe383b881383cf83bf8d3c8ac53d97154680b7f87798787ee15d8f4f96f75d5b11211496eb00a420dbfc1b8c2b3e384e2f1396ee3fd3161613072fb7ad5bc1cd0663b92d4d2c0624b8e2bc72806e3cdcf818d27525d14dd6a22001397ed17fcff908f26ecc6a6205906260fc3bca99872116d3e1eee3e1fb85763d3e5bdd756c6fcbb2715557ac53ea22d519713ed34170e00800c0fe8a2500481c469c482f567bcf6a10b187cb43bcf6a02245a40e09e132c254b60a54d7f5c67ccf48a8c9009cb760bfeffcc4629ab99cf9f2df4254782945469c4663473fbcf83ee15d8f4f96f75d5b1f30b8d2b85451594cec79181fdacee3305c701562873a0c9372562a35e60e7f4f4f0c14098d5f4f8fb26941066cd68d61648f17b387f15b09cda14c1885d0140101b8c68441191fe7847b016d5e18b8cd4dbe04cec4e26270a1e829f0476fdc2bb1e9f2deebab60b8d9d99187c273ff0080215298e38ef1e67fb40924425947cd0c38068b91e39efdb0069f19e40deb533815a0c83705ab0265330cfcf81ba75fe4a118b99ebed6a450082eac05284940b51664064e7fce553fa8689bc6f46ae90b36038953771ac6189e7c4e7b7ee15d8f4f96f65d5b3abfbbf8cca6381c6842ca61bc919c5bc76f3b018bec73dfc6a20cb896ce5b9aba5d133ce5bcd90322f0ad9bc8c38cd0c816de4e5c5c0a38643ac0b078d412a06ad004a135362e0e2c77caad6406e5da90d5810695721836381fd3128c005f11e23139eb471496e518353fc5b5939f071d9f70aec7a7cb7bdead9d6fddfc1062eba04d1c108e026f2befb08304aee2680494a095037d4773acab144d35a69d60e344995449dc3bcc2ac9b8493a5f7ddc2919c5753025f947d5282312bc54d026b0d128e5c07bcd6b1c4229062eba04d08c322e491b714795473135953c465238d302166450d1b547fd5505c226749b12a31476a1095273a424db40a8b101aaa906265d026846191de46dd29e04d08c5c7448f16019800b623c0e273d299a58839fc863eb4c45249112cd385763d3e5aa1fbe5b3ad7bbc701131332e86b5608c338580e2e6d2490c02cc31896afbd222e42254d2a220c5ad8c06eb6345254c404b9c674c92c71c30328df5128b498cb93eaa5003a1389b9ab8c5f22a4128a25c73f6341387094ceed6842e8364708dd57218c0b64723de2b52eb9c31318df4a13243564f51a170c5b460fba9c848e4323f4d4824cc60ea654b02b8152ade65c631be050904491cc549038fde90ab122e6bc70a9d540b76f804a63be85cee5d5abc3813ef59b3339c62ba4e8509744b62a31f749165358181c68bd917df23baf854c5ca408ba548a3032d5d29c09477a839053690b424991f5c1a4c059944ebe11174c5cb01c1a9d4fe83a26e4bd622999c5e5c4ddc29c2bb1e9f2d50dd92a30aeb5eef1f443d6eaba4d1e69ac3ecb9b0e9bdd57f097ad7454914c491c408a1003112e93bddd58e6f8afdad5d86e6af07387289f7add2fcb97c9954a38982c08caee226ae5a23359aef5bd183daea89c493a53953a296ae2fad62387e6baf538a9d01589db1d825b050230c55ac14cd845d1a3bb8eccd4f60e8d61a6de12eb5d37ba93777af8f6ad3c60046050960b38df2deefa000c0f1454e800beaf1c5c26852e31c29c2bb1e9f2d50fdd2a30aeb5eef1e72c442e9be895b0c978311d4d4a5c49346047035a90910594d0302ba6f7561717bd022c122af8b666fb9c3d1ab5c09a04951928066d596b5fefa506111096bc499732af2550c4178a8670e858dd1c4de8ea9460f6c55ee7dabd9531bbe3503707bb4dd4a3040075a6557364b6adba14100695243808a74df44ca83239626bca9259cae0e2b84f0a36530b07028eee3596a7e74ec3735869e91ef5d37ba8cb6996435b5ca6012516fb281b718613ceaf0b1cfad0488b78c56a03513e0040144239d22a8eae4f3e5cb75385763d3e5ae1bba54615d6bddfc2c4111a090e14a33cc17f5a01056388988b98d6398988b163662851d4a969cb254bfd5a88cb2acad033169c4491e543ce06719d4c486cb0c5c292c368c3752ccb9a6c44d442184646a70b16c00b53a9219b22c75aba217220c8a52dc5bda021594c34386c4921a570919430e75a900e2089e35854c6e86525ca8ee6286b76132c176a0630e234822d833033a55246108e74aa58c200ca8042494ab3cc17f52878583652568f2da512620719a0000c0f0bd8239a1c13da9c2bbde9f2db7bdba8c2bad7bbcd08e15def4f96f79d546149cae31952dd5fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf55fbafd57eebf54b1f4065698538577bd3e5bdff5d091526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b526b4a4635def4f969960428912bb6fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb4fe2bb6fe28f0080200323ff83e80242dc1a6c9d0c0ae016028e7878203d800a30b9a69b2737d8b00330392ebfd8a4e10148369ded938064091c6127ff1d41411119132a029b0967a39ceb3e12ba9ff00c162ec7a3ff1f4799581998f33138513c161c1247c05753fed22228b61201872afd27d54acc493408ffc8530e69558fc1c4e7e02ba9ed858da0209a2cddc52b26c020732fd6969378ff7a6213b93eb7eb56f17bbad25c5c7d7c6b10d4fc10714df5bc766ea6ba5f3550a96bea3189060f1adf3b3750fe0cccb8d800958a4dafc0025aae5cbd6a609affdea130b7324dd88a28a96a84b15bf4d686d89808efef6f753651cff00d282096e921c5b9c9ab6418ccaf5351d76280ab018ad02b0a34cb773e2d3510746f4855d9bbaff00bd1342346f4951bc003437732e25a864930d8202c817c9066d06f96c828deb0e55306d7fef51c41b8b8dd88a534802128306c8def5be766ea6d2acb4119c03603bf95772eea21c2c04dea1b755367edd1e8341c08c90fa3252486c5a3bd9870f4a3afe5640e679650b4ca8bda8706a7fba5ae89b9c699f066c92a03c4dd61e27796d2ba9eca73b0aaf84a6f7039d2aaab2b8ad475bbe10ea9b1498c433fbe29480c7b937e31d8fdb3e571c56a9edc3f862ec7a369a1670ec5b382635a09292b6dbc9029f0c13239cd259360f49057060a6e96b1c04c21ae86f694d630ad8a5d9e1b164919c9b07587620a72c370f94873765a8406b2cc19d63db4c6503b37ba7507d6c62ae29bb343a45280ab018b4dd274fb418f13f5b0a2f42a0d7041c5a3ad53e8c329eb1449e6886ccc2710d9da356c52b2ccd86cbf63fda054025700a3ddd26db79205151e9491a84b35851e5b2b26f20dcfbf1f2da2ab924445fb8e5c3643c04276dda0ff95d8efecea7b26ea66dcda3da79d14d0c1aad8a8398215d0bafc6ed850f022f9437945bc71ac629d944319cae73269519003b9f1e2ec7a369a283784597ec4d841c8f1cd46141d5c21bf49a156400b01814081224251e0c6f0114a1e23f4ec2a616208709a926099f4198e94eee07595deefa7274d35e952150519042c2e386c662e48e9539e41a363df6055228c0601e7141041b0b0a58cc0253be24e46ced1aa9c389cef58ea94aacacae740e2e88b2c0ea4796d3fe503033751a760903519ac03e1c127cb28a82c0d4122389492aff005ce230f4d9654275bc9fa1ff0076753a9b88063d292b7152d5e554be8f010607d642d38562946fa7c78bb1e8da689cf0e365da35565465726f4db767d75dd746c046f3debb47693014bbcc0739a6d9a7f9a81527479acad3d1945023189c76189ebf15343ccdf4ec099c0a36cc67121d8768d54b1b36e11d8fda092d2692576dfcd76dfcd591c44f51a4bb1184a87a11e5b45462048fc8e5c1c29d7b40422626dea75022c12282185c7054a663eae1a39e79b3124d865812ae01486cf3d50740a1b146466ac1403e1e88078f1763d1b4d752d9768d5597831765d75dd746c622088ee13c0f7d90dc14885c91754a201b300f56342cf596e8d83dce942bcc39c47b0761084c5c25bf501cf6c28b388e0447aa5f4d9da3551852227017a0ec7c9020b11c9329aeedf9aeedf9a1856309ca776fcd04720aac34bf975155ad884381cdefc76f53d9306b020b099f109f5d8cb219fdd467ecab371132e822687d3c1ee2321d5d8a9a7912d961c31f4fe18bb1e8da6ba96cbb46aacbc18bb2ebaeeba36036c97a903dc760534a43041a8251a3c9b32f03169c9a4c4126eb7bf06c063899f41f9a0a48333121a5d972d9d8bd291b2244611a1e610114ef22fbca7a48810ed5055e952b622f634376ced1aaa120a75121a73c382f20e24545a166c1d47725aa3436c28772d928826b19f916c38b490208844e01c0b54351888466a4ac12d230300562eff2ea2a0e2c9f1cf8d068e4bb3e07136753d901a880e2689a25489e4c80eeb0e3878242966113a171e385072880c5d55cd75fe18bb1e8da6ba96cbb46aacbc08bb2ebaeeba3609245c8cf78b329cffc432b79428c8c26744026a02ead0ca122630968b5e31d91fc4c2d4d804002ce760f4a32698193bf303b13146e8e62e01503c142c3d03df60bbdd2d89d45ac8ec21ca923f642778e66f36a6fb8263fe55a828ae340ebabe551d7a2c8344d854d5482a4d2de089388a834bec2a7ec4142eb6f03f56c655ce14b5a8d10564e046a3aa8008080f1b62e921e23b0a32c080300ab125a6332e6c2a3464224d2db0c30903823b0a6d5d24191676baed212ae2b49a6c04c260edbd4760503c97a52c06882841be73aa9690490690388ec28117880c071b9461152c33826a2853703a2d4cff19e1d11469690302371c2889e6051e852ec5873424bd4d8e9168c7583ef6bfd3763ad270874af49a1ea0e4b3deb502a1cb8b9ff00e5cc29044dd4dc85a698b74713606c75a9b01b2315312686d742a6701031a40a4bad00955c0a416163719f4c397fe78292781ca6e4cca8300c2e0715e9320343f153046837c51f3dc883d6f575f944bba938bbdfff002e109001c4e92ec2940048641adb8ecb2e641097025d85230650ea3583f9146880e174bec29a4d20c835b78d0a42310b74ec280ac09081d63ff1bb16770d363bc6bff1a648ed3a3676ed5ff895a6d0310131e7b28334d1352561cb6770d363bc6bff001a648ed3a3676ed5e6475324ab3c217694853302fab35732bbb2a2c00e404eb4d02495f54afb06c49688ae131b2c5919e2c182625f00ccb2104dcb9b2cbf60944d3189d92e8f03f40c5a5b0ec80beacf4a7011776543033550f52b0b0a5b37363cb69e2e05d4a5c0e1b2c646c4388218bc4d9dc34d8ef1aed2eca1813c0c5e54ce0d916f54f4a7108dff00551607cdf9128898b029f80e3cbc53245e7b247e81f3498ded4f2bd169e6700f4bf4ad6cfd2373a3c76f69d1b3b76ada26d6400e6d1206b7b961ed462f7cbded4d09da67f54a8948c809ccf2a86328fc0735b8f78a54692d95d926bb1d155908d40416558516e63a8d0805f13c45ea14e9167648c2524b1bf8870a9b331dc93b06a2aac80969346ddd09b1c882992086cd5828122da19e25eab4708992e3ebb8c5a50592a58340c8ddb24d76248393856a3446a09aee1c5647c5b16770d363bc6bb185db7175c28b54312c68323649aec46c890308d2a01d78e05f78c5f7f826397330a4a366785274aba95c76b9f5375b5a0a48730371c56ced3a3676ed5b235a7031590ded39645c5bf977ec93631645c5bdda3beae2d087159ade794c4248ca0c564e6ad30907990f58a06264401b806d3ab4bdd2d0c1ed5acf9324501960ec94182b790d8b73c0c601b0f679ecbf4c00f13aa4e5b2dcb1718c6ebd8e7b2fe721382ec3dde5b08ac3924eea1781caa5a08c581eb85079dbf327bcb8de991ae08b1ed478522c5a06324630d76085b3a6f11fbf16c59dc34d8ef1aec1e231092b60b91eed484399927a96a024d60d831233575a021c9e03daa2c484245864c24999d8666c86fb9f73c13196067c52c050b9c12cbd12dce933e1110270424a1551a0a5d5398e4ecbbcc826c048fa49b3b4e8d9dbb56cb7ce246c29e841cdd8710e8920863b89caa35118c27b5428a332660ca33365fb705700f927d0f2a48580c570a5514c7a7a54521fc227736fd28496001552625e7a6ceabb2bf6ab05d0c1e97e5b2cea5b996ebd6dcf65f06d2c33af5b72d97c5374b9807a5f9ed501588ce6a782328f40968d42f85eb9c3da9bc08c93c325ec1869b3bae9e2d8b3b869b1de35f027463387d05e8a42785b3ce1538cc65cdd8075d92b3e0e991480c780368f62aeeecbdb63890a394276769d1b3b76aa5825a576595dd28e9b25d99d1641396b35fb4faa2d264ab891a514090a0ea5319121e0b0f45f299d9915b3d90674ddcb6887016d8991ee6fab0a26da08256722d96bb0ded88c1090398d318442b9abf43d29e6040c92e5383104a7047c14a217766acb43f3064646ebd0682e0c064186c26b32f44fb014ac36d9e2e463ce760627b354f5c2ae2b41972816275cf619ddfb678b62cee1a6c778d7623bce5687e00cda9df0d91334b5de74e32e34394bae3d6a38940a66e09ebb04f8029823b185b4637a7eba30a0a66289e8a302bb4e8d9dbb554ed632f6a9ddc65efb21d1a8746a1d1a8746982190bd69cbb98f275015c0a5c4e3391607a05200bab05345e4741641c028000001914897013826c17fb5bb71c965355d7ac9cf6067ca03c331b11e25be9cd7b1ebb509293d163de5e7b20bbe51224036e74084180101488611f7b07e3676ddde2d8b3b869b1de35d86aa61c8104af55d82f10664b083098bcd072ac020298d44078a0d827c02982e10ab0210758da350696b67da5e5b01c99af0197a1b3b4e8d9dbb552490d19f0f259474d82c2498e310f59afcc522011d0062b428cbe03df4a14938e0f28df609420a033a8a50a84612e50ec8844a6a623451f1242f005dab90a98e4e2b7ba65b0d2c5a396d8e72485e43c98a6d8bbb2461d8fa80066ac05146498ccddd7aaed311316d149d1360d6f330210238db12852026c4bc0c5a258e218c1c8c8d0dfb01070ef49f8f16c59dc34d8ef1aec18498f66013df648bd0e91c2da4c229018930cfda82145a84308321d761bdfb42f04c1864c6944065906512f3bca741c122c728a2957646ebee7b6c6aa29131714e07beced3a3676ed5b2c41c50b0b3e247a3b1d9c79251c6d9aa3763cd84726f58682597b84e6d1664c68d2db80464898bbd07ca547924f2865331cfd76e2efab42da562403237ec9f3255c02d3ac78307783600b7c1e7b2ea37160e13eb7e5e09182c1b2c019a164d29114484c4763be8c814dd1c5819939e14174b039259a64e009cc7468441191c1f0ec59dc34d8ef1aec5811271c5b6f2fc69cb2a0108ec6f8d47c975adc737a513d1c4818a73f3ee037e9340480066248ed9861583c2524b73052dfbc0b26a399bf68e5920b0d353ba88ac119baaeabb3b4e8d9dbb56cb2ab1d5649a254ceb7321a6a78020b36105aafc635706928875786879505354c9177859f4a2fe905c44c341c26c722a088ca9a21156717ad08051912c75a000302db62a60e7744655df1f34db046edc618f87168c2c9c62cf3a5e60d001bdc1687836e94000000b014c397384abae35df1f35052009980b14709c976219db8d77c7cd351eb765595ddc6cb3d64e624c1b71aef8f9a74fc9bca1c4bbbb6bb69e63cf4c79d4f77915092ee083f4c3a51d6280200a99260215ad9aef8f9a83224d581817a323d2ec033b716bbe3e6810041956cc3177b5a758244e0e272a4451c92742d081185ff4105119f81806e0daefcc2d24e3835df1f347a8c56b030c5da185a01226f2b1a8057d1874a548334a054eb553d082b0d041c1c8ff00f8c6ffd9";
            Drawing drawing = null;
            wBorder border = new wBorder() { Val = BorderValues.None };
            string src = en.Attributes["src"];
            Uri uri = null;

            // Bug reported by Erik2014. Inline 64 bit images can be too big and Uri.TryCreate will fail silently with a SizeLimit error.
            // To circumvent this buffer size, we will work either on the Uri, either on the original src.
            if (src != null && (IO.DataUri.IsWellFormed(src) || Uri.TryCreate(src, UriKind.RelativeOrAbsolute, out uri)))
            {
                string alt = (en.Attributes["title"] ?? en.Attributes["alt"]) ?? String.Empty; // NAME OF THE IMAGE DO IT LATER

                Size preferredSize = Size.Empty;
                Unit wu = en.Attributes.GetAsUnit("width");
                if (!wu.IsValid) wu = en.StyleAttributes.GetAsUnit("width");
                Unit hu = en.Attributes.GetAsUnit("height");
                if (!hu.IsValid) hu = en.StyleAttributes.GetAsUnit("height");

                // % is not supported
                if (wu.IsFixed && wu.Value > 0)
                {
                    preferredSize.Width = wu.ValueInPx;
                }
                if (hu.IsFixed && hu.Value > 0)
                {
                    // Image perspective skewed. Bug fixed by ddeforge on github.com/onizet/html2openxml/discussions/350500
                    preferredSize.Height = hu.ValueInPx;
                }

                SideBorder attrBorder = en.StyleAttributes.GetAsSideBorder("border");
                if (attrBorder.IsValid)
                {
                    border.Val = attrBorder.Style;
                    border.Color = attrBorder.Color.ToHexString();
                    border.Size = (uint)attrBorder.Width.ValueInPx * 4;
                }
                else
                {
                    var attrBorderWidth = en.Attributes.GetAsUnit("border");
                    if (attrBorderWidth.IsValid)
                    {
                        border.Val = BorderValues.Single;
                        border.Size = (uint)attrBorderWidth.ValueInPx * 4;
                    }
                }

                Stream imageStream;
                ImagePart imagePart;
                string extension = Path.GetExtension(uri.ToString());

                switch (extension)
                {
                    case ".bmp":
                        imagePart = mainPart.AddImagePart(ImagePartType.Bmp);
                        break;
                    case ".gif":
                        imagePart = mainPart.AddImagePart(ImagePartType.Gif);
                        break;
                    case ".jpg":
                    case ".jpeg":
                        imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                        break;
                    case ".png":
                        imagePart = mainPart.AddImagePart(ImagePartType.Png);
                        break;
                    case ".svg":
                        imagePart = mainPart.AddImagePart(ImagePartType.Svg);
                        break;
                    default:
                        throw new ArgumentException("Unsupported image format");
                }

                HtmlImageInfo info = new HtmlImageInfo() { Source = src };

                try
                {
                    string image = storedImages.Where(x => x.src == "images/" + src).FirstOrDefault().hex;
                    if (image != null)
                    {
                        if (image.Length % 2 != 0)
                        {
                            image = "0" + image;
                        }
                        byte[] bytes = Enumerable.Range(0, image.Length)
                                        .Where(x => x % 2 == 0)
                                        .Select(x => Convert.ToByte(image.Substring(x, 2), 16))
                                        .ToArray();
                        //byte[] bytes = Encoding.UTF8.GetBytes(image);
                        imageStream = new MemoryStream(bytes);
                        imagePart.FeedData(imageStream);

                        imageStream.Seek(0L, SeekOrigin.Begin);
                        info.Size = ImageHeader.GetDimensions(imageStream);


                        drawing = AddImageToBody(mainPart.GetIdOfPart(imagePart), info.Size, preferredSize, mainPart, info.ImagePartId, alt, uri);
                    }
                    else
                    {
                        throw new Exception("Image wasn't found");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);

                    byte[] bytes = Enumerable.Range(0, defaultImage.Length)
                                    .Where(x => x % 2 == 0)
                                    .Select(x => Convert.ToByte(defaultImage.Substring(x, 2), 16))
                                    .ToArray();

                    imageStream = new MemoryStream(bytes);
                    imagePart.FeedData(imageStream);

                    imageStream.Seek(0L, SeekOrigin.Begin);
                    info.Size = new Size(32, 32);


                    drawing = AddImageToBody(mainPart.GetIdOfPart(imagePart), info.Size, preferredSize, mainPart, info.ImagePartId, alt, uri);
                }
            }

            if (drawing != null)
            {
                Run run = new Run(drawing);
                if (border.Val != BorderValues.None) run.InsertInProperties(prop => prop.Border = border);
                elements.Add(run);
            }
        }

        #endregion

        #region ProcessInput

        private void ProcessInput(HtmlEnumerator en)
        {

            if (en.Current.Contains("checked=\"checked\""))
            {
                en.IsChecked = true;
            }
        }

        #endregion

        #region ProcessLi

        private void ProcessLi(HtmlEnumerator en)
        {
            CompleteCurrentParagraph(false);
            currentParagraph = htmlStyles.Paragraph.NewParagraph();

            int numberingId = htmlStyles.NumberingList.ProcessItem(en);
            int level = htmlStyles.NumberingList.LevelIndex;
            bool isCheckBox = false;

            // Save the new paragraph reference to support nested numbering list.
            Paragraph p = currentParagraph;
            if (en.NextTag == "<input>")
            {
                isCheckBox = true;
            }
            else
            {
                currentParagraph.InsertInProperties(prop =>
                {
                    prop.ParagraphStyleId = new ParagraphStyleId() { Val = GetStyleIdForListItem(en) };
                    prop.Indentation = level < 2 ? null : new Indentation() { Left = (level * 780).ToString(CultureInfo.InvariantCulture) };
                    prop.NumberingProperties = new NumberingProperties
                    {
                        NumberingLevelReference = new NumberingLevelReference() { Val = level - 1 },
                        NumberingId = new NumberingId() { Val = numberingId }
                    };
                });
            }

            // Restore the original elements list
            AddParagraph(currentParagraph);

            // Continue to process the html until we found </li>
            HtmlStyles.Paragraph.ApplyTags(currentParagraph);
            AlternateProcessHtmlChunks(en, "</li>");
            if (isCheckBox)
            {
                SdtRun sdt = new SdtRun();
                SdtProperties sdtPr = new SdtProperties();
                w14.SdtContentCheckBox checkbox = new w14.SdtContentCheckBox();
                w14.Checked checkedd = new w14.Checked();
                Text text = new Text();

                if (en.IsChecked)
                {
                    checkedd.Val = w14.OnOffValues.One;
                    text.Text = "☒";
                    en.IsChecked = false;
                }
                else
                {
                    checkedd.Val = w14.OnOffValues.Zero;
                    text.Text = "☐";
                }

                w14.CheckedState checkedState = new w14.CheckedState() { Font = "MS Gothic", Val = "2612" };
                w14.UncheckedState uncheckedState = new w14.UncheckedState() { Font = "MS Gothic", Val = "2610" };

                sdt.Append(sdtPr);

                sdtPr.Append(checkbox);
                checkbox.Append(checkedd);
                checkbox.Append(checkedState);
                checkbox.Append(uncheckedState);

                SdtContentRun sdtContentRun = new SdtContentRun();
                Run run = new Run();


                run.Append(text);
                sdtContentRun.Append(run);
                sdt.Append(sdtContentRun);
                p.AppendChild(sdt);

                p.Append(elements);

            }
            else
            {
                p.Append(elements);
            }

            this.elements.Clear();
        }

        private string GetStyleIdForListItem(HtmlEnumerator en)
        {
            return GetStyleIdFromClasses(en.Attributes.GetAsClass())
                   ?? GetStyleIdFromClasses(htmlStyles.NumberingList.GetCurrentListClasses)
                   ?? htmlStyles.DefaultStyles.ListParagraphStyle;
        }

        private string GetStyleIdFromClasses(string[] classes)
        {
            if (classes != null)
            {
                foreach (string className in classes)
                {
                    string styleId = htmlStyles.GetStyle(className, StyleValues.Paragraph, ignoreCase: true);
                    if (styleId != null)
                    {
                        return styleId;
                    }
                }
            }

            return null;
        }

        #endregion

        #region ProcessLink

        private void ProcessLink(HtmlEnumerator en)
        {
            String att = en.Attributes["href"];
            Hyperlink h = null;
            Uri uri = null;


            if (!String.IsNullOrEmpty(att))
            {
                // handle link where the http:// is missing and that starts directly with www
                if (att.StartsWith("www.", StringComparison.OrdinalIgnoreCase))
                    att = "http://" + att;

                // is it an anchor?
                if (att[0] == '#' && att.Length > 1)
                {
                    // Always accept _top anchor
                    if (!this.ExcludeLinkAnchor || att == "#_top")
                    {
                        h = new Hyperlink(
                            )
                        { History = true, Anchor = att.Substring(1) };
                    }
                }
                // ensure the links does not start with javascript:
                else if (Uri.TryCreate(att, UriKind.Absolute, out uri) && uri.Scheme != "javascript")
                {
                    HyperlinkRelationship extLink = mainPart.AddHyperlinkRelationship(uri, true);

                    h = new Hyperlink(
                        )
                    { History = true, Id = extLink.Id };
                }
            }

            if (h == null)
            {
                // link to a broken url, simply process the content of the tag
                ProcessHtmlChunks(en, "</a>");
                return;
            }

            att = en.Attributes["title"];
            if (!String.IsNullOrEmpty(att)) h.Tooltip = att;

            AlternateProcessHtmlChunks(en, "</a>");

            if (elements.Count == 0) return;

            // Let's see whether the link tag include an image inside its body.
            // If so, the Hyperlink OpenXmlElement is lost and we'll keep only the images
            // and applied a HyperlinkOnClick attribute.
            List<OpenXmlElement> imageInLink = elements.FindAll(e => { return e.HasChild<Drawing>(); });
            if (imageInLink.Count != 0)
            {
                for (int i = 0; i < imageInLink.Count; i++)
                {
                    // Retrieves the "alt" attribute of the image and apply it as the link's tooltip
                    Drawing d = imageInLink[i].GetFirstChild<Drawing>();
                    var enDp = d.Descendants<pic.NonVisualDrawingProperties>().GetEnumerator();
                    String alt;
                    if (enDp.MoveNext()) alt = enDp.Current.Description;
                    else alt = null;

                    d.InsertInDocProperties(
                            new a.HyperlinkOnClick() { Id = h.Id ?? h.Anchor, Tooltip = alt });
                }
            }

            // Append the processed elements and put them to the Run of the Hyperlink
            h.Append(elements);

            // can't use GetFirstChild<Run> or we may find the one containing the image
            foreach (var el in h.ChildElements)
            {
                Run run = el as Run;
                if (run != null && !run.HasChild<Drawing>())
                {
                    run.InsertInProperties(prop =>
                        prop.RunStyle = new RunStyle() { Val = htmlStyles.GetStyle(htmlStyles.DefaultStyles.HyperlinkStyle, StyleValues.Character) });
                    break;
                }
            }

            this.elements.Clear();

            // Append the hyperlink
            elements.Add(h);

            if (imageInLink.Count > 0) CompleteCurrentParagraph(true);
        }

        #endregion

        #region ProcessNumberingList

        private void ProcessNumberingList(HtmlEnumerator en)
        {
            htmlStyles.NumberingList.BeginList(en);
        }

        #endregion

        #region ProcessCode

        private void ProcessCode(HtmlEnumerator en)
        {
            currentParagraph.AppendChild(new ParagraphProperties(
                       new Indentation() { Left = "500", Right = "500" },
                       new ParagraphBorders() { InnerXml = "<w:top w:val=\"single\" w:sz=\"6\" w:space=\"7\" w:color=\"CCCCCC\"/>\r\n<w:left w:val=\"single\" w:sz=\"6\" w:space=\"7\" w:color=\"CCCCCC\"/>\r\n<w:bottom w:val=\"single\" w:sz=\"6\" w:space=\"7\" w:color=\"CCCCCC\"/>\r\n<w:right w:val=\"single\" w:sz=\"6\" w:space=\"7\" w:color=\"CCCCCC\"/>" },
                       new Tabs() { InnerXml = "<w:tab w:val=\"left\" w:pos=\"916\"/>\r\n<w:tab w:val=\"left\" w:pos=\"1832\"/>\r\n<w:tab w:val=\"left\" w:pos=\"2748\"/>\r\n<w:tab w:val=\"left\" w:pos=\"3664\"/>\r\n<w:tab w:val=\"left\" w:pos=\"4580\"/>\r\n<w:tab w:val=\"left\" w:pos=\"5496\"/>\r\n<w:tab w:val=\"left\" w:pos=\"6412\"/>\r\n<w:tab w:val=\"left\" w:pos=\"7328\"/>\r\n<w:tab w:val=\"left\" w:pos=\"8244\"/>\r\n<w:tab w:val=\"left\" w:pos=\"9160\"/>\r\n<w:tab w:val=\"left\" w:pos=\"10076\"/>\r\n<w:tab w:val=\"left\" w:pos=\"10992\"/>\r\n<w:tab w:val=\"left\" w:pos=\"11908\"/>\r\n<w:tab w:val=\"left\" w:pos=\"12824\"/>\r\n<w:tab w:val=\"left\" w:pos=\"13740\"/>\r\n<w:tab w:val=\"left\" w:pos=\"14656\"/>" },
                       new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F8F8F8" },
                       new SpacingBetweenLines() { BeforeLines = 250 }
                       )
                   );

        }


        #endregion

        #region ProcessParagraph

        private void ProcessParagraph(HtmlEnumerator en)
        {
            CompleteCurrentParagraph(true);

            if (en.getParent() == "<blockquote>")
            {
                currentParagraph.AppendChild(new ParagraphProperties(
                    new Indentation() { Left = "500", Right = "500" },
                    new ParagraphBorders() { LeftBorder = new LeftBorder() { Size = 24, Space = 15, Color = "0000FF", Val = BorderValues.Single } }

                    ,
                    new SpacingBetweenLines() { BeforeLines = 250 }
                    )
                );

            }

            else
            {
                //add the spacing
                currentParagraph.AppendChild(new ParagraphProperties(

                    new SpacingBetweenLines() { BeforeLines = 250 }
                    )
                );
            }

            // Respect this order: this is the way the browsers apply them
            String attrValue = en.StyleAttributes["text-align"];
            if (attrValue == null) attrValue = en.Attributes["align"];

            if (attrValue != null)
            {
                JustificationValues? align = Converter.ToParagraphAlign(attrValue);
                if (align.HasValue)
                {
                    currentParagraph.InsertInProperties(prop => prop.Justification = new Justification { Val = align });
                }
            }

            List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();

            bool newParagraph = ProcessContainerAttributes(en, styleAttributes);



            if (styleAttributes.Count > 0)
                htmlStyles.Runs.BeginTag(en.CurrentTag, styleAttributes.ToArray());

            if (newParagraph)
            {
                AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);
                ProcessClosingParagraph(en);
            }

        }

        #endregion

        #region ProcessPre

        private void ProcessPre(HtmlEnumerator en)
        {
            CompleteCurrentParagraph();
            currentParagraph = htmlStyles.Paragraph.NewParagraph();

            // Oftenly, <pre> tag are used to renders some code examples. They look better inside a table
            /*       if (en.CurrentTag == "<pre>" && en.NextTag == "<code>")
                   {
                       StyleDefinitionsPart part = mainPart.StyleDefinitionsPart;
                       var styleid = "PlainTable43";
                       var stylename = "Plain Table 43";

                       if (part == null)
                       {
                           part = TableStyleCollection.AddStylesPartToPackage(mainPart);
                           TableStyleCollection.AddNewTableStyle(part, styleid, stylename);
                       }
                       else
                       {
                           // If the style is not in the document, add it.
                           if (TableStyleCollection.IsStyleIdInDocument(part, styleid) != true)
                           {
                               // No match on styleid, so let's try style name.
                               string styleidFromName = TableStyleCollection.GetStyleIdFromStyleName(mainPart, stylename);
                               if (styleidFromName == null)
                               {
                                   TableStyleCollection.AddNewTableStyle(part, styleid, stylename);
                               }
                               else
                                   styleid = styleidFromName;
                           }
                       }

                       Table currentTable = new Table(
                           new TableProperties (
                               new TableStyle() { Val = "PlainTable43" },
                               new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct } // 100% * 50
                           ),
                           new TableGrid(
                               new GridColumn() { Width = "5610" })
                       );
                       AddParagraph(currentTable);

                       AlternateProcessHtmlChunks(en, "</pre>");

                       foreach (OpenXmlElement element in elements)
                       {
                           if (element.InnerText != "")
                           {
                               var tr = currentTable.AppendChild(new TableRow());
                               var tc = tr.AppendChild(new TableCell());
                               var p = tc.AppendChild(new Paragraph());
                               p.Append(element);
                           }
                       }
                       elements.Clear();
                   }
                   else
                   {
                       AddParagraph(currentParagraph);
                   }
       */
            AddParagraph(currentParagraph);


            //Probably I'll need the loop here, when the elements are done

            //CompleteCurrentParagraph();
        }

        #endregion

        #region ProcessQuote

        private void ProcessQuote(HtmlEnumerator en)
        {
            // The browsers render the quote tag between a kind of separators.
            // We add the Quote style to the nested runs to match more Word.

            Run run = new Run(
                new Text(" " + HtmlStyles.QuoteCharacters.Prefix) { Space = SpaceProcessingModeValues.Preserve }
            );

            htmlStyles.Runs.ApplyTags(run);
            elements.Add(run);

            ProcessHtmlElement<RunStyle>(en, new RunStyle() { Val = htmlStyles.GetStyle(htmlStyles.DefaultStyles.QuoteStyle, StyleValues.Character) });
        }

        #endregion

        #region ProcessSpan

        private void ProcessSpan(HtmlEnumerator en)
        {
            // A span style attribute can contains many information: font color, background color, font size,
            // font family, ...
            // We'll check for each of these and add apply them to the next build runs.

            List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
            bool newParagraph = ProcessContainerAttributes(en, styleAttributes);

            if (styleAttributes.Count > 0)
                htmlStyles.Runs.MergeTag(en.CurrentTag, styleAttributes);

            if (newParagraph)
            {
                AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);
                CompleteCurrentParagraph(true);
            }
        }

        #endregion

        #region ProcessSubscript

        private void ProcessSubscript(HtmlEnumerator en)
        {
            ProcessHtmlElement<VerticalTextAlignment>(en, new VerticalTextAlignment() { Val = VerticalPositionValues.Subscript });
        }

        #endregion

        #region ProcessSuperscript

        private void ProcessSuperscript(HtmlEnumerator en)
        {
            ProcessHtmlElement<VerticalTextAlignment>(en, new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript });
        }

        #endregion

        #region ProcessUnderline

        private void ProcessUnderline(HtmlEnumerator en)
        {
            ProcessHtmlElement<Underline>(en, new Underline() { Val = UnderlineValues.Single });
        }

        #endregion

        #region ProcessTable

        private void ProcessTable(HtmlEnumerator en)
        {
            TableProperties properties = new TableProperties(
                new TableStyle() { Val = htmlStyles.GetStyle(htmlStyles.DefaultStyles.TableStyle, StyleValues.Table) }
            );
            Table currentTable = new Table(properties);

            string classValue = en.Attributes["class"];
            if (classValue != null)
            {
                classValue = htmlStyles.GetStyle(classValue, StyleValues.Table, ignoreCase: true);
                if (classValue != null)
                    properties.TableStyle.Val = classValue;
            }

            int? border = en.Attributes.GetAsInt("border");
            if (border.HasValue && border.Value > 0)
            {
                bool handleBorders = true;
                if (classValue != null)
                {
                    // check whether the style in use have borders
                    String styleId = this.htmlStyles.GetStyle(classValue, StyleValues.Table, true);
                    if (styleId != null)
                    {
                        var s = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().First(e => e.StyleId == styleId);
                        if (s.StyleTableProperties.TableBorders != null) handleBorders = false;
                    }
                }

                // If the border has been specified, we display the Table Grid style which display
                // its grid lines. Otherwise the default table style hides the grid lines.
                if (handleBorders && properties.TableStyle.Val != htmlStyles.DefaultStyles.TableStyle)
                {
                    uint borderSize = border.Value > 1 ? (uint)new Unit(UnitMetric.Pixel, border.Value).ValueInDxa : 1;
                    properties.TableBorders = new TableBorders()
                    {
                        TopBorder = new TopBorder { Val = BorderValues.None },
                        LeftBorder = new LeftBorder { Val = BorderValues.None },
                        RightBorder = new RightBorder { Val = BorderValues.None },
                        BottomBorder = new BottomBorder { Val = BorderValues.None },
                        InsideHorizontalBorder = new InsideHorizontalBorder { Val = BorderValues.Single, Size = borderSize },
                        InsideVerticalBorder = new InsideVerticalBorder { Val = BorderValues.Single, Size = borderSize }
                    };
                }
            }
            // is the border=0? If so, we remove the border regardless the style in use
            else if (border == 0)
            {
                properties.TableBorders = new TableBorders()
                {
                    TopBorder = new TopBorder { Val = BorderValues.None },
                    LeftBorder = new LeftBorder { Val = BorderValues.None },
                    RightBorder = new RightBorder { Val = BorderValues.None },
                    BottomBorder = new BottomBorder { Val = BorderValues.None },
                    InsideHorizontalBorder = new InsideHorizontalBorder { Val = BorderValues.None },
                    InsideVerticalBorder = new InsideVerticalBorder { Val = BorderValues.None }
                };
            }
            else
            {
                var styleBorder = en.StyleAttributes.GetAsBorder("border");
                if (!styleBorder.IsEmpty)
                {
                    properties.TableBorders = new TableBorders();

                    if (styleBorder.Left.IsValid)
                        properties.TableBorders.LeftBorder = new LeftBorder { Val = styleBorder.Left.Style, Color = StringValue.FromString(styleBorder.Left.Color.ToHexString()), Size = (uint)styleBorder.Left.Width.ValueInDxa };
                    if (styleBorder.Right.IsValid)
                        properties.TableBorders.RightBorder = new RightBorder { Val = styleBorder.Right.Style, Color = StringValue.FromString(styleBorder.Right.Color.ToHexString()), Size = (uint)styleBorder.Right.Width.ValueInDxa };
                    if (styleBorder.Top.IsValid)
                        properties.TableBorders.TopBorder = new TopBorder { Val = styleBorder.Top.Style, Color = StringValue.FromString(styleBorder.Top.Color.ToHexString()), Size = (uint)styleBorder.Top.Width.ValueInDxa };
                    if (styleBorder.Bottom.IsValid)
                        properties.TableBorders.BottomBorder = new BottomBorder { Val = styleBorder.Bottom.Style, Color = StringValue.FromString(styleBorder.Bottom.Color.ToHexString()), Size = (uint)styleBorder.Bottom.Width.ValueInDxa };
                }
            }

            Unit unit = en.StyleAttributes.GetAsUnit("width");
            if (!unit.IsValid) unit = en.Attributes.GetAsUnit("width");

            if (unit.IsValid)
            {
                switch (unit.Type)
                {
                    case UnitMetric.Percent:
                        properties.TableWidth = new TableWidth() { Type = TableWidthUnitValues.Pct, Width = (unit.Value * 50).ToString(CultureInfo.InvariantCulture) }; break;
                    case UnitMetric.Point:
                        properties.TableWidth = new TableWidth() { Type = TableWidthUnitValues.Dxa, Width = unit.ValueInDxa.ToString(CultureInfo.InvariantCulture) }; break;
                    case UnitMetric.Pixel:
                        properties.TableWidth = new TableWidth() { Type = TableWidthUnitValues.Dxa, Width = unit.ValueInDxa.ToString(CultureInfo.InvariantCulture) }; break;
                }
            }
            else
            {
                // Use Auto=0 instead of Pct=auto
                // bug reported by scarhand (https://html2openxml.codeplex.com/workitem/12494)
                properties.TableWidth = new TableWidth() { Type = TableWidthUnitValues.Auto, Width = "0" };
            }

            string align = en.Attributes["align"];
            if (align != null)
            {
                JustificationValues? halign = Converter.ToParagraphAlign(align);
                if (halign.HasValue)
                    properties.TableJustification = new TableJustification() { Val = halign.Value.ToTableRowAlignment() };
            }

            // only if the table is left aligned, we can handle some left margin indentation
            // Right margin + Right align has no equivalent in OpenXml
            if (align == null || align == "left")
            {
                Margin margin = en.StyleAttributes.GetAsMargin("margin");

                // OpenXml doesn't support table margin in Percent, but Html does
                // the margin part has been implemented by Olek (patch #8457)

                TableCellMarginDefault cellMargin = new TableCellMarginDefault();
                if (margin.Left.IsFixed)
                    cellMargin.TableCellLeftMargin = new TableCellLeftMargin() { Type = TableWidthValues.Dxa, Width = (short)margin.Left.ValueInDxa };
                if (margin.Right.IsFixed)
                    cellMargin.TableCellRightMargin = new TableCellRightMargin() { Type = TableWidthValues.Dxa, Width = (short)margin.Right.ValueInDxa };
                if (margin.Top.IsFixed)
                    cellMargin.TopMargin = new TopMargin() { Type = TableWidthUnitValues.Dxa, Width = margin.Top.ValueInDxa.ToString(CultureInfo.InvariantCulture) };
                if (margin.Bottom.IsFixed)
                    cellMargin.BottomMargin = new BottomMargin() { Type = TableWidthUnitValues.Dxa, Width = margin.Bottom.ValueInDxa.ToString(CultureInfo.InvariantCulture) };

                // Align table according to the margin 'auto' as it stands in Html
                if (margin.Left.Type == UnitMetric.Auto || margin.Right.Type == UnitMetric.Auto)
                {
                    TableRowAlignmentValues justification;

                    if (margin.Left.Type == UnitMetric.Auto && margin.Right.Type == UnitMetric.Auto)
                        justification = TableRowAlignmentValues.Center;
                    else if (margin.Left.Type == UnitMetric.Auto)
                        justification = TableRowAlignmentValues.Right;
                    else
                        justification = TableRowAlignmentValues.Left;

                    properties.TableJustification = new TableJustification() { Val = justification };
                }

                if (cellMargin.HasChildren)
                    properties.TableCellMarginDefault = cellMargin;
            }

            int? spacing = en.Attributes.GetAsInt("cellspacing");
            if (spacing.HasValue)
                properties.TableCellSpacing = new TableCellSpacing { Type = TableWidthUnitValues.Dxa, Width = new Unit(UnitMetric.Pixel, spacing.Value).ValueInDxa.ToString(CultureInfo.InvariantCulture) };

            int? padding = en.Attributes.GetAsInt("cellpadding");
            if (padding.HasValue)
            {
                int paddingDxa = (int)new Unit(UnitMetric.Pixel, padding.Value).ValueInDxa;

                TableCellMarginDefault cellMargin = new TableCellMarginDefault();
                cellMargin.TableCellLeftMargin = new TableCellLeftMargin() { Type = TableWidthValues.Dxa, Width = (short)paddingDxa };
                cellMargin.TableCellRightMargin = new TableCellRightMargin() { Type = TableWidthValues.Dxa, Width = (short)paddingDxa };
                cellMargin.TopMargin = new TopMargin() { Type = TableWidthUnitValues.Dxa, Width = paddingDxa.ToString(CultureInfo.InvariantCulture) };
                cellMargin.BottomMargin = new BottomMargin() { Type = TableWidthUnitValues.Dxa, Width = paddingDxa.ToString(CultureInfo.InvariantCulture) };
                properties.TableCellMarginDefault = cellMargin;
            }

            List<OpenXmlElement> runStyleAttributes = new List<OpenXmlElement>();
            htmlStyles.Tables.ProcessCommonAttributes(en, runStyleAttributes);
            if (runStyleAttributes.Count > 0)
                htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes.ToArray());


            // are we currently inside another table?
            if (tables.HasContext)
            {
                // Okay we will insert nested table but beware the paragraph inside TableCell should contains at least 1 run.

                TableCell currentCell = tables.CurrentTable.GetLastChild<TableRow>().GetLastChild<TableCell>();
                // don't add an empty paragraph if not required (bug #13608 by zanjo)
                if (elements.Count == 0) currentCell.Append(currentTable);
                else
                {
                    currentCell.Append(new Paragraph(elements), currentTable);
                    elements.Clear();
                }
            }
            else
            {
                CompleteCurrentParagraph();
                this.paragraphs.Add(currentTable);
            }

            tables.NewContext(currentTable);
        }

        #endregion

        #region ProcessTableCaption

        private void ProcessTableCaption(HtmlEnumerator en)
        {
            if (!tables.HasContext) return;

            string att = en.StyleAttributes["text-align"];
            if (att == null) att = en.Attributes["align"];

            ProcessHtmlChunks(en, "</caption>");

            var legend = new Paragraph(
                    new ParagraphProperties
                    {
                        ParagraphStyleId = new ParagraphStyleId() { Val = htmlStyles.GetStyle(htmlStyles.DefaultStyles.CaptionStyle, StyleValues.Paragraph) }
                    },
                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.Begin }),
                    new Run(
                        new FieldCode(" SEQ TABLE \\* ARABIC ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.End })
                );
            legend.Append(elements);
            elements.Clear();

            if (att != null)
            {
                JustificationValues? align = Converter.ToParagraphAlign(att);
                if (align.HasValue)
                    legend.InsertInProperties(prop => prop.Justification = new Justification { Val = align });
            }
            else
            {
                // If no particular alignement has been specified for the legend, we will align the legend
                // relative to the owning table
                TableProperties props = tables.CurrentTable.GetFirstChild<TableProperties>();
                if (props != null)
                {
                    TableJustification justif = props.GetFirstChild<TableJustification>();
                    if (justif != null) legend.InsertInProperties(prop =>
                        prop.Justification = new Justification { Val = justif.Val.Value.ToJustification() });
                }
            }

            if (this.TableCaptionPosition == CaptionPositionValues.Above)
                this.paragraphs.Insert(this.paragraphs.Count - 1, legend);
            else
                this.paragraphs.Add(legend);
        }

        #endregion

        #region ProcessTableRow

        private void ProcessTableRow(HtmlEnumerator en)
        {
            // in case the html is bad-formed and use <tr> outside a <table> tag, we will ensure
            // a table context exists.
            if (!tables.HasContext) return;

            TableRowProperties properties = new TableRowProperties();
            List<OpenXmlElement> runStyleAttributes = new List<OpenXmlElement>();

            htmlStyles.Tables.ProcessCommonAttributes(en, runStyleAttributes);

            Unit unit = en.StyleAttributes.GetAsUnit("height");
            if (!unit.IsValid) unit = en.Attributes.GetAsUnit("height");

            switch (unit.Type)
            {
                case UnitMetric.Point:
                    properties.AddChild(new TableRowHeight() { HeightType = HeightRuleValues.AtLeast, Val = (uint)(unit.Value * 20) });
                    break;
                case UnitMetric.Pixel:
                    properties.AddChild(new TableRowHeight() { HeightType = HeightRuleValues.AtLeast, Val = (uint)unit.ValueInDxa });
                    break;
            }

            // Do not explicitly set the tablecell spacing in order to inherit table style (issue 107)
            //properties.AddChild(new TableCellSpacing() { Type = TableWidthUnitValues.Dxa, Width = "0" });

            TableRow row = new TableRow();
            row.TableRowProperties = properties;

            htmlStyles.Runs.ProcessCommonAttributes(en, runStyleAttributes);
            if (runStyleAttributes.Count > 0)
                htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes.ToArray());

            tables.CurrentTable.Append(row);
            tables.CellPosition = new CellPosition(tables.CellPosition.Row + 1, 0);
        }

        #endregion

        #region ProcessTableColumn

        private void ProcessTableColumn(HtmlEnumerator en)
        {
            if (!tables.HasContext) return;

            TableCellProperties properties = new TableCellProperties();
            // in Html, table cell are vertically centered by default
            properties.TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
            List<OpenXmlElement> runStyleAttributes = new List<OpenXmlElement>();

            Unit unit = en.StyleAttributes.GetAsUnit("width");
            if (!unit.IsValid) unit = en.Attributes.GetAsUnit("width");

            // The heightUnit used to retrieve a height value.
            Unit heightUnit = en.StyleAttributes.GetAsUnit("height");
            if (!heightUnit.IsValid) heightUnit = en.Attributes.GetAsUnit("height");

            switch (unit.Type)
            {
                case UnitMetric.Percent:
                    properties.TableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Pct, Width = (unit.Value * 50).ToString(CultureInfo.InvariantCulture) };
                    break;
                case UnitMetric.Point:
                    // unit.ValueInPoint used instead of ValueInDxa
                    properties.TableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Auto, Width = (unit.ValueInPoint * 20).ToString(CultureInfo.InvariantCulture) };
                    break;
                case UnitMetric.Pixel:
                    properties.TableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = (unit.ValueInDxa).ToString(CultureInfo.InvariantCulture) };
                    break;
            }

            // fix an issue when specifying the RowSpan or ColSpan=1 (reported by imagremlin)
            int? colspan = en.Attributes.GetAsInt("colspan");
            if (colspan.HasValue && colspan.Value > 1)
            {
                properties.GridSpan = new GridSpan() { Val = colspan };
            }

            int? rowspan = en.Attributes.GetAsInt("rowspan");
            if (rowspan.HasValue && rowspan.Value > 1)
            {
                properties.VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Restart };

                var p = tables.CellPosition;
                int shift = 0;
                // if there is already a running rowSpan on a left-sided column, we have to shift this position
                foreach (var rs in tables.RowSpan)
                    if (rs.CellOrigin.Row < p.Row && rs.CellOrigin.Column <= p.Column + shift) shift++;

                p.Offset(0, shift);
                tables.RowSpan.Add(new HtmlTableSpan(p)
                {
                    RowSpan = rowspan.Value - 1,
                    ColSpan = colspan.HasValue && rowspan.Value > 1 ? colspan.Value : 0
                });
            }

            // Manage vertical text (only for table cell)
            string direction = en.StyleAttributes["writing-mode"];
            if (direction != null)
            {
                switch (direction)
                {
                    case "tb-lr":
                        properties.TextDirection = new TextDirection() { Val = TextDirectionValues.BottomToTopLeftToRight };
                        properties.TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
                        htmlStyles.Tables.BeginTagForParagraph(en.CurrentTag, new Justification() { Val = JustificationValues.Center });
                        break;
                    case "tb-rl":
                        properties.TextDirection = new TextDirection() { Val = TextDirectionValues.TopToBottomRightToLeft };
                        properties.TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
                        htmlStyles.Tables.BeginTagForParagraph(en.CurrentTag, new Justification() { Val = JustificationValues.Center });
                        break;
                }
            }

            var padding = en.StyleAttributes.GetAsMargin("padding");
            if (!padding.IsEmpty)
            {
                TableCellMargin cellMargin = new TableCellMargin();
                var cellMarginSide = new List<KeyValuePair<Unit, TableWidthType>>();
                cellMarginSide.Add(new KeyValuePair<Unit, TableWidthType>(padding.Top, new TopMargin()));
                cellMarginSide.Add(new KeyValuePair<Unit, TableWidthType>(padding.Left, new LeftMargin()));
                cellMarginSide.Add(new KeyValuePair<Unit, TableWidthType>(padding.Bottom, new BottomMargin()));
                cellMarginSide.Add(new KeyValuePair<Unit, TableWidthType>(padding.Right, new RightMargin()));

                foreach (var pair in cellMarginSide)
                {
                    if (!pair.Key.IsValid || pair.Key.Value == 0) continue;
                    if (pair.Key.Type == UnitMetric.Percent)
                    {
                        pair.Value.Width = (pair.Key.Value * 50).ToString(CultureInfo.InvariantCulture);
                        pair.Value.Type = TableWidthUnitValues.Pct;
                    }
                    else
                    {
                        pair.Value.Width = pair.Key.ValueInDxa.ToString(CultureInfo.InvariantCulture);
                        pair.Value.Type = TableWidthUnitValues.Dxa;
                    }

                    cellMargin.AddChild(pair.Value);
                }

                properties.TableCellMargin = cellMargin;
            }

            var border = en.StyleAttributes.GetAsBorder("border");
            if (!border.IsEmpty)
            {
                properties.TableCellBorders = new TableCellBorders();

                if (border.Left.IsValid)
                    properties.TableCellBorders.LeftBorder = new LeftBorder { Val = border.Left.Style, Color = StringValue.FromString(border.Left.Color.ToHexString()), Size = (uint)border.Left.Width.ValueInDxa };
                if (border.Right.IsValid)
                    properties.TableCellBorders.RightBorder = new RightBorder { Val = border.Right.Style, Color = StringValue.FromString(border.Right.Color.ToHexString()), Size = (uint)border.Right.Width.ValueInDxa };
                if (border.Top.IsValid)
                    properties.TableCellBorders.TopBorder = new TopBorder { Val = border.Top.Style, Color = StringValue.FromString(border.Top.Color.ToHexString()), Size = (uint)border.Top.Width.ValueInDxa };
                if (border.Bottom.IsValid)
                    properties.TableCellBorders.BottomBorder = new BottomBorder { Val = border.Bottom.Style, Color = StringValue.FromString(border.Bottom.Color.ToHexString()), Size = (uint)border.Bottom.Width.ValueInDxa };
            }

            htmlStyles.Tables.ProcessCommonAttributes(en, runStyleAttributes);
            if (styleAttributes.Count > 0)
                htmlStyles.Tables.BeginTag(en.CurrentTag, styleAttributes);
            if (runStyleAttributes.Count > 0)
                htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes.ToArray());

            TableCell cell = new TableCell();
            if (properties.HasChildren) cell.TableCellProperties = properties;

            // The heightUnit value used to append a height to the TableRowHeight.
            var row = tables.CurrentTable.GetLastChild<TableRow>();

            switch (heightUnit.Type)
            {
                case UnitMetric.Point:
                    row.TableRowProperties.AddChild(new TableRowHeight() { HeightType = HeightRuleValues.AtLeast, Val = (uint)(heightUnit.Value * 20) });

                    break;
                case UnitMetric.Pixel:
                    row.TableRowProperties.AddChild(new TableRowHeight() { HeightType = HeightRuleValues.AtLeast, Val = (uint)heightUnit.ValueInDxa });
                    break;
            }

            row.Append(cell);

            if (en.IsSelfClosedTag) // Force a call to ProcessClosingTableColumn
                ProcessClosingTableColumn(en);
            else
            {
                // we create a new currentParagraph to add new runs inside the TableCell
                cell.Append(currentParagraph = new Paragraph());
            }
        }

        #endregion

        #region ProcessTablePart

        private void ProcessTablePart(HtmlEnumerator en)
        {
            List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();

            htmlStyles.Tables.ProcessCommonAttributes(en, styleAttributes);

            if (styleAttributes.Count > 0)
                htmlStyles.Tables.BeginTag(en.CurrentTag, styleAttributes.ToArray());
        }

        #endregion

        #region ProcessXmlDataIsland

        private void ProcessXmlDataIsland(HtmlEnumerator en)
        {
            // Process inner Xml data island and do nothing.
            // The Xml has this format:
            /* <?xml:namespace prefix=o ns="urn:schemas-microsoft-com:office:office">
			   <globalGuideLine>
				   <employee>
					  <FirstName>Austin</FirstName>
					  <LastName>Hennery</LastName>
				   </employee>
			   </globalGuideLine>
			 */

            // Move to the first root element of the Xml then process until the end of the xml chunks.
            while (en.MoveNext() && !en.IsCurrentHtmlTag) ;

            if (en.Current != null)
            {
                string xmlRootElement = en.ClosingCurrentTag;
                while (en.MoveUntilMatch(xmlRootElement)) ;
            }
        }

        #endregion

        // Closing tags

        #region ProcessClosingBlockQuote

        private void ProcessClosingBlockQuote(HtmlEnumerator en)
        {
            CompleteCurrentParagraph(true);
            htmlStyles.Paragraph.EndTag("<blockquote>");
        }

        #endregion

        #region ProcessClosingDiv

        private void ProcessClosingDiv(HtmlEnumerator en)
        {
            // Mimic the rendering of the browser:
            ProcessBr(en);
            ProcessClosingTag(en);
        }

        #endregion

        #region ProcessClosingTag

        private void ProcessClosingTag(HtmlEnumerator en)
        {
            string openingTag = en.CurrentTag.Replace("/", "");
            htmlStyles.Runs.EndTag(openingTag);
            htmlStyles.Paragraph.EndTag(openingTag);
        }

        #endregion

        #region ProcessClosingNumberingList

        private void ProcessClosingNumberingList(HtmlEnumerator en)
        {
            htmlStyles.NumberingList.EndList();

            // If we are no more inside a list, we move to another paragraph (as we created
            // one for containing all the <li>. This will ensure the next run will not be added to the <li>.
            if (htmlStyles.NumberingList.LevelIndex == 0)
                AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
        }

        #endregion

        #region ProcessClosingCode

        private void ProcessClosingCode(HtmlEnumerator en)
        {



        }

        #endregion

        #region ProcessClosingParagraph

        private void ProcessClosingParagraph(HtmlEnumerator en)
        {
            CompleteCurrentParagraph(true);

            string tag = en.CurrentTag.Replace("/", "");
            htmlStyles.Runs.EndTag(tag);
            htmlStyles.Paragraph.EndTag(tag);
        }

        #endregion

        #region ProcessClosingQuote

        private void ProcessClosingQuote(HtmlEnumerator en)
        {
            Run run = new Run(
                new Text(HtmlStyles.QuoteCharacters.Suffix) { Space = SpaceProcessingModeValues.Preserve }
            );
            htmlStyles.Runs.ApplyTags(run);
            elements.Add(run);

            htmlStyles.Runs.EndTag("<q>");
        }

        #endregion

        #region ProcessClosingTable

        private void ProcessClosingTable(HtmlEnumerator en)
        {
            htmlStyles.Tables.EndTag("<table>");
            htmlStyles.Runs.EndTag("<table>");

            TableRow row = tables.CurrentTable.GetFirstChild<TableRow>();
            // Is this a misformed or empty table?
            if (row != null)
            {
                // Count the number of tableCell and add as much GridColumn as we need.
                TableGrid grid = new TableGrid();
                foreach (TableCell cell in row.Elements<TableCell>())
                {
                    // If that column contains some span, we need to count them also
                    int count = cell.TableCellProperties?.GridSpan?.Val ?? 1;
                    for (int i = 0; i < count; i++)
                    {
                        grid.Append(new GridColumn());
                    }
                }

                tables.CurrentTable.InsertAt<TableGrid>(grid, 1);
            }

            tables.CloseContext();

            if (!tables.HasContext)
                AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
        }

        #endregion

        #region ProcessClosingTablePart

        private void ProcessClosingTablePart(HtmlEnumerator en)
        {
            string closingTag = en.CurrentTag.Replace("/", "");

            htmlStyles.Tables.EndTag(closingTag);
        }

        #endregion

        #region ProcessClosingTableRow

        private void ProcessClosingTableRow(HtmlEnumerator en)
        {
            if (!tables.HasContext) return;
            TableRow row = tables.CurrentTable.GetLastChild<TableRow>();
            if (row == null) return;

            // Word will not open documents with empty rows (reported by scwebgroup)
            if (row.GetFirstChild<TableCell>() == null)
            {
                row.Remove();
                return;
            }

            // Add empty columns to fill rowspan
            if (tables.RowSpan.Count > 0)
            {
                int rowIndex = tables.CellPosition.Row;

                for (int i = 0; i < tables.RowSpan.Count; i++)
                {
                    HtmlTableSpan tspan = tables.RowSpan[i];
                    if (tspan.CellOrigin.Row == rowIndex) continue;

                    TableCell emptyCell = new TableCell(new TableCellProperties
                    {
                        TableCellWidth = new TableCellWidth() { Width = "0" },
                        VerticalMerge = new VerticalMerge()
                    },
                                        new Paragraph());

                    tspan.RowSpan--;
                    if (tspan.RowSpan == 0) { tables.RowSpan.RemoveAt(i); i--; }

                    // in case of both colSpan + rowSpan on the same cell, we have to reverberate the rowSpan on the next columns too
                    if (tspan.ColSpan > 0) emptyCell.TableCellProperties.GridSpan = new GridSpan() { Val = tspan.ColSpan };

                    TableCell cell = row.GetFirstChild<TableCell>();
                    if (tspan.CellOrigin.Column == 0 || cell == null)
                    {
                        row.InsertAfter(emptyCell, row.TableRowProperties);
                        continue;
                    }

                    // find the good column position, taking care of eventual colSpan
                    int columnIndex = 0;
                    while (columnIndex < tspan.CellOrigin.Column)
                    {
                        columnIndex += cell.TableCellProperties?.GridSpan?.Val ?? 1;
                    }
                    //while ((cell = cell.NextSibling<TableCell>()) != null);

                    if (cell == null) row.AppendChild(emptyCell);
                    else row.InsertAfter<TableCell>(emptyCell, cell);
                }
            }

            htmlStyles.Tables.EndTag("<tr>");
            htmlStyles.Runs.EndTag("<tr>");
        }

        #endregion

        #region ProcessClosingTableColumn

        private void ProcessClosingTableColumn(HtmlEnumerator en)
        {
            if (!tables.HasContext)
            {
                // When the Html is bad-formed and doesn't contain <table>, the browser renders the column separated by a space.
                // So we do the same here
                Run run = new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve });
                htmlStyles.Runs.ApplyTags(run);
                elements.Add(run);
                return;
            }
            TableCell cell = tables.CurrentTable.GetLastChild<TableRow>().GetLastChild<TableCell>();

            // As we add automatically a paragraph to the cell once we create it, we'll remove it if finally, it was not used.
            // For all the other children, we will ensure there is no more empty paragraphs (similarly to what we do at the end
            // of the convert processing).
            // use a basic loop instead of foreach to allow removal (bug reported by antgraf)
            for (int i = 0; i < cell.ChildElements.Count;)
            {
                Paragraph p = cell.ChildElements[i] as Paragraph;
                // care of hyperlinks as they are not inside Run (bug reported by mdeclercq github.com/onizet/html2openxml/workitem/11162)
                if (p != null && !p.HasChild<Run>() && !p.HasChild<Hyperlink>()) p.Remove();
                else i++;
            }

            // We add this paragraph regardless it has elements or not. A TableCell requires at least a Paragraph, as the last child of
            // of a table cell.
            // additional check for a proper cleaning (reported by antgraf github.com/onizet/html2openxml/discussions/272744)
            if (!(cell.LastChild is Paragraph) || elements.Count > 0) cell.Append(new Paragraph(elements));

            htmlStyles.Tables.ApplyTags(cell);

            // Reset all our variables and move to next cell
            this.elements.Clear();
            String openingTag = en.CurrentTag.Replace("/", "");
            htmlStyles.Tables.EndTag(openingTag);
            htmlStyles.Runs.EndTag(openingTag);

            var pos = tables.CellPosition;
            pos.Column++;
            tables.CellPosition = pos;
        }

        #endregion
    }
}
