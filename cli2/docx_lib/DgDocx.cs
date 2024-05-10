namespace docx_lib;
using Markdig;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using System.Text;
using System.Collections.Generic;
using System.Text.Json;
using Mammoth;

public class DgDocx
{
    private static IEnumerable<HyperlinkRelationship> hyperlinks;
    private static Dictionary<string, Stream> images;
    private static WordprocessingDocument wordDoc;
    private static MainDocumentPart mainPart;

    private static int linksCount = 0;

    // stream here because anticipating zip.
    public async static Task md_to_docx(string md, Stream outputStream) //String mdFile, String docxFile, String template)
    {
        MarkdownPipeline pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();

        var html = Markdown.ToHtml(md, pipeline);

        //All the document is being saved in the stream
        using (wordDoc = WordprocessingDocument.Create(outputStream, WordprocessingDocumentType.Document, true))
        {
            mainPart = wordDoc.AddMainDocumentPart();

            // Create the document structure and add some text.
            mainPart.Document = new Document();


            HtmlConverter converter = new HtmlConverter(mainPart);
            converter.ParseHtml(html);
            mainPart.Document.Save();
        }
    }

    public async static Task md_to_docx(JsonElement[] mdFiles, string images, Dictionary<string, MemoryStream> outputData) //String mdFile, String docxFile, String template)
    {
        MarkdownPipeline pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
        for (int i = 0; i < mdFiles.Length; i++)
        {
            MemoryStream outputStream = new MemoryStream();
            var html = Markdown.ToHtml(mdFiles[i].GetProperty("file").GetString(), pipeline);
            //edit on debug the h
            //All the document is being saved in the stream
            using (wordDoc = WordprocessingDocument.Create(outputStream = new MemoryStream(), WordprocessingDocumentType.Document, true))
            {
                mainPart = wordDoc.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();

                HtmlConverter converter = new HtmlConverter(mainPart);
                converter.ParseHtml(html, images);
                mainPart.Document.Save();
            }
            outputData.Add(mdFiles[i].GetProperty("src").GetString(), outputStream);
        }
    }

    public async static Task docx_to_md(Stream infile, Stream outfile, Dictionary<string, Stream> images, string name = "")
    {
        wordDoc = WordprocessingDocument.Open(infile, false);
        mainPart = wordDoc.MainDocumentPart;
        DgDocx.images = images;

        StringBuilder textBuilder = new StringBuilder();
        StyleDefinitionsPart styleDefinitionsPart = wordDoc.MainDocumentPart.StyleDefinitionsPart;
        hyperlinks = wordDoc.MainDocumentPart.HyperlinkRelationships;
        Body body = mainPart.Document.Body;

        if (body != null)
        {
            //var asd = parts.Descendants<HyperlinkList>();
            foreach (var block in body.ChildElements)
            {
                //This method is for manipulating the style of Paragraphs and text inside
                if (block is Paragraph) MdController.ProcessParagraph((Paragraph)block, textBuilder, mainPart, images, hyperlinks);

                if (block is Table) MdController.ProcessTable((Table)block, textBuilder);

            }
        }

        //This code is replacing the below one because I need to check the .md file faster
        //writing the .md file in test_result folder
        if (name != "")
        {
            var writer = new StreamWriter(outfile);
            string s = textBuilder.ToString();
            writer.Write(s);
            writer.Flush();
        }
        else
        {
            var writer = new StreamWriter(outfile);
            string s = textBuilder.ToString();
            writer.Write(s);
            writer.Flush();
        }

    }

    public async static Task docx_to_html(Stream infile, StringBuilder outfile, string name = "")
    {
        var converter = new DocumentConverter();
    // .ImageConverter(image => {
    //     using (var stream = image.GetStream()) {
    //         var base64 = Convert.ToBase64String(stream);
    //         var src = "data:" + image.ContentType + ";base64," + base64;
    //         return new Dictionary<string, string> { { "src", src } };
    //     }
    // });
        var result = converter.ConvertToHtml(infile);
        outfile.Append(result.Value); // The generated HTML
        var warnings = result.Warnings; // Any warnings during conversion
    }
}



