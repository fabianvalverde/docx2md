using docx_lib;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO.Compression;
using System.Numerics;

internal class Program
{
    static async Task Main(string[] args)
    {
        Dictionary<string, Stream> images = new Dictionary<string, Stream>();
        //Directory to check results
        var outdir = @"./../../../../docx_md/test_results/";
        var outdirMedia = @"./../../../../docx_md/test_results/results/media/";

        //To check many md files at the same time
        string[] files = Directory.GetFiles(@"./../../../../docx_md/folder_tests/", "*.md", SearchOption.TopDirectoryOnly);


        foreach (var mdFile in files)
        {
            //Placing the route to store the results
            string fn = Path.GetFileNameWithoutExtension(mdFile);
            string outFile = outdir + fn.Replace("_test", "_result");
            try
            {
                // markdown to docx
                var md = File.ReadAllText(mdFile);
                var inputStream = new MemoryStream();
                await DgDocx.md_to_docx(md, inputStream);

                //inputStream is writing into the .docx file
                File.WriteAllBytes(outFile+".docx", inputStream.ToArray());



                //---------------------------------------TESTING PURPOSES---------------------------------------
                // convert the docx back to markdown.
                using (var instream = File.Open(outFile+".docx", FileMode.Open))
                {
                    var outstream = new MemoryStream();
                    await DgDocx.docx_to_md(instream, outstream, images, "asfa");//Previous: instream, outstream, fn.Replace("_md", "")

                    //The commented code is for .zip files

                    using (var fileStream = new FileStream(outFile+".md", FileMode.Create))
                    {
                        outstream.Seek(0, SeekOrigin.Begin);
                        outstream.CopyTo(fileStream);
                    }                        

                    //pull the images from "/media"
                    using (ZipArchive archive = new ZipArchive(instream, ZipArchiveMode.Update, true))
                    {
                        string subDirectory = "media/";
                        // Loop through each entry in the zip file
                        foreach (ZipArchiveEntry entry in archive.Entries)
                        {
                            // Check if the entry is a directory and its name matches the specified subdirectory
                            if (entry.FullName.Contains(subDirectory) && !entry.Name.EndsWith("/"))
                            {
                                Directory.CreateDirectory(outdirMedia);
                                // Extract the entry to the specified extract path
                                entry.ExtractToFile(outdirMedia + entry.Name.Replace(".bin",".jpeg"), true);
                            }
                        }

                    }

                }
                //---------------------------------------TESTING PURPOSES---------------------------------------



                using (ZipArchive archive = ZipFile.OpenRead(outFile+".docx"))
                {
                    archive.ExtractToDirectory(outdir + "test.unzipped", true);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"{mdFile} failed {e}");
            }
        }
    }

    static void AssertThatOpenXmlDocumentIsValid(WordprocessingDocument wpDoc)
    {

        var validator = new OpenXmlValidator(FileFormatVersions.Office2010);
        var errors = validator.Validate(wpDoc);

        if (!errors.GetEnumerator().MoveNext())
            return;

        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("The document doesn't look 100% compatible with Office 2010.\n");

        Console.ForegroundColor = ConsoleColor.Gray;
        foreach (ValidationErrorInfo error in errors)
        {
            Console.Write("{0}\n\t{1}", error.Path.XPath, error.Description);
            Console.WriteLine();
        }

        Console.ReadLine();
    }
}