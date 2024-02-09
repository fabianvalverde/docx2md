

```
mkdir cli2
cd cli2
code .
dotnet new sln

dotnet new classlib -o docx_lib
cd docx_lib
dotnet add package Markdig --version 0.30.4
dotnet add package NS.HtmlToOpenXml --version 1.1.0
cd ..

dotnet new console -o docx_md
dotnet new blazorwasm -o docx_web

dotnet sln add docx_md/docx_md.csproj
dotnet sln add docx_web/docx_web.csproj 
dotnet sln add docx_lib/docx_lib.csproj 

dotnet add docx_md/docx_md.csproj reference docx_lib/docx_lib.csproj

dotnet add docx_web/docx_web.csproj reference docx_lib/docx_lib.csproj

dotnet build
dotnet publish -c release
cd bin/Release/net6.0/publish/wwwroot
surge . mddocx2.surge.sh

```

# About

"Package name" allows you to translate Markdown file to .docx files and vice versa. It's able to read and translate to .docx following Markdown features:

- Headings
- Paragraphs
- Bold text (Only available for * symbol)
- Italic text (Only available for * symbol)
- Bold and Italic text (Only available for * symbol)
- Blockquotes
- Unordered lists
- Code blocks

All that is available following markdown best practices 

# Getting Started

This section only wants to cover the basics to translate from markdown files to .docx format.

To use this package just clone it and reference cli2\docx_lib\docx_lib.csproj into your main project where you want to use it.

## Read Markdown Files (.md to .docx)

This is the method you need to get the .md file 
DgDocx.md_to_docx(JsonElement[] mdFiles, string images, Dictionary<string, MemoryStream> outputData)

mdFiles = a JsonElement that must have the space "src" (file name) and the space "file" that will store all the data you can read.
outfile = empty new Dictionary<string, Stream>(). To store the name of the file and the file as a Stream.
images = Must be a Json that will store the "src" (file name) and the image as a hexadecimal in a string.
name = file's name you want to convert.

Having all that data you can follow this piece of code (.net8) to zip everything.
```
    public static async Task<byte[]> convertToZipDocx(object[] mdObject, string images)
    {
        string json = JsonSerializer.Serialize(mdObject);
        JsonElement[] mdFiles = JsonSerializer.Deserialize<JsonElement[]>(json);

        try
        {
            Dictionary<string, MemoryStream> dictionary = new Dictionary<string, MemoryStream>();
            

            await DgDocx.md_to_docx(mdFiles, images, dictionary);

            MemoryStream zipStream = new MemoryStream();
            //The .zip file stored in the zipStream
            using (ZipArchive zipArchive = new ZipArchive(zipStream, ZipArchiveMode.Create, true))
            {
                for (int i = 0; i < dictionary.Count; i++)
                {
                    //The entry already created in the .zip
                    ZipArchiveEntry entry = zipArchive.CreateEntry($"DGConvertor/Articles/{dictionary.ElementAt(i).Key.Replace("articles/", "").Replace(".md", "")}.docx");
                    using (Stream entryStream = entry.Open())
                    {
                        try
                        {
                            dictionary.ElementAt(i).Value.Position = 0;
                            //byte[] bytes = Encoding.ASCII.GetBytes(mdFiles[i]);
                            await dictionary.ElementAt(i).Value.CopyToAsync(entryStream);
                            //entryStream.Write(bytes);
                            entryStream.Close();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                        }
                    }
                }
            }
            return zipStream.ToArray();
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            return new byte[] { };
        }

    }
```
## Read .docx Files (.docx to md)

This is the method you need to get the .md file 
DgDocx.docx_to_md(Stream infile, Stream outfile, Dictionary<string, Stream> images, string name = "")

infile = .docx file you want to convert.
outfile = empty new MemoryStream(). Here you'll store the .docx data.
images = empty new Dictionary<string, Stream>(). Here you'll get the .docx images.
name = file's name you want to convert.

Having all that data you can follow this piece of code (.net8).

```
public static async Task<byte[]> docxToZipMd(byte[] file, string name)
    {
        MemoryStream outStream = new MemoryStream();
        
        Dictionary<string, Stream> images = new Dictionary<string, Stream>();

        using (MemoryStream data = new MemoryStream(file))
        {
            await DgDocx.docx_to_md(data, outStream, images);
            @* StreamReader reader = new StreamReader(outStream);
            md = reader.ReadToEnd(); *@
        }

        MemoryStream zipStream = new MemoryStream();
        using (ZipArchive zipArchive = new ZipArchive(zipStream, ZipArchiveMode.Create, true))
        {
            //The entry already created in the .zip
            ZipArchiveEntry documentEntry = zipArchive.CreateEntry($"DGConvertor/Articles/{name.Replace(".docx", "")}.md");
            using (Stream entryStream = documentEntry.Open())
            {
                try
                {
                    outStream.Position = 0;
                    //byte[] bytes = Encoding.ASCII.GetBytes(mdFiles[i]);
                    await outStream.CopyToAsync(entryStream);
                    //entryStream.Write(bytes);
                    entryStream.Close();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
            foreach (KeyValuePair<string, Stream> img in images)
            {
                ZipArchiveEntry imagesEntry = zipArchive.CreateEntry($"DGConvertor/Images/{img.Key}");
                using (Stream entryStream = imagesEntry.Open())
                {
                    try
                    {
                        img.Value.Position = 0;
                        await img.Value.CopyToAsync(entryStream);
                         entryStream.Close();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }
            }
        }
        return zipStream.ToArray();
    }
```

## Read .docx Files (.docx to .html)

This is the method you need to get the .md file 
DgDocx.docx_to_html(Stream infile, StringBuilder outfile, string name = "")

infile = .docx file you want to convert.
outfile = empty new StringBuilder(). Here you'll store the .docx data.
name = file's name you want to convert.

Having all that data you can follow this piece of code (.net8).
```
public static async Task<string> docxToZipHtml(byte[] file, string name)
    {
        StringBuilder html = new StringBuilder();
        string htmlFinal;

        using (MemoryStream data = new MemoryStream(file))
        {
            await DgDocx.docx_to_html(data, html, name);
            @* StreamReader reader = new StreamReader(outStream);
            md = reader.ReadToEnd(); *@
            htmlFinal = html.ToString();
        }
        return htmlFinal;
    }
```
# Bugs

# Future Development

- Code blocks (Missing .docx to md)
- Images 
- Links (Every type of links)
- Escape characters (https://www.markdownguide.org/basic-syntax/#escaping-characters)
- Check box lists (Missing .docx to md)
- Code (https://www.markdownguide.org/basic-syntax/#code)
- Ordered lists
