import * as JSZip from './jszip.min.js';
import Image from "./images.js";
// import {unified} from './node_modules/unified/index.js';
// import remarkParse from './node_modules/remark-parse/index.js';
// import remarkRehype from './node_modules/remark-rehype/index.js';
// import rehypeStringify from './node_modules/rehype-stringify/index.js';
// import rehypeSanitize from './node_modules/rehype-sanitize/index.js';
// import rehypeSlug from './node_modules/rehype-slug/index.js';


window.convertToMd = (input) => {

  const file = input.files[0];

  var reader = new FileReader();
  reader.onload = async function (e) {
    // The file's text will be printed here
    var string = await DotNet.invokeMethodAsync("blazorwasm", "openDocxFile", new Uint8Array(reader.result));

    downloadBlob(string, 'test.md', 'application/octet-stream');
  };
  reader.readAsArrayBuffer(file);
}

window.convertToDocx = (input) => {

  const file = input.files[0];
  const jszip = new window.JSZip();
  var jsonString = "";
  var mdString = [];
  const images = [];


  jszip.loadAsync(file).then(async function (zip) {
    zip.forEach(function (relativePath, zipEntry) {
      if (zipEntry.name === 'images/') {
        zip.folder("images/").forEach(function (relativePath, zipFile) {
          if (!file.dir) {
            console.log(zipFile);

            // The file's text will be printed here
            let imageHex = convertToHex(zipFile._data.compressedContent);
            images.push(new Image(zipFile.name, imageHex));
          }
        })

      }
      if (zipEntry.name === 'articles/') {
        zip.folder("articles/").forEach(function (relativePath, file) {
          if (!file.dir) {
            console.log("arrived");
            console.log(file.name);
            mdString.push(String.fromCharCode.apply(null, file._data.compressedContent));
            //md2html(file);
          }
        });
      }
    });
    jsonString = createJsonImages(images);

    //Here I need to pass the md array and json images. It'll return bytes.

    var bytes = await DotNet.invokeMethodAsync("blazorwasm", "openDocxFile", new Uint8Array(reader.result));
  })
  

};


function createJsonImages(images) {

  var json = {}

  images.forEach(function (img) {
    json[img.src] = img.hex;
  });

  var jsonString = JSON.stringify(json);

  return jsonString;

}

//-------------------------------------------------
// Function below is to convert .md file to html
//-------------------------------------------------

// async function md2html(md){
//   const file = await unified()
//       .use(remarkParse)
//       .use(remarkRehype)
//       .use(rehypeSanitize)
//       .use(rehypeStringify)
//       .use(rehypeSlug)
//       .process(md ?? "error")
//   return String(file)
// }


//-------------------------------------------------
// Function below is to convert and image to hex format
//-------------------------------------------------

function convertToHex(image) {
  let hex = '';
  const hexArray = image;
  hexArray.forEach(function (byte) {
    hex += byte.toString(16).padStart(2, '0');
  });
  console.log(hex);
  return hex;
}


//-------------------------------------------------
// Functions below are to download the files after the conversion
//-------------------------------------------------

function downloadBlob(data, fileName, mimeType) {
  var blob = new Blob([data], {
    type: mimeType
  });
  var url = window.URL.createObjectURL(blob);
  downloadURL(url, fileName);
  setTimeout(function () {
    return window.URL.revokeObjectURL(url);
  }, 1000);
};

var downloadURL = function (data, fileName) {
  var a;
  a = document.createElement('a');
  a.href = data;
  a.download = fileName;
  document.body.appendChild(a);
  a.style = 'display: none';
  a.click();
  a.remove();
};


//-------------------------------------------------
// Functions below failed
//-------------------------------------------------

// async function readImages(htmlFile) {
//   const parser = new DOMParser();
//   let response

//   const htmlTags = parser.parseFromString(htmlFile, "text/html");

//   const imgTags = htmlTags.querySelectorAll("img");
//   for (const jpg of imgTags) {
//     console.log(jpg.src);
//     const result = await fetch(jpg.src);
//   }

//   imgTags.push(htmlTags.querySelector("img").forEach((e) => e.src));
// }



//Previous convertToDocx (Doesn't work because of CORS)
// function convertToDocx(input) {
//   const file = input.files[0];

//   var reader = new FileReader();
//   reader.onload = async function (e) {

//     var byte = await DotNet.invokeMethodAsync("blazorwasm", "openMdFile", new Uint8Array(reader.result));;

//     downloadBlob(byte, "test.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
//   }
//   reader.readAsArrayBuffer(file)
// }