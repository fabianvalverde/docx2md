function convertToMd(input) {

  const file = input.files[0];

  var reader = new FileReader();
  reader.onload = async function (e) {
    // The file's text will be printed here
    var string = await DotNet.invokeMethodAsync("blazorwasm", "openDocxFile", new Uint8Array(reader.result));

    downloadBlob(string, 'test.md', 'application/octet-stream');
  };
  reader.readAsArrayBuffer(file);
}

function convertToDocx(input) {
  const file = input.files[0];

  var reader = new FileReader();
  reader.onload = async function (e) {

    var md = window.markdownit();
    var htmlFile = md.render(reader.result);

    readImages(htmlFile);
  }
  reader.readAsBinaryString(file);
}

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

function readImages(htmlFile) {
  let text = "";
  let tags = [];
  let insideTag = false;
  let currentTag = "";

  for (let i = 0; i < htmlString.length; i++) {
    if (htmlString[i] === "<") {
      insideTag = true;
      currentTag = "<";
    } else if (htmlString[i] === ">") {
      insideTag = false;
      currentTag += ">";
      tags.push(currentTag);
    } else if (insideTag) {
      currentTag += htmlString[i];
    } else {
      text += htmlString[i];
    }
  }

  return { text: text, tags: tags };
}



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