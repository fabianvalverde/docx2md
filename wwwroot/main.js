// const df = document.getElementById('docxFile')
// const mf = document.getElementById('mdFile')
// console.log("attaching...", df)


// function downloadMdFile(){

// }

// df.addEventListener('change', (event) => {
//     const file = event.target.files[0]
//     console.log(file)
//     var reader = new FileReader();
//     reader.onload = async (event) => {
//         //console.log("xx", event)
//         //console.log(reader.result);
//         //console.log(exports.TodoMVC.MainJS.openFile(new Uint8Array(reader.result)))
//         var string = exports.TodoMVC.MainJS.openDocxFile(new Uint8Array(reader.result));

//         downloadBlob(string, 'test.md', 'application/octet-stream');
//     }
//     reader.readAsArrayBuffer(file);
// }, true);

// mf.addEventListener('change', (event) => {
//     const file = event.target.files[0]
//     console.log(file)
//     var reader = new FileReader();
//     reader.onload = async (event) => {
//         //console.log("xx", event)
//         //console.log(reader.result);
//         //console.log(exports.TodoMVC.MainJS.openFile(new Uint8Array(reader.result)))
//         var string = exports.TodoMVC.MainJS.openFile(new Uint8Array(reader.result));

//         downloadBlob(string, 'test.md', 'application/octet-stream');
//     }
//     reader.readAsArrayBuffer(file);
// }, true);


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

    var byte = await DotNet.invokeMethodAsync("blazorwasm", "openMdFile", new Uint8Array(reader.result));;

    downloadBlob(byte, "test.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  }
  reader.readAsArrayBuffer(file)
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