import * as JSZip from './jszip.min.js';
var fileSaved;

window.prueba = (input) => {

    fileSaved = input.files[0];
}

window.ejecutar = () => {
    if (fileSaved.name.endsWith('.docx')) {
        toMd(fileSaved);
    } else if (fileSaved.name.endsWith('.md')) {
        toDocx(fileSaved);
    } else {
        toDocx(fileSaved);
    }
}

//-------------------------------------------------
// Function below is to convert docx to md
//-------------------------------------------------
function toMd(file) {
    var reader = new FileReader();
    reader.onload = async function (e) {

        var zipBytes = await DotNet.invokeMethodAsync("docx2md-wasm", "openDocxZipFile", new Uint8Array(reader.result), file.name);

        downloadBlob(zipBytes, 'testMd.zip', 'application/octet-stream');
    };
    reader.readAsArrayBuffer(file);
}

function toHtml(file) {

    var reader = new FileReader();
    reader.onload = async function (e) {

        var zipBytes = await DotNet.invokeMethodAsync("docx2md-wasm", "openDocxZipFile2", new Uint8Array(reader.result), file.name);

        downloadBlob(zipBytes, 'testHtml.html', 'application/octet-stream');
    };
    reader.readAsArrayBuffer(file);
}

//-------------------------------------------------
// Function below is to convert md to docx
//-------------------------------------------------
async function toDocx(file) {

    if (file.name.endsWith('.zip')) {
        const jszip = new window.JSZip();
        const decoder = new TextDecoder();
        var mdString = [];
        var test = [];
        const images = [];
        const zipFiles = [];


        jszip.loadAsync(file).then(async function (zip) {

            zipFiles.push(...zip.folder("articles").file(/^[^\.]/));
            zipFiles.push(...zip.folder("images").file(/^[^\.]/));

            const DGCoverterPromises = [];
            const fileNames = [];

            zipFiles.forEach(async function (file, index) {
                if (file.name.startsWith('images/')) {
                    DGCoverterPromises.push(file.async('uint8array'))
                    fileNames.push(file.name)
                }
                else if (file.name.startsWith('articles/')) {
                    DGCoverterPromises.push(file.async("text"))
                    fileNames.push(file.name)
                }
            })

            await Promise.all(DGCoverterPromises).then(function (data) {
                data.forEach(function (file, index) {
                    if (typeof file === 'string') {
                        mdString.push({ src: fileNames[index], file: file });
                    } else {
                        let imageHex = Array.from(file)
                            .map(b => b.toString(16).padStart(2, '0'))
                            .join('');
                        if (imageHex.length % 2 !== 0) {
                            // Add a padding of 0 to the end of hexString
                            imageHex += '0';
                        }
                        images.push({ src: fileNames[index], hex: imageHex });
                    }
                })
            });

            const jsonString = JSON.stringify(images);

            var zipBytes = await DotNet.invokeMethodAsync("docx2md-wasm", "openMdZipFile", mdString, jsonString);

            downloadBlob(zipBytes, 'test.zip', 'application/octet-stream');
        })
    } else {
        var reader = new FileReader();
        reader.onload = async function (e) {

            var byte = await DotNet.invokeMethodAsync("docx2md-wasm", "openMdFile", new Uint8Array(reader.result));;

            downloadBlob(byte, "test.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        }
        reader.readAsArrayBuffer(file)
    }
};


//-------------------------------------------------
// Function below is to convert an image to hex format
//-------------------------------------------------

function convertToHex(image) {
    let hex = '';
    let finalHex = '';
    for (let i = 0; i < image.length; i++) {
        hex += ("00" + image.charCodeAt(i).toString(16)).slice(-4);
    }
    return hex;
}


//-------------------------------------------------
// Function below is to convert a byte to hex format
//-------------------------------------------------

function byteToHex(byte) {
    return byte.toString(16).padStart(2, '0');
}

function bytesToHex(bytes) {
    return bytes.map(byteToHex);
}

function bytesToHexString(bytes) {
    return bytesToHex(bytes).join('');
}

//-------------------------------------------------
// Function below are to download the files after the conversion
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