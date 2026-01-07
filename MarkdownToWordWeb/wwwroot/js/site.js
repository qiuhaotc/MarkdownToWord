// Download file function for Blazor
window.downloadFile = function (fileName, base64Content) {
    const linkSource = `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${base64Content}`;
    const downloadLink = document.createElement("a");
    downloadLink.href = linkSource;
    downloadLink.download = fileName;
    downloadLink.click();
};

