window.downloadFile = (filename, contentType, base64) => {
    const a = document.createElement('a');
    a.href = `data:${contentType};base64,${base64}`;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
};

window.triggerClick = (elementId) => {
    const el = document.getElementById(elementId);
    if (el) el.click();
};
