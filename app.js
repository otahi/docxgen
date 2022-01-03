function generate() {
    const txt = document.querySelector('#txt').value;

    const lines = txt.split('\n');
    const items = new Array();
    const re = /^(?<level>#+).*$/
    for (var i = 0; i < lines.length; i++) {
        const match = re.exec(lines[i])
        if (match) {
            const level = match.groups.level.length;
            items.push({
                level: level,
                text: lines[i].substr(level).trim()
            })
        } else {
            items.push({
                level: 0,
                text: lines[i].trim()
            })
        }
    }

    const body = new Array();
    for (var i = 0; i < items.length; i++) {
        const item = items[i];
        const level = item.level;
        const text = item.text;

        if (level == 1) {
            body.push(new docx.Paragraph({ text: text, heading: docx.HeadingLevel.TITLE }));
        } else if (level > 1 && level <= 6) {
            body.push(new docx.Paragraph({ text: text, heading: eval("docx.HeadingLevel.HEADING_" + String(level - 1)) }));
        } else {
            body.push(new docx.Paragraph({ text: text }));
        }
    }

    const doc = new docx.Document({
        sections: [{
            properties: {},
            children: body
        }]
    });

    docx.Packer.toBlob(doc).then((blob) => {
        saveAs(blob, "generated.docx");
        console.log("Document created successfully");
    });
}