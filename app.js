function parse(txt) {
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
    return items;
}

function generateDocx() {
    const txt = document.querySelector('#txt').value;
    const items = parse(txt);

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

function generatePptx() {
    const TITLE_SLIDE = "Title";
    const TITLE_CONTENTS_SLIDE = "Title and Contents";

    const txt = document.querySelector('#txt').value;
    const items = parse(txt);

    const pptx = new PptxGenJS();

    pptx.defineSlideMaster({
        title: TITLE_SLIDE,
        type: 'ctrTitle',
        objects: [{
            placeholder: {
                options: { name: "title", type: "title", x: 1.25, y: 0.92, w: 7.5, h: 1.95 },
            },
        }, ],
    });

    pptx.defineSlideMaster({
        title: TITLE_CONTENTS_SLIDE,
        type: 'ctr',
        objects: [{
                placeholder: {
                    options: { name: "title", type: "title", x: 0.69, y: 0.3, w: 8.6, h: 1.09 },
                },
            },
            {
                placeholder: {
                    options: { name: "body", type: "body", x: 0.7, y: 1.5, w: 8.6, h: 3.57 },
                },
            },
        ],
    });

    let slide = null;
    let body = null;
    for (var i = 0; i < items.length; i++) {
        const item = items[i];
        const level = item.level;
        const text = item.text;

        if (level == 1) {
            slide = pptx.addSlide({ masterName: TITLE_SLIDE });
            slide.addText(text, { placeholder: "title" });
        } else if (level == 2) {
            if (slide != null && body != null && body.length > 0) {
                slide.addText(body, { placeholder: "body" });
                slide = null;
                body = null;
            }

            slide = pptx.addSlide({ masterName: TITLE_CONTENTS_SLIDE });
            slide.addText(text, { placeholder: "title" });
        } else if (text != null && text.length > 0) {
            console.log(text);
            if (body == null) {
                body = new Array();
            }
            if (level == 0) {
                body.push({ text: text, options: { fontSize: 20, softBreakBefore: true } });
            } else {
                body.push({ text: text, options: { fontSize: 24 - 2 * level, indentLevel: level - 3, bullet: true } });
            }
        }
    }

    // For final slide
    if (slide != null && body != null && body.length > 0) {
        slide.addText(body, { placeholder: "body" });
        slide = null;
        body = null;
    }

    pptx.writeFile({ fileName: "generated.pptx" });
}