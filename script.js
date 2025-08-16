list.innerHTML = "";
document.getElementById("certificateUpload").addEventListener("change", function(event) {
    const files = event.target.files;
    const list = document.getElementById("certificateList");

    for (let file of files) {
        const box = document.createElement("div");
        box.className = "certificate-box";

        if (file.type.startsWith("image/")) {
            const img = document.createElement("img");
            img.className = "certificate-preview";
            img.src = URL.createObjectURL(file);
            box.appendChild(img);
            addFileName(box, file);
            list.appendChild(box);
        } else if (file.type === "application/pdf") {
            previewPDF(file, box, list);
        }
    }

    // Reset input to allow uploading same file again
    event.target.value = "";

    function addFileName(box, file) {
        const link = document.createElement("a");
        link.href = URL.createObjectURL(file);
        link.target = "_blank";
        link.textContent = file.name;
        box.appendChild(link);
    }

    function previewPDF(file, box, list) {
        const canvas = document.createElement("canvas");
        canvas.className = "certificate-preview";

        const reader = new FileReader();
        reader.onload = function() {
            const loadingTask = pdfjsLib.getDocument({ data: reader.result });
            loadingTask.promise.then(function(pdf) {
                pdf.getPage(1).then(function(page) {
                    const scale = 0.5;
                    const viewport = page.getViewport({ scale: scale });
                    const context = canvas.getContext("2d");
                    canvas.height = 120;
                    canvas.width = 120;

                    page.render({
                        canvasContext: context,
                        viewport: viewport
                    }).promise.then(() => {
                        box.appendChild(canvas);
                        addFileName(box, file);
                        list.appendChild(box);
                    });
                });
            });
        };
        reader.readAsArrayBuffer(file);
    }
});

const { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType } = docx;

const form = document.getElementById("question-form");

form.addEventListener("submit", async (e) => {
  e.preventDefault();

  const name = form.name.value.trim();
  const email = form.email.value.trim();
  const question = form.question.value.trim();

  if (!name || !email || !question) {
    alert("Please fill all fields.");
    return;
  }

  // Optional: check word count limit
  const wordCount = question.split(/\s+/).filter(Boolean).length;
  if (wordCount > 500) {
    alert("Question exceeds 500 words limit.");
    return;
  }

  // Create a Word document
  const doc = new Document();

  // Create a table with two columns: Field and Value
  const table = new Table({
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            children: [new Paragraph({ text: "Field", bold: true })],
          }),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            children: [new Paragraph({ text: "Information", bold: true })],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph("Name")] }),
          new TableCell({ children: [new Paragraph(name)] }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph("Email ID")] }),
          new TableCell({ children: [new Paragraph(email)] }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph("Question")] }),
          new TableCell({ children: [new Paragraph(question)] }),
        ],
      }),
    ],
  });

  doc.addSection({
    children: [
      new Paragraph({ text: "User Question Submission", heading: docx.HeadingLevel.HEADING_1 }),
      table,
    ],
  });

  // Generate and download the Word document
  const blob = await Packer.toBlob(doc);
  window.saveAs(blob, "question_submission.docx");

  form.reset();
});
