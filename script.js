// main.js

// 1) Generar y guardar .docx en IndexedDB
document.getElementById("contractForm").addEventListener("submit", async (e) => {
  e.preventDefault();

  const data = {
    NOMBRE: document.getElementById("nombre").value,
    DIRECCION: document.getElementById("direccion").value,
    PLAN: document.getElementById("plan").value,
    VALOR_PLAN: document.getElementById("valorPlan").value,
    VALOR_PROMO: document.getElementById("valorPromo").value,
    CICLO: document.getElementById("ciclo").value,
    FECHA: document.getElementById("fecha").value,
  };

  try {
    const content = await loadFile("contrato_template.docx");
    const zip = new PizZip(content);
    const doc = new window.docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: { start: "<<", end: ">>" },
    });

    doc.render(data);
    const blob = doc.getZip().generate({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });

    await saveContrato(blob);
    console.log("DOCX generado y guardado");
    document.getElementById("preview").innerHTML =
      "<p>Contrato generado y guardado. Pulsa “Visualizar contrato”.</p>";
  } catch (err) {
    console.error("Error generando .docx:", err);
    alert("Error generando contrato. Revisa la consola.");
  }
});

// 2) Renderizar y exportar a PDF (multipágina)
document.getElementById("visualizarButton").addEventListener("click", async () => {
  try {
    // 2.1) Leer el blob .docx de IndexedDB
    const blob = await getContrato();
    if (!blob) {
      alert("No hay contrato generado.");
      return;
    }

    // 2.2) Renderizar el .docx en pantalla
    const archivo = new File([blob], "Contrato.docx", { type: blob.type });
    const container = document.getElementById("preview");
    container.innerHTML = "";
    await window.docx.renderAsync(archivo, container);
    console.log("Contrato renderizado en pantalla");

    // 2.3) Limpieza visual
    const imgs = container.querySelectorAll("img");
    if (imgs.length > 1) imgs[1].remove();
    const hdr = container.querySelector("div");
    if (hdr) Object.assign(hdr.style, { margin: "0", padding: "0", float: "none", display: "block" });
    const first = container.firstElementChild;
    if (first) Object.assign(first.style, { margin: "0", padding: "0" });
    Object.assign(container.style, { margin: "0", padding: "0" });

    // 2.4) Esperar al siguiente frame para asegurar render completo
    await new Promise(requestAnimationFrame);

    // 2.5) Forzar crossOrigin y esperar a que se carguen todas las imágenes
    const capture = document.getElementById("pdf-capture");
    const allImgs = capture.querySelectorAll("img");
    allImgs.forEach(img => (img.crossOrigin = "anonymous"));
    await Promise.all(
      Array.from(allImgs).map(
        img =>
          new Promise(resolve => {
            if (img.complete) return resolve();
            img.onload = resolve;
            img.onerror = resolve;
          })
      )
    );

    // 2.6) Capturar #pdf-capture en un solo canvas
    console.log("Iniciando html2canvas...");
    const canvas = await html2canvas(capture, {
      scale: 2,
      useCORS: true,
      allowTaint: false,
      scrollX: 0,
      scrollY: -window.scrollY,
      width: capture.offsetWidth,
      height: capture.scrollHeight,
    });
    console.log("Canvas capturado:", canvas.width, "×", canvas.height);

    // 2.7) Configurar jsPDF y paginar
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ unit: "mm", format: "letter", orientation: "portrait" });
    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();
    const margin = 5;
    const pdfW = pageW - margin * 2;
    const pdfH = pageH - margin * 2;
    const pxPerMm = canvas.width / pdfW;
    const pagePxH = Math.floor(pdfH * pxPerMm);

    const imgData = canvas.toDataURL("image/jpeg", 1.0);
    let renderedH = 0;
    let pageCount = 0;

    while (renderedH < canvas.height) {
      const fragH = Math.min(pagePxH, canvas.height - renderedH);
      const pageCanvas = document.createElement("canvas");
      pageCanvas.width = canvas.width;
      pageCanvas.height = fragH;
      pageCanvas.getContext("2d").drawImage(
        canvas,
        0,
        renderedH,
        canvas.width,
        fragH,
        0,
        0,
        canvas.width,
        fragH
      );

      const fragImg = pageCanvas.toDataURL("image/jpeg", 1.0);
      if (pageCount > 0) pdf.addPage();
      pdf.addImage(fragImg, "JPEG", margin, margin, pdfW, (fragH / canvas.width) * pdfW);

      renderedH += fragH;
      pageCount++;
    }

    // 2.8) Guardar PDF
    pdf.save("Contrato.pdf");
    console.log("PDF generado en", pageCount, "páginas");
  } catch (err) {
    console.error("Error exportando PDF:", err);
    alert("Error exportando PDF. Revisa la consola.");
  }
});

// Helper para cargar plantilla .docx
function loadFile(url) {
  return new Promise((resolve, reject) => {
    window.PizZipUtils.getBinaryContent(url, (err, data) =>
      err ? reject(err) : resolve(data)
    );
  });
}
