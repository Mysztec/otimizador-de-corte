document.getElementById("uploadForm").addEventListener("submit", async function(e) {
    e.preventDefault();

    const formData = new FormData(this);

    document.getElementById("status").innerText = "Processando PDF...";

    const response = await fetch("/organizar", {
        method: "POST",
        body: formData
    });

    if (response.ok) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);

        const a = document.createElement("a");
        a.href = url;
        a.download = "pdf_organizado.pdf";
        document.body.appendChild(a);
        a.click();
        a.remove();

        document.getElementById("status").innerText = "PDF organizado com sucesso!";
    } else {
        document.getElementById("status").innerText = "Erro ao processar o PDF.";
    }
});