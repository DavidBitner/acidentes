// Validate Noc
function validateNoc() {
  const nOc = document.getElementById('nOc');
  nOc.value = nOc.value.toUpperCase();

  const nOcError = document.getElementById('nOcError');
  const nOcPattern = /^6C\d{4}$/;

  if (!nOcPattern.test(nOc.value)) {
    nOc.classList.add('invalid');
    nOc.classList.remove('valid');
    nOcError.style.display = 'inline';
  } else {
    nOc.classList.remove('invalid');
    nOc.classList.add('valid');
    nOcError.style.display = 'none';
  }
}

// Validate Date (only check if the field is filled correctly)
function validateDate() {
  const dateField = document.getElementById('date');
  const dateError = document.getElementById('dateError');

  // If the date field is empty or the browser didn't accept it, mark it as invalid
  if (!dateField.value) {
    dateField.classList.add('invalid');
    dateField.classList.remove('valid');
    dateError.style.display = 'inline';
  } else {
    dateField.classList.remove('invalid');
    dateField.classList.add('valid');
    dateError.style.display = 'none';
  }
}

// Validate Ocorrencia (Free-text)
function validateOcorrencia() {
  const ocorrencia = document.getElementById('ocorrencia');
  ocorrencia.value = ocorrencia.value.toUpperCase();

  const isValid = ocorrencia.value.trim() !== '';
  const ocorrenciaError = document.getElementById("ocorrenciaError");

  if (!isValid) {
    ocorrencia.classList.add('invalid');
    ocorrencia.classList.remove('valid');
    ocorrenciaError.style.display = 'inline';
  } else {
    ocorrencia.classList.remove('invalid');
    ocorrencia.classList.add('valid');
    ocorrenciaError.style.display = 'none';
  }
}

document.getElementById('nOc').addEventListener('blur', validateNoc);
document.getElementById('ocorrencia').addEventListener('blur', validateOcorrencia);
document.getElementById('date').addEventListener('blur', validateDate);

document.addEventListener("DOMContentLoaded", async function () {
  // Save and restore only <input> and <textarea> fields
  const inputs = document.querySelectorAll("input, textarea");

  inputs.forEach(input => {
    const savedValue = localStorage.getItem(input.id);
    if (savedValue) {
      input.value = savedValue;
    }

    input.addEventListener("input", () => {
      localStorage.setItem(input.id, input.value.trim());
    });
  });
});

// Clear Fields
document.getElementById("clear").addEventListener("click", function () {
  localStorage.clear();
  document.querySelectorAll('input, textarea').forEach(input => {
    input.value = '';
  });
  document.querySelector("#excel").classList.add("hidden");
});

function clearStorage() {
  localStorage.clear();
  location.reload(); // Reload to reflect cleared values
}

// Modal notes
const modal = document.getElementById("modal");
const notesButton = document.getElementById("notes");

notesButton.addEventListener("click", function () {
  modal.classList.add("show");
});

modal.addEventListener("click", function (event) {
  const modalContent = document.querySelector(".modal-content");
  if (!modalContent.contains(event.target)) {
    modal.classList.remove("show");
  }
});

// Word Generating
async function loadTemplate() {
  try {
    const response = await fetch('template.docx'); // Load default template
    if (!response.ok) throw new Error("Failed to load template");
    return await response.arrayBuffer();
  } catch (error) {
    console.error("Error loading template:", error);
    alert("Error loading template: " + error.message);
    throw error;
  }
}

// Generate Ocorrência
document.getElementById("generateWord").addEventListener("click", async function () {
  try {
    const nOc = document.getElementById("nOc").value.toUpperCase();
    const ocorrencia = document.getElementById("ocorrencia").value.toUpperCase();
    let dateValue = document.getElementById("date").value;
    let dateParts = dateValue.split("-");
    let date = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]).toLocaleDateString("pt-BR").toUpperCase();
    const inicio = document.getElementById("inicio").value;
    let termino = document.getElementById("termino").value;
    const logradouro = document.getElementById("logradouro").value.toUpperCase();
    let numero = document.getElementById("numero").value.toUpperCase();
    const bairro = document.getElementById("bairro").value.toUpperCase();
    const inicioFato = document.getElementById("inicioFato").value.trim().toUpperCase();
    let motivo = document.getElementById("motivo").value.trim().toUpperCase();
    let respObra = document.getElementById("respObra").value.trim().toUpperCase();
    let desviosbc = document.getElementById("desviosbc").value.trim().toUpperCase();
    let desvioscb = document.getElementById("desvioscb").value.trim().toUpperCase();
    let linhasAfetadas = document.getElementById("linhasAfetadas").value.trim().toUpperCase();
    let alerta = document.getElementById("alerta").value.toUpperCase();
    let linha = document.getElementById("linha").value.toUpperCase();
    let ocSptrans = document.getElementById("ocSptrans").value.toUpperCase();
    let contato = document.getElementById("contato").value.toUpperCase();
    const cco = document.getElementById("cco").value.toUpperCase();
    let operacional = document.getElementById("operacional").value.toUpperCase();

    if (!nOc) {
      alert("Insira um numero de ocorrência. Ex: 6C1234");
      return;
    }
    if (!ocorrencia) {
      alert("Insira um tipo de ocorrência. Ex: Instabilidade Sistema SIM...");
      return;
    }
    if (!linha) {
      linha = "NÃO HOUVE"
    }
    if (!date) {
      alert("Insira uma data válida para a ocorrência");
      return;
    }
    if (!inicio) {
      alert("Insira o horário no qual a interferência se iniciou");
      return;
    }
    if (!termino) {
      termino = "-";
    }
    if (!logradouro) {
      alert("Insira um endereço para o ocorrido");
      return;
    }
    if (!numero) {
      alert("Insira um numero para o logradouro");
      return;
    }
    if (!bairro) {
      alert("Insira um bairro para o logradouro");
      return;
    }
    if (!inicioFato) {
      alert("Insira corpo da ocorrência");
      return;
    }
    if (!motivo) {
      motivo = "NÃO CABE"
    }
    if (!respObra) {
      respObra = "NÃO CABE"
    }
    if (!desviosbc) {
      desviosbc = "NÃO HOUVE"
    }
    if (!desvioscb) {
      desvioscb = "NÃO HOUVE"
    }
    if (!linhasAfetadas) {
      linhasAfetadas = "NÃO HOUVE"
    }
    if (!ocSptrans) {
      ocSptrans = "NÃO HOUVE"
    }
    if (!contato) {
      contato = "NÃO HOUVE"
    }
    if (!alerta) {
      alerta = "NÃO HOUVE"
    }
    if (!cco) {
      alert("Insira o nome do responsável do CCO pela elaboração da ocorrência");
      return;
    }
    if (!operacional) {
      operacional = "NÃO HOUVE"
    }

    // Load the template
    const content = await loadTemplate();
    const zip = new PizZip(content);
    const doc = new docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

    doc.setData(
      {
        nOc: nOc + "/2025",
        ocorrencia: ocorrencia,
        date: date,
        inicio: inicio,
        termino: termino,
        logradouro: logradouro,
        numero: numero,
        bairro: bairro,
        inicioFato: inicioFato,
        motivo: motivo,
        respObra: respObra,
        desviosbc: desviosbc,
        desvioscb: desvioscb,
        linhasAfetadas: linhasAfetadas,
        alerta: alerta,
        linha: linha,
        ocSptrans: ocSptrans,
        contato: contato,
        cco: cco,
        operacional: operacional,
      }); // Replacing {nOc} in the template
    doc.render();

    const blob = new Blob([doc.getZip().generate({ type: "arraybuffer" })], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    const link = document.createElement("a");
    const nomeOcorrencia = `${nOc} - ${date.replace(/\//g, '.').slice(0, 5)}${linha != "NÃO HOUVE" ? " - " + linha : ""} - ${ocorrencia} - ${logradouro}`
    link.href = URL.createObjectURL(blob);
    link.download = `${nomeOcorrencia}.docx`;
    link.click();

    document.getElementById("td-nOc").textContent = nOc + "/2025";
    document.getElementById("td-date").textContent = date;
    document.getElementById("td-alerta").textContent = alerta;
    document.getElementById("td-ocorrencia").textContent = ocorrencia;
    document.getElementById("td-logradouro").textContent = logradouro;
    document.getElementById("td-operacional").textContent = operacional;
    document.getElementById("td-ocSptrans").textContent = ocSptrans;
    document.getElementById("td-contato").textContent = contato;
    document.getElementById("td-termino").textContent = termino;
    document.getElementById("td-cco").textContent = cco;
    document.getElementById("td-fechamento").textContent = cco;
    document.getElementById('excel').classList.remove('hidden');
  } catch (error) {
    console.error("Error generating document:", error);
    alert("Error generating document: " + error.message);
  }
});

document.getElementById("copy").addEventListener("click", function () {
  let row = document.querySelector("#excel tbody tr"); // Select the first row
  if (!row) return;

  let text = Array.from(row.querySelectorAll("td"))
    .map(td => td.innerText)
    .join("\t"); // Format for Excel

  navigator.clipboard.writeText(text).then(() => {
    alert("Linha copiada!");
  }).catch(err => {
    console.error("Erro ao copiar:", err);
  });
});
