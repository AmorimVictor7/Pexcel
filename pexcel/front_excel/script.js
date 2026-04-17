// ================================
// CONFIG
// ================================

const siteUrl =
"https://pecege.sharepoint.com/:l:/r/sites/grupoweb/Lists/ListaAcoes?e=6Py78q";

const listaNome = "ListaAcoes";

const form = document.getElementById('form');
const botao = form.querySelector("button");

let enviando = false;


// ================================
// CARREGAR SELECTS
// (se vierem de lista SharePoint)
// ================================

window.onload = () => {

  carregarSelects();

};

async function carregarSelects() {

  try {

    const url =
      siteUrl +
      "/_api/web/lists/getbytitle('BaseSelects')/items";

    const response = await fetch(url, {
      method: "GET",
      headers: {
        "Accept": "application/json;odata=nometadata"
      }
    });

    const data = await response.json();

    preencherSelect("tipo_acao", data.value, "TipoAcao");
    preencherSelect("curso", data.value, "Curso");
    preencherSelect("campanha", data.value, "Campanha");
    preencherSelect("marca", data.value, "Marca");
    preencherSelect("ativo", data.value, "AcaoAtiva");

  } catch (error) {

    console.error("Erro ao carregar selects:", error);

  }

}


// ================================
// PREENCHER SELECT
// ================================

function preencherSelect(id, dados, campo) {

  const select = document.getElementById(id);

  const valoresUnicos =
    [...new Set(dados.map(item => item[campo]))];

  select.innerHTML =
    '<option value="">Selecione...</option>';

  valoresUnicos.forEach(valor => {

    if (valor) {

      let option =
        document.createElement("option");

      option.value = valor;
      option.textContent = valor;

      select.appendChild(option);

    }

  });

}


// ================================
// ENVIAR PARA SHAREPOINT
// ================================

form.addEventListener('submit', async function (e) {

  e.preventDefault();

  if (enviando) return;

  const data = {

    tipo_acao:
      document.getElementById('tipo_acao').value,

    curso:
      document.getElementById('curso').value,

    campanha:
      document.getElementById('campanha').value,

    marca:
      document.getElementById('marca').value,

    data_acao:
      document.getElementById('data_acao').value,

    ativo:
      document.getElementById('ativo').value,

    link:
      document.getElementById('link').value

  };

  enviando = true;

  botao.disabled = true;
  botao.innerText = "Enviando...";

  try {

    const url =
      siteUrl +
      "/_api/web/lists/getbytitle('" +
      listaNome +
      "')/items";

    const response = await fetch(url, {

      method: "POST",

      headers: {
        "Accept":
          "application/json;odata=nometadata",

        "Content-Type":
          "application/json;odata=nometadata"
      },

      body: JSON.stringify({

        Title: data.tipo_acao,
        Curso: data.curso,
        Campanha: data.campanha,
        Marca: data.marca,
        DataAcao: data.data_acao,
        Ativo: data.ativo,
        Link: data.link

      })

    });

    if (response.ok) {

      alert("Enviado com sucesso! ✅");

      form.reset();

    } else {

      alert("Erro ao salvar ❌");

    }

  } catch (error) {

    console.error(error);

    alert("Erro ao enviar ❌");

  }

  enviando = false;

  botao.disabled = false;
  botao.innerText = "Salvar Ação";

});


// ================================
// DARK MODE
// ================================

const toggle =
document.getElementById("toggleTheme");

if (localStorage.getItem("theme") === "dark") {

  document.body.classList.add("dark-mode");

}

toggle.addEventListener("click", () => {

  document.body.classList.toggle("dark-mode");

  if (document.body.classList.contains("dark-mode")) {

    localStorage.setItem("theme", "dark");

  } else {

    localStorage.setItem("theme", "light");

  }

});
