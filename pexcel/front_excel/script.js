const url = "https://pecege.sharepoint.com/:l:/r/sites/grupoweb/Lists/ListaAcoes?e=5CQWdP";

const form = document.getElementById('form');
const botao = form.querySelector("button");

let enviando = false;


  // CARREGAR SELECTS

window.onload = () => {
  fetch(url)
    .then(res => res.json())
    .then(data => {
      console.log("Dados recebidos:", data);

      preencherSelect("tipo_acao", data.tipo_acao);
      preencherSelect("curso", data.curso);
      preencherSelect("campanha", data.campanha);
      preencherSelect("marca", data.marca);
      preencherSelect("ativo", data.ativo);
    })
    .catch(error => {
      console.error("Erro ao carregar selects:", error);
    });
};


   //FUNÇÃO PREENCHER SELECT

function preencherSelect(id, lista) {
  const select = document.getElementById(id);

  if (!lista) return;

  lista.forEach(item => {
    const option = document.createElement("option");
    option.value = item;
    option.textContent = item;
    select.appendChild(option);
  });
}


   //ENVIAR FORMULÁRIO

form.addEventListener('submit', function (e) {
  e.preventDefault();

  if (enviando) return; // bloqueia clique duplo
  enviando = true;

  // trava botão
  botao.disabled = true;
  botao.innerText = "Enviando...";

  const data = {
    tipo_acao: document.getElementById('tipo_acao').value,
    curso: document.getElementById('curso').value,
    campanha: document.getElementById('campanha').value,
    marca: document.getElementById('marca').value,
    data_acao: document.getElementById('data_acao').value,
    ativo: document.getElementById('ativo').value,
    link: document.getElementById('link').value
  };

  fetch(url, {
    method: "POST",
    body: JSON.stringify(data)
  })
    .then(response => response.text())
    .then(result => {
      console.log(result);

      alert("Enviado com sucesso! ✅");
      form.reset();
    })
    .catch(err => {
      console.error(err);
      alert("Erro ao enviar ❌");
    })
    .finally(() => {
      enviando = false;

      // libera botão
      botao.disabled = false;
      botao.innerText = "Salvar Ação";
    });
});


   //DARK MODE

const toggle = document.getElementById("toggleTheme");

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
