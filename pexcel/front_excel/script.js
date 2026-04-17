// ============================================
// CONFIGURAÇÃO DO SHAREPOINT
// ============================================
// ATENÇÃO: Você PRECISA configurar estas 3 variáveis:
const siteUrl = "https://[seu-tenant].sharepoint.com/sites/[seu-site]"; // URL do seu site SharePoint
const listaNome = "NomeDaSuaLista"; // Nome da sua lista no SharePoint
const accessToken = "SEU_TOKEN_AQUI"; // Veja explicação sobre autenticação abaixo

// URLs da API REST do SharePoint
const getUrl = `${siteUrl}/_api/web/lists/getbytitle('${listaNome}')/items`;
const postUrl = `${siteUrl}/_api/web/lists/getbytitle('${listaNome}')/items`;

// ============================================
// FUNÇÃO PARA FAZER REQUISIÇÕES AO SHAREPOINT
// ============================================
async function requestSharePoint(url, method, data = null) {
    const headers = {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'Authorization': 'Bearer ' + accessToken
    };

    const options = {
        method: method,
        headers: headers
    };

    if (data && method === 'POST') {
        options.body = JSON.stringify(data);
    }

    try {
        const response = await fetch(url, options);
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        return await response.json();
    } catch (error) {
        console.error('Erro na requisição SharePoint:', error);
        throw error;
    }
}

// ============================================
// CARREGAR SELECTS (SUBSTITUI O GET DO GOOGLE)
// ============================================
window.onload = async () => {
    try {
        console.log("Carregando dados do SharePoint...");
        
        // Busca todos os itens da lista
        const response = await requestSharePoint(getUrl, 'GET');
        const itens = response.d.results;
        
        console.log("Dados recebidos:", itens);
        
        // Extrai valores únicos para cada campo (como fazia o Google Sheets)
        const tipo_acao = [...new Set(itens.map(item => item.TipoAcao || item.Tipo_acao))];
        const curso = [...new Set(itens.map(item => item.Curso))];
        const campanha = [...new Set(itens.map(item => item.Campanha))];
        const marca = [...new Set(itens.map(item => item.Marca))];
        const ativo = [...new Set(itens.map(item => item.Ativo))];
        
        // Preenche os selects
        preencherSelect("tipo_acao", tipo_acao.filter(Boolean));
        preencherSelect("curso", curso.filter(Boolean));
        preencherSelect("campanha", campanha.filter(Boolean));
        preencherSelect("marca", marca.filter(Boolean));
        preencherSelect("ativo", ativo.filter(Boolean));
        
    } catch (error) {
        console.error("Erro ao carregar dados do SharePoint:", error);
        alert("Erro ao carregar os dados. Verifique o console para mais detalhes.");
    }
};

// ============================================
// FUNÇÃO PREENCHER SELECT (MESMA DO SEU CÓDIGO)
// ============================================
function preencherSelect(id, lista) {
    const select = document.getElementById(id);
    if (!select || !lista) return;
    
    // Limpa opções existentes (exceto a primeira se for placeholder)
    while (select.options.length > 0) {
        select.remove(0);
    }
    
    // Adiciona opção padrão
    const defaultOption = document.createElement("option");
    defaultOption.value = "";
    defaultOption.textContent = "Selecione...";
    select.appendChild(defaultOption);
    
    // Adiciona as opções da lista
    lista.forEach(item => {
        if (item && item.trim() !== "") {
            const option = document.createElement("option");
            option.value = item;
            option.textContent = item;
            select.appendChild(option);
        }
    });
}

// ============================================
// ENVIAR FORMULÁRIO (SUBSTITUI O POST DO GOOGLE)
// ============================================
const form = document.getElementById('form');
const botao = form ? form.querySelector("button") : null;
let enviando = false;

if (form && botao) {
    form.addEventListener('submit', async function (e) {
        e.preventDefault();
        
        if (enviando) return;
        enviando = true;
        
        // Trava botão
        botao.disabled = true;
        botao.innerText = "Enviando...";
        
        // Prepara os dados no formato que o SharePoint espera
        const formData = {
            __metadata: { type: `SP.Data.${listaNome.replace(/ /g, '')}ListItem` },
            Title: document.getElementById('tipo_acao')?.value || "Ação", // Título é obrigatório no SharePoint
            TipoAcao: document.getElementById('tipo_acao')?.value || "",
            Curso: document.getElementById('curso')?.value || "",
            Campanha: document.getElementById('campanha')?.value || "",
            Marca: document.getElementById('marca')?.value || "",
            DataAcao: document.getElementById('data_acao')?.value || null,
            Ativo: document.getElementById('ativo')?.value || "",
            Link: document.getElementById('link')?.value || ""
        };
        
        try {
            // IMPORTANTE: Você precisa criar estas colunas no SharePoint com os nomes exatos:
            // Title (padrão), TipoAcao, Curso, Campanha, Marca, DataAcao, Ativo, Link
            
            const response = await requestSharePoint(postUrl, 'POST', formData);
            console.log("Resposta do SharePoint:", response);
            
            alert("Enviado com sucesso! ✅");
            form.reset();
            
            // Recarrega os selects para mostrar novos dados (opcional)
            // window.location.reload(); // ou chama a função de carregar novamente
            
        } catch (error) {
            console.error("Erro detalhado:", error);
            alert(`Erro ao enviar: ${error.message} ❌\nVerifique o console para mais detalhes.`);
        } finally {
            enviando = false;
            botao.disabled = false;
            botao.innerText = "Salvar Ação";
        }
    });
}

// ============================================
// DARK MODE (MESMO DO SEU CÓDIGO)
// ============================================
const toggle = document.getElementById("toggleTheme");

if (localStorage.getItem("theme") === "dark") {
    document.body.classList.add("dark-mode");
}

if (toggle) {
    toggle.addEventListener("click", () => {
        document.body.classList.toggle("dark-mode");
        
        if (document.body.classList.contains("dark-mode")) {
            localStorage.setItem("theme", "dark");
        } else {
            localStorage.setItem("theme", "light");
        }
    });
}
