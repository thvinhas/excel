import config from '../config/categorias_config.json' with { type: "json" };
let finalvalue = [];
let last_upload = null;

// Função que será chamada quando ambos os arquivos forem processados
function onFinalValueReady() {
    new DataTable('#example', {
        columns: [
            { title: 'Data' },
            { title: 'Descrição' },
            { title: 'Valor' },
            { title: 'Conta' },
            { title: 'Categoria' }
        ],
        data: finalvalue,
        layout: {
            topStart: {
              buttons: [{
extend: 'excel',
title: ''
}, [  {
    extend: 'csvHtml5',
    fieldSeparator: ';'
    //..other options
  }]]
            }
        }
    });
    // Aqui você pode fazer o que precisar com o `finalvalue`
}

// Modificar formValidation para resolver as promessas

    document.getElementById('form').addEventListener('submit', function(e) {
        e.preventDefault(); // Evita o envio do formulário padrão
    
        const files = e.target.elements.arquivos.files; // Obtém todos os arquivos selecionados
        last_upload = e.target.elements.last_upload.value;
        if (files.length === 0) {
            alert('Por favor, selecione pelo menos um arquivo.');
            return;
        }
    
        // Cria uma lista de promessas para cada arquivo
        const filePromises = Array.from(files).map(file => processFile(file));
    
        // Quando todas as promessas forem resolvidas, chama a função
        Promise.all(filePromises)
            .then(() => {
                onFinalValueReady(); // Chama a função quando todos os arquivos forem processados
            })
            .catch(error => console.error("Erro ao processar os arquivos:", error));
    });


 
        // Função para processar cada arquivo
        function processFile(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
    
                reader.onload = function(event) {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array',  raw: false , cellText: true});
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const json = XLSX.utils.sheet_to_json(firstSheet, {raw: true } );

    
                    // Verifica se o arquivo é do tipo AIB ou Revolut
                    if (json[0].hasOwnProperty("Transaction Type")) {
                        generatetableAib(json);
                    } else {
                        generatetableRevolut(json);
                    }
                    resolve(); // Resolva a promessa ao concluir
                };
    
                reader.onerror = reject; // Em caso de erro
                reader.readAsArrayBuffer(file);
            });
        }

// Funções para gerar tabelas (mantidas as mesmas, com finalvalue sendo preenchido)
function generatetableAib(value) {
    let valor = 0;
    value.forEach(function(data) {
        if(data['Transaction Type'] == "Debit" || data['Transaction Type'] == "Direct Debit") {
            valor = `-${data[" Debit Amount"]}`
        }else {
            data[" Credit Amount"];
        }
       
        generateTable(data[' Posted Transactions Date'], data[' Description1'], valor, "AIB");
    });
}

function generatetableRevolut(value) {
    value.forEach(function(data) {
        generateTable(data['Started Date'], data['Description'], data['Amount'], "Revolut")
    });
}

function generateTable (dataTransacao, descricao, valor, banco) {
    let categoriasConfig = config.categorias;
    let transferenciasShow = config.transferencias_show;
    let transferenciasRemove = config.transferencias_remove;
     dataTransacao = dateFormat(dataTransacao,banco);
     descricao = cleanDescription(descricao);
    let categoria = foundCategory(descricao, categoriasConfig);

    if (valor == 0) {
        return;
    }

    if (isTransferencia(descricao, transferenciasShow)) {
        let transefersDiv = document.getElementById('trasnfers');
        let text = `Transferência do ${banco} para: ${descricao}, data: ${dataTransacao}, valor: ${valor}`;

        transefersDiv.insertAdjacentHTML('afterend', `<p >${text}</p>`);
        return; // Não insere na planilha final
    }else if (isTransferencia(descricao, transferenciasRemove)) {
        return;
    }

    finalvalue.push([dataTransacao, descricao, valor, banco, categoria]);
}

function checkIfTracaoIsMenor(dataTransacao, last_upload) {
    let [day, month, year] = dataTransacao.split('/')
    let dateObj = new Date("20"+year, +month - 1, +day)       
    return dateObj < last_upload;
}

function excelDateToJSDate(serial) {
    const daysSince1900 = serial - 25569; // O número de dias desde 1º de janeiro de 1970 (25569 = dias entre 1/1/1900 e 1/1/1970)
    const msPerDay = 86400000; // Milissegundos por dia
    const date = new Date(daysSince1900 * msPerDay); // Converte para timestamp e cria uma data
    const time = new Date((serial % 1) * msPerDay); // Converte a parte decimal (hora)
    
    // Combina a data e a hora
    const hours = time.getUTCHours();
    const minutes = time.getUTCMinutes();
    
    return `${date.toLocaleDateString()}`;
}

// Outras funções auxiliares
function isTransferencia(palavraChave, config) {
    return config.includes(palavraChave);
}

function foundCategory(description, categorias) {
    let categoria = "Sem Categoria";
    Object.keys(categorias).forEach(function(key) {
        if (categorias[key].includes(description)) {
           categoria = key;
        }
    });
    return categoria;
}

function dateFormat(date, type) {
    if (typeof date === 'string') {
        let [d, m, y] = date.split(/\D/);
        return `${d}/${m}/${y}`;
    }else{
        const excelDate = new Date((date - (25567 + 2)) * 86400 * 1000);
        // Formata a data no formato correto (DD/MM/AAAA)
        const day = excelDate.getDate().toString().padStart(2, "0");
        const month = (excelDate.getMonth() +1).toString().padStart(2, "0");
        const year = excelDate.getFullYear();
        if (type === 'Revolut') {
            return `${day}/${month}/${year}`;
        }else {
            return `${month}/${day}/${year}`;

        }
    }
}

function cleanDescription(text) {
    if (typeof text !== 'string') {
        return ''; // Se `text` não for uma string, retorna uma string vazia ou outro valor padrão
    }

    let values = ['VDC', 'VDP', '*MOBI', "VDA"] ;
    values.forEach(x => {
        text = text.replace(x, '');
    });
    text = text.replace(/^-/, '');
    return text.trimStart();
}
