// ========================================
// DADOS GLOBAIS
// ========================================

let dadosOriginais = [];
let dadosFiltrados = [];
let agrupado = {};
let mapaApanhas = {};
let mapaPulmoes = {};

// ========================================
// PROCESSAR ARQUIVO
// ========================================

async function processarArquivo(){

    const arquivo =
    document.getElementById(
        "arquivo"
    ).files[0];

    if(!arquivo){

        alert(
            "Selecione o arquivo principal."
        );

        return;
    }

    try{

        // Carrega as apanhas se houver arquivo
        if(
            document.getElementById(
                "arquivoApanha"
            )?.files[0]
        ){

            await carregarApanhas();

        }
if(
    document.getElementById(
        "arquivoPulmao"
    )?.files[0]
){

    await carregarPulmoes();

}
        
        const nome =
        arquivo.name.toLowerCase();

        if(
            nome.endsWith(".csv")
        ){

            lerCSV(
                arquivo
            );

        }else{

            lerExcel(
                arquivo
            );

        }

    }catch(erro){

        console.error(
            erro
        );

        alert(
            "Erro ao ler arquivo."
        );

    }

}
// ========================================
// LEITURA CSV
// ========================================

function lerCSV(arquivo){

    const reader =
    new FileReader();

    reader.onload =
    function(e){

        const texto =
        e.target.result;

        const linhas =
        texto
        .split(/\r?\n/)
        .filter(
            l=>l.trim()
        );

        const cabecalho =
        linhas[0]
        .split(";");

        const dados = [];

        for(
            let i = 1;
            i < linhas.length;
            i++
        ){

            const valores =
            linhas[i]
            .split(";");

            const obj = {};

            cabecalho.forEach(
                (col,idx)=>{

                    obj[
                        col.trim()
                    ] =
                    valores[idx]
                    ? valores[idx].trim()
                    : "";

                }
            );

            dados.push(obj);

        }

        console.log(
            "COLUNAS CSV:",
            Object.keys(dados[0])
        );

        tratarDados(dados);

    };

    reader.readAsText(
        arquivo,
        "latin1"
    );

}
// ========================================
// LEITURA XLSX
// ========================================

function lerExcel(
    arquivo
){

    const reader =
    new FileReader();

    reader.onload =
    function(e){

        const data =
        new Uint8Array(
            e.target.result
        );

        const workbook =
        XLSX.read(
            data,
            {
                type:"array"
            }
        );

        const aba =
        workbook.SheetNames[0];

        const dados =
        XLSX.utils.sheet_to_json(
            workbook.Sheets[aba],
            {
                defval:""
            }
        );

        tratarDados(
            dados
        );

    };

    reader.readAsArrayBuffer(
        arquivo
    );

}

// ========================================
// LEITURA APANHAS
// ========================================

async function carregarApanhas(){

    const arquivo =
    document.getElementById(
        "arquivoApanha"
    ).files[0];

    if(!arquivo) return;

    const reader =
    new FileReader();

    return new Promise(resolve=>{

        reader.onload =
        function(e){

            const texto =
            e.target.result;

            const linhas =
            texto
            .split(/\r?\n/)
            .filter(
                l => l.trim()
            );

            mapaApanhas = {};

            for(
                let i = 1;
                i < linhas.length;
                i++
            ){

                const colunas =
                linhas[i]
                .split(";");

                const endereco =
                colunas[0]?.trim();

                const sku =
                colunas[1]?.trim();

                if(sku){

                    mapaApanhas[sku] =
                    endereco;

                }

            }

            console.log(
                "Apanhas carregadas:",
                Object.keys(
                    mapaApanhas
                ).length
            );

            console.log(mapaApanhas);
            
            resolve();

        };

        reader.readAsText(
            arquivo,
            "latin1"
        );

    });

}

// ========================================
// LEITURA PULMÕES
// ========================================

async function carregarPulmoes(){

    const arquivo =
    document.getElementById(
        "arquivoPulmao"
    ).files[0];

    if(!arquivo) return;

    const reader =
    new FileReader();

    return new Promise(resolve=>{

        reader.onload =
        function(e){

            const texto =
            e.target.result;

            const linhas =
            texto
            .split(/\r?\n/)
            .filter(
                l => l.trim()
            );

            mapaPulmoes = {};

            for(
                let i = 1;
                i < linhas.length;
                i++
            ){

                const colunas =
                linhas[i].split(";");

                const rua =
                colunas[1]?.trim();

                const predio =
                colunas[2]?.trim();

                const apartamento =
                colunas[3]?.trim();

                const sala =
                colunas[4]?.trim();

                const sku =
                colunas[5]?.trim();

                if(!sku) continue;

                const endereco =
                `${rua}.${predio}.${apartamento}.${sala}`;

                if(!mapaPulmoes[sku]){

                    mapaPulmoes[sku] = [];

                }

                if(
                    !mapaPulmoes[sku]
                    .includes(endereco)
                ){

                    mapaPulmoes[sku]
                    .push(endereco);

                }

            }

            console.log(
                "Pulmões carregados:",
                Object.keys(
                    mapaPulmoes
                ).length
            );

            resolve();

        };

        reader.readAsText(
            arquivo,
            "latin1"
        );

    });

}



// ========================================
// TRATAMENTO INICIAL
// ========================================

function tratarDados(dados){

    console.log("COLUNAS:");
    console.log(Object.keys(dados[0]));

    dadosOriginais =
    dados.map(linha=>{

        const loja =
        String(
            linha["tipEspecie"] ||
            linha["Espécie"] ||
            ""
        )
        .replace(
            /^S/i,
            ""
        );

        const skuPrincipal =
        String(
            linha["Código do produto"]
        ).trim();

        console.log(
            "SKU:",
            skuPrincipal,
            "Apanha encontrada:",
            mapaApanhas[skuPrincipal]
        );

        return{

            loja,

            ptl:
            linha["(Palete)Posição"] ||
            linha["Posição"] ||
            "",

            sku:
            skuPrincipal,

            descricao:
            linha["Produto"] ||
            "",

            volumes:
            Number(
                linha["quantidadeTotal"]
            ) || 0,

            master:
            linha["(Palete)Master"] ||
            "",

            apanha:
            mapaApanhas[skuPrincipal] ||
            "Sem Apanha"

        };

    });

    dadosFiltrados =
    [...dadosOriginais];

    gerarAgrupamento();

}

// ========================================
// AGRUPAMENTO
// ========================================

function gerarAgrupamento(){

    agrupado = {};

    dadosFiltrados.forEach(item=>{

        const loja =
        item.loja || "SEM LOJA";

        const ptl =
        item.ptl || "SEM PTL";

        if(!agrupado[loja]){

            agrupado[loja] = {};

        }

        if(!agrupado[loja][ptl]){

            agrupado[loja][ptl] = [];

        }

        agrupado[loja][ptl].push(
            item
        );

    });

    atualizarKPIs();

    renderizar();

}

// ========================================
// KPIs
// ========================================

function atualizarKPIs(){

    const lojas =
    Object.keys(
        agrupado
    ).length;

    let ptls = 0;

    let volumes = 0;

    const skusUnicos =
    new Set();

    Object.values(
        agrupado
    ).forEach(loja=>{

        ptls +=
        Object.keys(
            loja
        ).length;

        Object.values(
            loja
        ).forEach(ptl=>{

            ptl.forEach(item=>{

                if(item.sku){

                    skusUnicos.add(
                        String(item.sku).trim()
                    );

                }

                volumes++;

            });

        });

    });

    document.getElementById(
        "kpiLojas"
    ).textContent = lojas;

    document.getElementById(
        "kpiPtls"
    ).textContent = ptls;

    document.getElementById(
        "kpiSkus"
    ).textContent =
    skusUnicos.size;

    document.getElementById(
        "kpiVolumes"
    ).textContent =
    volumes.toLocaleString(
        "pt-BR"
    );

}
// ========================================
// FILTROS
// ========================================

function aplicarFiltros(){

    const lojaFiltro =
    document
    .getElementById(
        "filtroLoja"
    )
    .value
    .toLowerCase();

    const ptlFiltro =
    document
    .getElementById(
        "filtroPTL"
    )
    .value
    .toLowerCase();

    const skuFiltro =
    document
    .getElementById(
        "filtroSKU"
    )
    .value
    .toLowerCase();

    dadosFiltrados =
    dadosOriginais.filter(
        item=>{

            return (
                item.loja
                    .toLowerCase()
                    .includes(lojaFiltro)
                &&
                item.ptl
                    .toLowerCase()
                    .includes(ptlFiltro)
                &&
                String(item.sku)
                    .toLowerCase()
                    .includes(skuFiltro)
            );

        }
    );

    agrupado = {};

    dadosFiltrados.forEach(item=>{

        if(!agrupado[item.loja]){

            agrupado[item.loja] = {};

        }

        if(!agrupado[item.loja][item.ptl]){

            agrupado[item.loja][item.ptl] = [];

        }

        agrupado[item.loja][item.ptl]
            .push(item);

    });

    atualizarKPIs();

    renderizar();

}

// ========================================
// RENDERIZAÇÃO
// ========================================

function renderizar(){

    const resultado =
    document.getElementById(
        "resultado"
    );

    resultado.innerHTML = "";

    const lojas =
    Object.keys(
        agrupado
    )
    .sort(
        (
            a,
            b
        )=>
        Number(a) -
        Number(b)
    );

    lojas.forEach(loja=>{

        const cardLoja =
        document.createElement(
            "div"
        );

        cardLoja.className =
        "loja-card";

        cardLoja.innerHTML =
        `
        <div class="loja-titulo">
            🏪 LOJA ${loja}
        </div>
        `;

        const ptls =
        Object.keys(
            agrupado[loja]
        ).sort();

        ptls.forEach(ptl=>{

            const itens =
            agrupado[
                loja
            ][ptl];

            let totalVolumes =
            0;

            itens.forEach(i=>{

                totalVolumes +=
                Number(
                    i.volumes
                ) || 0;

            });

            let htmlTabela =
            `
            <div class="ptl-card">

                <div class="ptl-titulo">
                    📦 ${ptl}
                </div>

                <table class="tabela">

                <thead>
                    <tr>
                        <th>SKU</th>
                        <th>Descrição</th>
<th>Apanha</th>
<th>Volumes</th>
                    </tr>
                </thead>

                <tbody>
            `;

const skuAgrupado = {};

itens.forEach(item=>{

    const chave =
    item.sku;

    if(!skuAgrupado[chave]){

       skuAgrupado[chave] = {

    sku: item.sku,

    descricao:
    item.descricao,

    apanha:
    item.apanha || "Sem Apanha",

    volumes: 0

};

    }

    skuAgrupado[chave]
    .volumes += 1;

});

Object.values(
    skuAgrupado
).forEach(item=>{

    htmlTabela +=
    `
    <tr>
        <td>${item.sku}</td>
        <td>${item.descricao}</td>
<td>${item.apanha}</td>
<td>${item.volumes}</td>
    </tr>
    `;

});
            

            htmlTabela +=
            `
                </tbody>
                </table>

                <div class="resumo-ptl">

                   SKUs:
${Object.keys(skuAgrupado).length}

                    |

                    Volumes:
                    ${totalVolumes}

                </div>

            </div>
            `;

            cardLoja.innerHTML +=
            htmlTabela;

        });

        resultado.appendChild(
            cardLoja
        );

    });

}

// ========================================
// TEXTO WHATSAPP
// ========================================

function gerarTextoWhatsapp(){

    let texto =
`🚨 PENDÊNCIAS PTL

`;

    let totalLojas = 0;
    let totalPTLs = 0;
    let totalVolumes = 0;

    const skusUnicos =
    new Set();

    const lojas =
    Object.keys(agrupado)
    .sort(
        (a,b)=>
        Number(a)-Number(b)
    );

    lojas.forEach(loja=>{

        totalLojas++;

        texto +=
`🏪 LOJA ${loja}

`;

        const ptls =
        Object.keys(
            agrupado[loja]
        ).sort();

        ptls.forEach(ptl=>{

            totalPTLs++;

            const itens =
            agrupado[loja][ptl];

            let volumesPTL = 0;

            itens.forEach(item=>{

                if(item.sku){

                    skusUnicos.add(
                        String(item.sku).trim()
                    );

                }

                volumesPTL++;

            });

            totalVolumes +=
            volumesPTL;

            texto +=
`📦 ${ptl}
SKUs: ${new Set(itens.map(i => i.sku)).size}
Volumes: ${volumesPTL}

`;

        });

        texto +=
`────────────────

`;

    });

    texto +=
`📊 RESUMO

Lojas: ${totalLojas}
PTLs: ${totalPTLs}
SKUs: ${skusUnicos.size}
Volumes: ${totalVolumes.toLocaleString("pt-BR")}
`;

    document.getElementById(
        "textoWhatsapp"
    ).value = texto;

}

// ========================================
// COPIAR WHATSAPP
// ========================================

function copiarWhatsapp(){

    gerarTextoWhatsapp();

    const campo =
    document.getElementById(
        "textoWhatsapp"
    );

    campo.select();

    document.execCommand(
        "copy"
    );

    alert(
        "Texto copiado!"
    );

}

// ========================================
// EXPORTAR EXCEL
// ========================================

function exportarExcel(){

    const exportacao = [];

    Object.keys(
        agrupado
    ).forEach(loja=>{

        Object.keys(
            agrupado[loja]
        ).forEach(ptl=>{

            agrupado[
                loja
            ][ptl]
            .forEach(item=>{

                exportacao.push({

                    Loja:
                    item.loja,

                    PTL:
                    item.ptl,

                    SKU:
                    item.sku,

                    Descricao:
                    item.descricao,

                    Volumes:
                    item.volumes,

                    Master:
                    item.master

                });

            });

        });

    });

    const ws =
    XLSX.utils.json_to_sheet(
        exportacao
    );

    const wb =
    XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
        wb,
        ws,
        "Pendencias"
    );

    XLSX.writeFile(
        wb,
        "Pendencia_PTL.xlsx"
    );

}

// ========================================
// RESUMO AUTOMÁTICO
// ========================================

function atualizarResumo(){

    gerarTextoWhatsapp();

}

// ========================================
// PROCESSAMENTO FINAL
// ========================================

const gerarAgrupamentoOriginal =
gerarAgrupamento;

gerarAgrupamento =
function(){

    gerarAgrupamentoOriginal();

    atualizarResumo();

};

// ========================================
// ATALHO ENTER FILTROS
// ========================================

document
.querySelectorAll(
    ".filtros input"
)
.forEach(input=>{

    input.addEventListener(
        "keyup",
        ()=>{
            aplicarFiltros();
        }
    );

});

// ========================================
// CARREGAMENTO
// ========================================

window.onload =
function(){

    document.getElementById(
        "textoWhatsapp"
    ).value =
`Faça upload do arquivo para gerar o relatório.`;

};


function copiarWhatsappCompleto(){
    copiarWhatsapp();
}

function copiarWhatsappResumo(){

    const textoOriginal =
    document.getElementById(
        "textoWhatsapp"
    ).value;

    const linhas =
    textoOriginal.split("\n");

    const resumo = [];

    resumo.push(
        "🚨 PENDÊNCIAS PTL\n"
    );

    const resumoGeral =
    linhas.find(
        l=>l.includes("Lojas:")
    );

    const ptls =
    linhas.find(
        l=>l.includes("PTLs:")
    );

    const skus =
    linhas.find(
        l=>l.includes("SKUs:")
    );

    const volumes =
    linhas.find(
        l=>l.includes("Volumes:")
    );

    resumo.push(
        "📊 RESUMO GERAL\n"
    );

    if(resumoGeral)
    resumo.push(resumoGeral);

    if(ptls)
    resumo.push(ptls);

    if(skus)
    resumo.push(skus);

    if(volumes)
    resumo.push(volumes);

    navigator.clipboard.writeText(
        resumo.join("\n")
    );

    alert(
        "Resumo copiado!"
    );

}

function baixarImagemResumo(){

    const card =
    document.createElement(
        "div"
    );

    card.style.width =
    "900px";

    card.style.padding =
    "30px";

    card.style.background =
    "#ffffff";

    card.style.color =
    "#000";

    card.style.fontFamily =
    "Segoe UI";

    card.innerHTML = `

    <h1>
    🚨 Pendência PTL
    </h1>

    <pre style="font-size:20px">
${document.getElementById(
"textoWhatsapp"
).value}
    </pre>

    `;

    document.body.appendChild(
        card
    );

    html2canvas(card)
    .then(canvas=>{

        const link =
        document.createElement(
            "a"
        );

        link.download =
        "pendencia-ptl.png";

        link.href =
        canvas.toDataURL(
            "image/png"
        );

        link.click();

        card.remove();

    });

}


function copiarWhatsappCompleto(){

    gerarTextoWhatsapp();

    const texto =
    document.getElementById(
        "textoWhatsapp"
    ).value;

    navigator.clipboard.writeText(
        texto
    );

    alert(
        "Relatório completo copiado!"
    );

}


function copiarWhatsappResumo(){

    gerarTextoWhatsapp();

    const textoOriginal =
    document.getElementById(
        "textoWhatsapp"
    ).value;

    const linhas =
    textoOriginal.split("\n");

    const resumo = [];

    resumo.push(
        "🚨 PENDÊNCIAS PTL\n"
    );

    linhas.forEach(linha=>{

        if(
            linha.includes("Lojas:")
            ||
            linha.includes("PTLs:")
            ||
            linha.includes("SKUs:")
            ||
            linha.includes("Volumes:")
        ){

            resumo.push(linha);

        }

    });

    navigator.clipboard.writeText(
        resumo.join("\n")
    );

    alert(
        "Resumo copiado!"
    );

}

async function baixarImagemResumo(){

    const blocos = [];

    Object.keys(agrupado)
    .sort((a,b)=>Number(a)-Number(b))
    .forEach(loja=>{

        Object.keys(
            agrupado[loja]
        )
        .sort()
        .forEach(ptl=>{

            const itens =
            agrupado[loja][ptl];

            const skus =
            new Set(
                itens.map(
                    i => i.sku
                )
            ).size;

            const volumes =
            itens.length;

            blocos.push(
`🏪 LOJA ${loja}
📦 ${ptl}
SKUs: ${skus}
Volumes: ${volumes}`
            );

        });

    });

    const itensPorImagem = 10;

    const paginas = [];

    for(
        let i = 0;
        i < blocos.length;
        i += itensPorImagem
    ){

        paginas.push(
            blocos.slice(
                i,
                i + itensPorImagem
            )
        );

    }

    const totalPaginas =
    Math.min(
        paginas.length,
        18
    );

    for(
        let pagina = 0;
        pagina < totalPaginas;
        pagina++
    ){

        const card =
        document.createElement(
            "div"
        );

        card.style.width =
        "1200px";

        card.style.padding =
        "30px";

        card.style.background =
        "#fff";

        card.style.color =
        "#000";

        card.style.fontFamily =
        "Segoe UI";

        card.innerHTML = `
            <h1>🚨 Pendência PTL</h1>
            <h3>Página ${pagina+1}/${totalPaginas}</h3>

            <pre style="
                font-size:22px;
                line-height:1.5;
                white-space:pre-wrap;
            ">
${paginas[pagina].join("\n\n")}
            </pre>
        `;

        document.body.appendChild(
            card
        );

        const canvas =
        await html2canvas(
            card,
            {
                scale:2
            }
        );

        const link =
        document.createElement(
            "a"
        );

        link.download =
        `pendencia-ptl-${String(
            pagina+1
        ).padStart(
            2,
            "0"
        )}.png`;

        link.href =
        canvas.toDataURL(
            "image/png"
        );

        link.click();

        card.remove();

        await new Promise(
            r=>setTimeout(
                r,
                300
            )
        );

    }

}


function imprimirPorVolume(){

    const dados = [];

    Object.keys(agrupado).forEach(loja=>{

        Object.keys(agrupado[loja]).forEach(ptl=>{

            const itens = agrupado[loja][ptl];

            const skuAgrupado = {};

            itens.forEach(item=>{

                if(!skuAgrupado[item.sku]){

                    skuAgrupado[item.sku] = {
                        loja: item.loja,
                        ptl: item.ptl,
                        sku: item.sku,
                        descricao: item.descricao,
                        apanha: item.apanha || "Sem Apanha",
                        volumes: 0
                    };

                }

                skuAgrupado[item.sku].volumes++;

            });

            dados.push(...Object.values(skuAgrupado));

        });

    });

    const grupos = {
        "🔴 Acima de 50 Volumes":
            dados.filter(x => x.volumes > 50),

        "🟠 De 20 a 49 Volumes":
            dados.filter(x => x.volumes >= 20 && x.volumes <= 49),

        "🟡 De 10 a 19 Volumes":
            dados.filter(x => x.volumes >= 10 && x.volumes <= 19),

        "🟢 Até 9 Volumes":
            dados.filter(x => x.volumes <= 9)
    };

    let html = `
    <html>
    <head>
        <style>
            body{
                font-family:Arial;
                padding:20px;
            }

            h2{
                background:#eee;
                padding:10px;
            }

            table{
                width:100%;
                border-collapse:collapse;
                margin-bottom:30px;
            }

            th,td{
                border:1px solid #ccc;
                padding:8px;
            }
        </style>
    </head>

    <body>

    <h1>📦 Pendências por Volume</h1>
    `;

    for(const titulo in grupos){

        if(grupos[titulo].length === 0)
            continue;

        html += `<h2>${titulo}</h2>`;

        html += `
        <table>
            <tr>
                <th>Loja</th>
                <th>PTL</th>
                <th>SKU</th>
                <th>Descrição</th>
                <th>Apanha</th>
                <th>Volumes</th>
            </tr>
        `;

        grupos[titulo]
        .sort((a,b)=>b.volumes-a.volumes)
        .forEach(item=>{

            html += `
            <tr>
                <td>${item.loja}</td>
                <td>${item.ptl}</td>
                <td>${item.sku}</td>
                <td>${item.descricao}</td>
                <td>${item.apanha || "Sem Apanha"}</td>
                <td>${item.volumes}</td>
            </tr>
            `;

        });

        html += `</table>`;
    }

    html += `
    </body>
    </html>
    `;

    const janela = window.open("", "_blank");

    janela.document.write(html);
    janela.document.close();

    setTimeout(()=>{
        janela.print();
    },500);

}
