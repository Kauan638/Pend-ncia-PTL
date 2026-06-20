// ========================================
// DADOS GLOBAIS
// ========================================

let dadosOriginais = [];
let dadosFiltrados = [];
let agrupado = {};

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
            "Selecione um arquivo."
        );

        return;
    }

    try{

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

function lerCSV(
    arquivo
){

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
        .split("\t");

        const dados =
        [];

        for(
            let i=1;
            i<linhas.length;
            i++
        ){

            const valores =
            linhas[i]
            .split("\t");

            const obj = {};

            cabecalho.forEach(
                (
                    col,
                    idx
                )=>{

                    obj[
                        col.trim()
                    ] =
                    valores[idx]
                    ? valores[idx].trim()
                    : "";

                }
            );

            dados.push(
                obj
            );

        }

        tratarDados(
            dados
        );

    };

    reader.readAsText(
        arquivo,
        "utf-8"
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
// TRATAMENTO INICIAL
// ========================================

function tratarDados(
    dados
){

    dadosOriginais =
    dados.map(
        linha=>{

            const loja =
            String(
                linha[
                    "tipEspecie"
                ] || ""
            )
            .replace(
                /^S/i,
                ""
            );

            return{

                loja,

                ptl:
                linha[
                    "(Palete)Posição"
                ] || "",

                sku:
                linha[
                    "Código do produto"
                ] || "",

                descricao:
                linha[
                    "Produto"
                ] || "",

                volumes:
                Number(
                    linha[
                        "quantidadeTotal"
                    ]
                ) || 0,

                master:
                linha[
                    "(Palete)Master"
                ] || ""

            };

        }
    );

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

    let skus = 0;

    let volumes = 0;

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

            skus +=
            ptl.length;

            ptl.forEach(item=>{

                volumes +=
                Number(
                    item.volumes
                ) || 0;

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
    ).textContent = skus;

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

            const okLoja =
            item.loja
            .toLowerCase()
            .includes(
                lojaFiltro
            );

            const okPTL =
            item.ptl
            .toLowerCase()
            .includes(
                ptlFiltro
            );

            const okSKU =
            String(
                item.sku
            )
            .toLowerCase()
            .includes(
                skuFiltro
            );

            return (
                okLoja &&
                okPTL &&
                okSKU
            );

        }
    );

    gerarAgrupamento();

}

// ========================================
// RENDERIZAÇÃO
// ========================================

function renderizar(){

    if(
        document
        .activeElement
        ?.id
        ?.includes(
            "filtro"
        )
    ){

        aplicarFiltros();

        return;
    }

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
                        <th>Volumes</th>
                    </tr>
                </thead>

                <tbody>
            `;

            itens.forEach(item=>{

                htmlTabela +=
                `
                <tr>
                    <td>${item.sku}</td>
                    <td>${item.descricao}</td>
                    <td>${item.volumes}</td>
                </tr>
                `;

            });

            htmlTabela +=
            `
                </tbody>
                </table>

                <div class="resumo-ptl">

                    Itens:
                    ${itens.length}

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
