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
