import { initializeApp } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js";
import { getFirestore, doc, setDoc, getDoc } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-firestore.js";

// As chaves que extraí diretamente do seu ecrã do Firebase
const firebaseConfig = {
  apiKey: "AIzaSyDBXVR2RESuvVbx7gQ9PrlJluycJWzGdNw",
  authDomain: "sistema-precos-1b770.firebaseapp.com",
  projectId: "sistema-precos-1b770",
  storageBucket: "sistema-precos-1b770.firebasestorage.app",
  messagingSenderId: "833660695608",
  appId: "1:833660695608:web:e67eecd3a7dfdf4daae2cd"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

document.addEventListener('DOMContentLoaded', async () => {
    const btnToggleAdmin = document.getElementById('btn-toggle-admin');
    const adminArea = document.getElementById('admin-area');
    const btnUpload = document.getElementById('btn-upload');
    const fileInput = document.getElementById('excel-file');
    const statusMsg = document.getElementById('upload-status');
    const errorLog = document.getElementById('error-log');
    const lastUpdateText = document.getElementById('last-update');
    const searchInput = document.getElementById('search-input');
    const btnSearch = document.getElementById('btn-search');
    const btnClear = document.getElementById('btn-clear');
    const resultsList = document.getElementById('results-list');

    let baseDeDados = []; // Os produtos vivem nesta variável agora

    // Variáveis do Carrinho
    let carrinho = JSON.parse(localStorage.getItem('carrinhoCotacao') || '[]');
    if (carrinho.length > 0 && carrinho[0].valorTotalLinha === undefined) {
        carrinho = [];
        localStorage.removeItem('carrinhoCotacao');
    }

    const cartList = document.getElementById('cart-list');
    const cartCount = document.getElementById('cart-count');
    const cartTotal = document.getElementById('cart-total');
    const btnCopyCart = document.getElementById('btn-copy-cart');
    const btnClearCart = document.getElementById('btn-clear-cart');
    const btnExportCsv = document.getElementById('btn-export-csv');

    // Variáveis do Modal
    const modal = document.getElementById('add-modal');
    const btnModalCancel = document.getElementById('modal-cancel');
    const btnModalConfirm = document.getElementById('modal-confirm');
    const inputQty = document.getElementById('modal-qty');
    const inputDiscount = document.getElementById('modal-discount');
    let itemPendente = null;

    // 1. CARREGAR DADOS DO FIREBASE AO ABRIR O SISTEMA
    statusMsg.innerText = "A ligar à nuvem e carregar tabela...";
    try {
        const docSnap = await getDoc(doc(db, "catalogo", "precos"));
        if (docSnap.exists()) {
            const dados = docSnap.data();
            baseDeDados = dados.produtos || [];
            lastUpdateText.innerText = `Última atualização: ${dados.dataAtualizacao}`;
            statusMsg.innerText = "";
        } else {
            statusMsg.innerText = "A tabela na nuvem está vazia. Faça o upload do Excel.";
        }
    } catch (error) {
        console.error(error);
        statusMsg.style.color = "red";
        statusMsg.innerText = "Erro ao ligar ao Firebase. O Firestore está ativo no Modo Teste?";
    }

    renderizarCarrinho();

    // =============== ÁREA DE ADMIN E UPLOAD (FIREBASE) ===============
    btnToggleAdmin.addEventListener('click', () => {
        adminArea.classList.toggle('hidden');
        errorLog.classList.add('hidden');
        statusMsg.innerText = '';
    });

    btnUpload.addEventListener('click', () => {
        const file = fileInput.files[0];
        if (!file) { alert("Por favor, selecione um ficheiro Excel primeiro."); return; }

        const reader = new FileReader();
        reader.onload = async (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            const res = processarDados(jsonData);
            const dataAtual = new Date().toLocaleString('pt-BR');

            statusMsg.style.color = "blue";
            statusMsg.innerText = "A guardar na nuvem. Aguarde...";

            try {
                // Grava na Nuvem (Substitui o localStorage)
                await setDoc(doc(db, "catalogo", "precos"), {
                    produtos: res.produtos,
                    dataAtualizacao: dataAtual
                });

                baseDeDados = res.produtos; // Atualiza memória da página
                lastUpdateText.innerText = `Última atualização: ${dataAtual}`;

                statusMsg.style.color = "green";
                statusMsg.innerText = `Sucesso! ${res.produtos.length} variações guardadas na NUVEM.`;
                
                if (res.erros.length > 0) {
                    errorLog.classList.remove('hidden');
                    errorLog.innerText = "Atenção! As seguintes linhas não foram importadas:\n\n" + res.erros.join('\n');
                } else {
                    errorLog.classList.add('hidden');
                    setTimeout(() => { adminArea.classList.add('hidden'); }, 3000);
                }
            } catch (error) {
                console.error(error);
                statusMsg.style.color = "red";
                statusMsg.innerText = "Erro ao guardar na nuvem.";
            }
        };
        reader.readAsArrayBuffer(file);
    });

    // Função interna de processamento
    function processarDados(rows) {
        let produtos = [];
        let erros = [];
        for (let i = 0; i < rows.length; i++) {
            let col1 = String(rows[i][0] || '').trim();
            let col2 = String(rows[i][1] || '').trim();
            if (!col1 && !col2) continue; 
            const matchCol1 = col1.match(/^([A-Za-z]{2}\s\d{7})[\s-]*(.*)$/);
            if (!matchCol1) {
                if (i === 0) continue; 
                erros.push(`Linha ${i + 1} - Produto mal formatado: "${col1}"`);
                continue;
            }
            let baseCode = matchCol1[1];
            let nomeProduto = matchCol1[2];
            let variacoes = col2.split('|');
            variacoes.forEach(variacao => {
                const matchCol2 = variacao.trim().match(/^([A-Za-z0-9]{2})[^\w]*?(.*?)\s*(R\$.*)$/i);
                if (matchCol2) {
                    let finalCodeNumber = matchCol2[1];
                    let embalagem = matchCol2[2].replace(/\*/g, '').trim(); 
                    let preco = matchCol2[3];
                    let codigoCompleto = `${baseCode}${finalCodeNumber}`; 
                    produtos.push({
                        codigoBase: baseCode, codigo: codigoCompleto,
                        codigoBusca: codigoCompleto.replace(/\s/g, '').toLowerCase(),
                        nome: nomeProduto, embalagem: embalagem, preco: preco
                    });
                } else {
                    if (variacao.trim() !== '') {
                        erros.push(`Linha ${i + 1} - Preço mal formatado: "${variacao.trim()}"`);
                    }
                }
            });
        }
        return { produtos, erros };
    }

    // =============== BUSCA ===============
    searchInput.addEventListener('input', () => {
        if (searchInput.value.trim().length > 0) btnClear.classList.remove('hidden');
        else { btnClear.classList.add('hidden'); resultsList.innerHTML = ''; }
    });

    btnClear.addEventListener('click', () => {
        searchInput.value = ''; btnClear.classList.add('hidden');
        resultsList.innerHTML = ''; searchInput.focus();
    });

    btnSearch.addEventListener('click', realizarBusca);
    searchInput.addEventListener('keypress', (e) => { if (e.key === 'Enter') realizarBusca(); });

    function realizarBusca() {
        const query = searchInput.value.trim().toLowerCase();
        resultsList.innerHTML = '';

        if (baseDeDados.length === 0) { resultsList.innerHTML = '<li><p>Nenhum dado carregado na nuvem.</p></li>'; return; }
        if (query === '') return;

        const queryLimpa = query.replace(/\s/g, '');
        const resultados = baseDeDados.filter(prod => {
            return prod.codigoBusca.includes(queryLimpa) || 
                   prod.nome.toLowerCase().includes(query) ||
                   prod.codigoBase.toLowerCase().replace(/\s/g, '').includes(queryLimpa);
        });

        if (resultados.length === 0) { resultsList.innerHTML = '<li><p>Nenhum produto encontrado.</p></li>'; return; }

        const produtosAgrupados = {};
        resultados.forEach(prod => {
            if (!produtosAgrupados[prod.codigoBase]) {
                produtosAgrupados[prod.codigoBase] = { nome: prod.nome, codigoBase: prod.codigoBase, variacoes: [] };
            }
            produtosAgrupados[prod.codigoBase].variacoes.push(prod);
        });

        Object.values(produtosAgrupados).forEach(grupo => {
            const li = document.createElement('li');
            li.className = 'result-item';
            let htmlCard = `<h3>${grupo.nome}</h3><span class="base-code">Linha Base: ${grupo.codigoBase}</span><ul class="variant-list">`;
            grupo.variacoes.forEach(v => {
                htmlCard += `
                    <li class="variant-item">
                        <div class="variant-info">
                            <strong>${v.embalagem}</strong><br>
                            <span class="code-tag">${v.codigo}</span>
                        </div>
                        <div class="variant-actions">
                            <span class="price-tag">${v.preco}</span>
                            <button class="btn-add" title="Adicionar à cotação" 
                                data-codigo="${v.codigo}" 
                                data-nome="${grupo.nome}" 
                                data-embalagem="${v.embalagem}" 
                                data-preco="${v.preco}">+</button>
                        </div>
                    </li>
                `;
            });
            htmlCard += `</ul>`;
            li.innerHTML = htmlCard;
            resultsList.appendChild(li);
        });
    }

    // =============== MODAL DE ADIÇÃO ===============
    resultsList.addEventListener('click', (e) => {
        if (e.target.classList.contains('btn-add')) {
            const btn = e.target;
            itemPendente = {
                codigo: btn.dataset.codigo,
                nome: btn.dataset.nome,
                embalagem: btn.dataset.embalagem,
                precoOriginalString: btn.dataset.preco
            };
            document.getElementById('modal-product-name').innerText = `[${itemPendente.codigo}] ${itemPendente.nome} (${itemPendente.embalagem})`;
            inputQty.value = 1;
            inputDiscount.value = 0;
            modal.classList.remove('hidden');
            inputQty.focus();
        }
    });

    btnModalCancel.addEventListener('click', () => { modal.classList.add('hidden'); itemPendente = null; });

    btnModalConfirm.addEventListener('click', () => {
        if (!itemPendente) return;

        const indexExistente = carrinho.findIndex(item => item.codigo === itemPendente.codigo);
        if (indexExistente !== -1) {
            alert('Este item já está no carrinho! Você pode alterar a quantidade diretamente por lá.');
            return; 
        }

        const qtd = parseInt(inputQty.value) || 1;
        let desc = parseFloat(inputDiscount.value) || 0;
        if (desc < 0) desc = 0;
        if (desc > 100) desc = 100;

        let valorBaseNum = parseFloat(itemPendente.precoOriginalString.replace('R$', '').replace(/\./g, '').replace(',', '.').trim());
        let valorComDesconto = valorBaseNum * (1 - (desc / 100));
        let valorTotalLinha = valorComDesconto * qtd;

        carrinho.push({
            codigo: itemPendente.codigo, nome: itemPendente.nome, embalagem: itemPendente.embalagem,
            quantidade: qtd, descontoAplicado: desc, valorUnitarioFinal: valorComDesconto, valorTotalLinha: valorTotalLinha
        });

        localStorage.setItem('carrinhoCotacao', JSON.stringify(carrinho));
        renderizarCarrinho();
        modal.classList.add('hidden');
        itemPendente = null;
    });

    // =============== CARRINHO E EXPORTAÇÕES ===============
    cartList.addEventListener('click', (e) => {
        const index = e.target.dataset.index;
        if (e.target.classList.contains('btn-remove')) { carrinho.splice(index, 1); } 
        else if (e.target.classList.contains('btn-qty-plus')) {
            carrinho[index].quantidade++;
            carrinho[index].valorTotalLinha = carrinho[index].quantidade * carrinho[index].valorUnitarioFinal;
        } 
        else if (e.target.classList.contains('btn-qty-minus')) {
            if (carrinho[index].quantidade > 1) {
                carrinho[index].quantidade--;
                carrinho[index].valorTotalLinha = carrinho[index].quantidade * carrinho[index].valorUnitarioFinal;
            }
        } else { return; }

        localStorage.setItem('carrinhoCotacao', JSON.stringify(carrinho));
        renderizarCarrinho();
    });

    function renderizarCarrinho() {
        cartList.innerHTML = '';
        let totalGeral = 0;

        if (carrinho.length === 0) {
            cartList.innerHTML = '<li><p class="empty-cart">Seu carrinho está vazio.</p></li>';
            cartCount.innerText = '(0)';
            cartTotal.innerText = 'R$ 0,00';
            return;
        }

        cartCount.innerText = `(${carrinho.length})`;

        carrinho.forEach((item, index) => {
            const li = document.createElement('li');
            li.className = 'cart-item';
            let htmlDesconto = item.descontoAplicado > 0 ? `<span class="cart-discount">(-${item.descontoAplicado}%)</span>` : '';
            let unitarioFormatado = item.valorUnitarioFinal.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
            let totalLinhaFormatado = item.valorTotalLinha.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });

            li.innerHTML = `
                <div style="flex: 1;">
                    <span class="cart-item-title">${item.nome}</span>
                    <span class="cart-item-code">${item.codigo} | ${item.embalagem}</span>
                    <div style="display: flex; align-items: center; margin-top: 4px;">
                        <div class="cart-qty-controls">
                            <button class="btn-qty-minus" data-index="${index}">-</button>
                            <span>${item.quantidade}x</span>
                            <button class="btn-qty-plus" data-index="${index}">+</button>
                        </div>
                        ${unitarioFormatado} ${htmlDesconto}
                    </div>
                    <span class="cart-price-total">${totalLinhaFormatado}</span>
                </div>
                <button class="btn-remove" data-index="${index}" title="Remover item">X</button>
            `;
            cartList.appendChild(li);
            totalGeral += item.valorTotalLinha;
        });

        cartTotal.innerText = totalGeral.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
    }

    btnClearCart.addEventListener('click', () => {
        if (confirm('Tem certeza que deseja esvaziar a cotação?')) {
            carrinho = [];
            localStorage.setItem('carrinhoCotacao', JSON.stringify(carrinho));
            renderizarCarrinho();
        }
    });

    btnCopyCart.addEventListener('click', () => {
        if (carrinho.length === 0) { alert('Adicione produtos antes de copiar!'); return; }
        let textoCopia = "🛒 LISTA DE COTAÇÃO:\n\n";
        carrinho.forEach(item => {
            let unitFormatado = item.valorUnitarioFinal.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
            let totalFormatado = item.valorTotalLinha.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
            textoCopia += `[${item.codigo}] ${item.nome} (${item.embalagem})\n`;
            textoCopia += `Qtd: ${item.quantidade} | Unit: ${unitFormatado}\n`;
            textoCopia += `Total do Item: ${totalFormatado}\n\n`;
        });
        textoCopia += `TOTAL ESTIMADO: ${cartTotal.innerText}`;

        navigator.clipboard.writeText(textoCopia).then(() => {
            const textoOriginal = btnCopyCart.innerText;
            btnCopyCart.innerText = 'Copiado com Sucesso! ✅';
            btnCopyCart.classList.add('btn-success-feedback');
            setTimeout(() => {
                btnCopyCart.innerText = textoOriginal;
                btnCopyCart.classList.remove('btn-success-feedback');
            }, 2000);
        }).catch(() => { alert('Selecione o texto e copie manualmente.'); });
    });

    btnExportCsv.addEventListener('click', () => {
        if (carrinho.length === 0) { alert('Adicione produtos antes de exportar!'); return; }
        let csvContent = "data:text/csv;charset=utf-8,\uFEFF"; 
        csvContent += "Código;Nome;Embalagem;Quantidade;Valor Unitário (R$);Valor Total (R$)\n";
        carrinho.forEach(item => {
            let nomeStr = `"${item.nome}"`; 
            let unitarioStr = item.valorUnitarioFinal.toFixed(2).replace('.', ',');
            let totalStr = item.valorTotalLinha.toFixed(2).replace('.', ',');
            csvContent += `${item.codigo};${nomeStr};${item.embalagem};${item.quantidade};${unitarioStr};${totalStr}\n`;
        });
        const encodedUri = encodeURI(csvContent);
        const link = document.createElement("a");
        link.setAttribute("href", encodedUri);
        link.setAttribute("download", `cotacao_${new Date().getTime()}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });
});