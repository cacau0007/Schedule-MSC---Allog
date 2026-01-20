// ============================================
// SERVIDOR DE SCHEDULES - MSC
// Versão corrigida - Excel funcional + Extração individual
// ============================================

const express = require('express');
const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

puppeteer.use(StealthPlugin());

const app = express();
const PORT = process.env.PORT || 3000;
const EXPORTS_DIR = path.join(__dirname, '../exports');

if (!fs.existsSync(EXPORTS_DIR)) fs.mkdirSync(EXPORTS_DIR, { recursive: true });

app.use(express.json());
app.use('/exports', express.static(EXPORTS_DIR));
app.use(express.static(path.join(__dirname, '../frontend')));

// ============================================
// SERVIÇOS MSC 2026
// ============================================
const ALL_SERVICES = ['Carioca', 'Ipanema', 'Santana', 'Jade'];

// POLs de CONEXÃO (via Singapore/Busan) - não filtrar
const CONNECTION_POLS = [
    'Jakarta', 'Surabaya', 'Panjang', 'Belawan', 'Semarang',
    'Laem Chabang', 'Bangkok', 'Haiphong', 'Ho Chi Minh', 'Phnom Penh',
    'Port Klang', 'Penang', 'Tanjung Pelepas',
    'Xingang', 'Tianjin', 'Dalian', 'Incheon',
    'Yokohama', 'Tokyo', 'Kobe', 'Osaka', 'Nagoya',
    'Kaohsiung', 'Keelung'
];

function isConnectionRoute(pol) {
    return CONNECTION_POLS.includes(pol);
}

// ============================================
// MAPEAMENTO DE SERVIÇOS POR ROTA
// ============================================
const SERVICE_ROUTES = {
    // SANTOS - Todos os 4 serviços
    'Shanghai-Santos': ['Carioca', 'Ipanema', 'Santana', 'Jade'],
    'Ningbo-Santos': ['Carioca', 'Ipanema', 'Santana', 'Jade'],
    'Shekou-Santos': ['Carioca', 'Ipanema', 'Santana', 'Jade'],
    'Busan-Santos': ['Carioca', 'Ipanema', 'Santana', 'Jade'],
    'Singapore-Santos': ['Carioca', 'Ipanema', 'Santana', 'Jade'],
    'Qingdao-Santos': ['Carioca', 'Santana', 'Jade'],
    'Yantian-Santos': ['Ipanema', 'Jade'],
    'Hong Kong-Santos': ['Ipanema', 'Jade'],
    'Xiamen-Santos': ['Jade'],
    'Nansha-Santos': ['Jade'],
    'Fuzhou-Santos': ['Jade'],

    // RIO DE JANEIRO - Carioca, Santana, Jade
    'Shanghai-Rio de Janeiro': ['Carioca', 'Santana', 'Jade'],
    'Ningbo-Rio de Janeiro': ['Carioca', 'Santana', 'Jade'],
    'Shekou-Rio de Janeiro': ['Carioca', 'Santana', 'Jade'],
    'Busan-Rio de Janeiro': ['Carioca', 'Santana', 'Jade'],
    'Singapore-Rio de Janeiro': ['Carioca', 'Santana', 'Jade'],
    'Qingdao-Rio de Janeiro': ['Carioca', 'Santana', 'Jade'],
    'Yantian-Rio de Janeiro': ['Jade'],
    'Hong Kong-Rio de Janeiro': ['Jade'],
    'Xiamen-Rio de Janeiro': ['Jade'],
    'Nansha-Rio de Janeiro': ['Jade'],
    'Fuzhou-Rio de Janeiro': ['Jade'],

    // PARANAGUÁ - Todos os 4 serviços
    'Shanghai-Paranagua': ['Carioca', 'Ipanema', 'Santana', 'Jade'],
    'Ningbo-Paranagua': ['Carioca', 'Ipanema', 'Santana', 'Jade'],
    'Shekou-Paranagua': ['Carioca', 'Ipanema', 'Santana', 'Jade'],
    'Busan-Paranagua': ['Carioca', 'Ipanema', 'Santana', 'Jade'],
    'Singapore-Paranagua': ['Carioca', 'Ipanema', 'Santana', 'Jade'],
    'Qingdao-Paranagua': ['Carioca', 'Santana', 'Jade'],
    'Yantian-Paranagua': ['Ipanema', 'Jade'],
    'Hong Kong-Paranagua': ['Ipanema', 'Jade'],
    'Xiamen-Paranagua': ['Jade'],
    'Nansha-Paranagua': ['Jade'],
    'Fuzhou-Paranagua': ['Jade'],

    // NAVEGANTES - Ipanema, Santana, Jade (SEM Carioca)
    'Shanghai-Navegantes': ['Ipanema', 'Santana', 'Jade'],
    'Ningbo-Navegantes': ['Ipanema', 'Santana', 'Jade'],
    'Shekou-Navegantes': ['Ipanema', 'Santana', 'Jade'],
    'Busan-Navegantes': ['Ipanema', 'Santana', 'Jade'],
    'Singapore-Navegantes': ['Ipanema', 'Santana', 'Jade'],
    'Qingdao-Navegantes': ['Santana', 'Jade'],
    'Yantian-Navegantes': ['Ipanema', 'Jade'],
    'Hong Kong-Navegantes': ['Ipanema', 'Jade'],
    'Xiamen-Navegantes': ['Jade'],
    'Nansha-Navegantes': ['Jade'],
    'Fuzhou-Navegantes': ['Jade'],

    // ITAPOÁ - Carioca, Jade
    'Shanghai-Itapoa': ['Carioca', 'Jade'],
    'Ningbo-Itapoa': ['Carioca', 'Jade'],
    'Shekou-Itapoa': ['Carioca', 'Jade'],
    'Busan-Itapoa': ['Carioca', 'Jade'],
    'Singapore-Itapoa': ['Carioca', 'Jade'],
    'Qingdao-Itapoa': ['Carioca', 'Jade'],
    'Yantian-Itapoa': ['Jade'],
    'Hong Kong-Itapoa': ['Jade'],
    'Xiamen-Itapoa': ['Jade'],
    'Nansha-Itapoa': ['Jade'],
    'Fuzhou-Itapoa': ['Jade'],

    // ITAGUAÍ - Carioca, Santana
    'Shanghai-Itaguai': ['Carioca', 'Santana'],
    'Ningbo-Itaguai': ['Carioca', 'Santana'],
    'Shekou-Itaguai': ['Carioca', 'Santana'],
    'Busan-Itaguai': ['Carioca', 'Santana'],
    'Singapore-Itaguai': ['Carioca', 'Santana'],
    'Qingdao-Itaguai': ['Carioca', 'Santana'],

    // IMBITUBA - Santana
    'Shanghai-Imbituba': ['Santana'],
    'Ningbo-Imbituba': ['Santana'],
    'Shekou-Imbituba': ['Santana'],
    'Busan-Imbituba': ['Santana'],
    'Singapore-Imbituba': ['Santana'],
    'Qingdao-Imbituba': ['Santana'],

    // ITAJAÍ - Santana
    'Shanghai-Itajai': ['Santana'],
    'Ningbo-Itajai': ['Santana'],
    'Shekou-Itajai': ['Santana'],
    'Busan-Itajai': ['Santana'],
    'Singapore-Itajai': ['Santana'],
    'Qingdao-Itajai': ['Santana'],

    // SUAPE - Santana
    'Shanghai-Suape': ['Santana'],
    'Ningbo-Suape': ['Santana'],
    'Shekou-Suape': ['Santana'],
    'Busan-Suape': ['Santana'],
    'Singapore-Suape': ['Santana'],
    'Qingdao-Suape': ['Santana'],

    // SALVADOR - Santana
    'Shanghai-Salvador': ['Santana'],
    'Ningbo-Salvador': ['Santana'],
    'Shekou-Salvador': ['Santana'],
    'Busan-Salvador': ['Santana'],
    'Singapore-Salvador': ['Santana'],
    'Qingdao-Salvador': ['Santana'],

    // MONTEVIDEO - Ipanema
    'Shanghai-Montevideo': ['Ipanema'],
    'Ningbo-Montevideo': ['Ipanema'],
    'Shekou-Montevideo': ['Ipanema'],
    'Busan-Montevideo': ['Ipanema'],
    'Singapore-Montevideo': ['Ipanema'],
    'Yantian-Montevideo': ['Ipanema'],
    'Hong Kong-Montevideo': ['Ipanema'],

    // BUENOS AIRES - Ipanema
    'Shanghai-Buenos Aires': ['Ipanema'],
    'Ningbo-Buenos Aires': ['Ipanema'],
    'Shekou-Buenos Aires': ['Ipanema'],
    'Busan-Buenos Aires': ['Ipanema'],
    'Singapore-Buenos Aires': ['Ipanema'],
    'Yantian-Buenos Aires': ['Ipanema'],
    'Hong Kong-Buenos Aires': ['Ipanema'],

    // RIO GRANDE - Ipanema
    'Shanghai-Rio Grande': ['Ipanema'],
    'Ningbo-Rio Grande': ['Ipanema'],
    'Shekou-Rio Grande': ['Ipanema'],
    'Busan-Rio Grande': ['Ipanema'],
    'Singapore-Rio Grande': ['Ipanema'],
    'Yantian-Rio Grande': ['Ipanema'],
    'Hong Kong-Rio Grande': ['Ipanema'],

    // MANAUS - Santana
    'Shanghai-Manaus': ['Santana'],
    'Ningbo-Manaus': ['Santana'],
    'Qingdao-Manaus': ['Santana'],
    'Busan-Manaus': ['Santana'],
    'Shekou-Manaus': ['Santana'],
    'Singapore-Manaus': ['Santana'],
    'Yantian-Manaus': ['Santana'],

    // VITÓRIA - Santana, Carioca
    'Shanghai-Vitoria': ['Santana', 'Carioca'],
    'Ningbo-Vitoria': ['Santana', 'Carioca'],
    'Qingdao-Vitoria': ['Santana', 'Carioca'],
    'Busan-Vitoria': ['Santana', 'Carioca'],
    'Shekou-Vitoria': ['Santana', 'Carioca'],
    'Singapore-Vitoria': ['Santana', 'Carioca'],

    // PECÉM - Santana
    'Shanghai-Pecem': ['Santana'],
    'Ningbo-Pecem': ['Santana'],
    'Qingdao-Pecem': ['Santana'],
    'Busan-Pecem': ['Santana'],
    'Shekou-Pecem': ['Santana'],
    'Singapore-Pecem': ['Santana'],

    // FORTALEZA - Santana
    'Shanghai-Fortaleza': ['Santana'],
    'Ningbo-Fortaleza': ['Santana'],
    'Qingdao-Fortaleza': ['Santana'],
    'Busan-Fortaleza': ['Santana'],
    'Shekou-Fortaleza': ['Santana'],
    'Singapore-Fortaleza': ['Santana'],

    // BELÉM - Santana
    'Shanghai-Belem': ['Santana'],
    'Ningbo-Belem': ['Santana'],
    'Qingdao-Belem': ['Santana'],
    'Busan-Belem': ['Santana'],
    'Shekou-Belem': ['Santana'],
    'Singapore-Belem': ['Santana']
};

function getAvailableServices(pol, pod) {
    if (isConnectionRoute(pol)) return null;
    return SERVICE_ROUTES[`${pol}-${pod}`] || null;
}

// ============================================
// ENDPOINT: Serviços disponíveis
// ============================================
app.get('/api/available-services', (req, res) => {
    const { pol, pod } = req.query;
    
    if (!pol || !pod) {
        return res.json({ services: ALL_SERVICES, message: 'Selecione POL e POD' });
    }
    
    if (isConnectionRoute(pol)) {
        return res.json({ 
            services: ALL_SERVICES,
            mapped: false,
            isConnection: true,
            message: `Rota de conexão (${pol})`
        });
    }
    
    const services = getAvailableServices(pol, pod);
    const isMapped = SERVICE_ROUTES.hasOwnProperty(`${pol}-${pod}`);
    
    if (!isMapped) {
        return res.json({ 
            services: null,
            mapped: false,
            isConnection: true,
            message: `Rota não mapeada`
        });
    }
    
    return res.json({ 
        services,
        mapped: true,
        isConnection: false,
        message: `Serviços: ${services.join(', ')}`
    });
});

// ============================================
// ENDPOINT: Buscar schedules
// ============================================
app.post('/api/schedules', async (req, res) => {
    const { pol, pod, carriers, service } = req.body;
    
    console.log(`\n=== BUSCA: ${pol} → ${pod} | Serviço: ${service || 'ALL'} ===`);
    
    const results = [];
    
    try {
        if (carriers.includes('MSC')) {
            const mscData = await scrapeMSC(pol, pod, service);
            results.push(...mscData);
        }
        
        // Gerar Excel
        const filename = `Schedules_${pol.replace(/ /g, '_')}_${pod.replace(/ /g, '_')}_${new Date().toISOString().slice(0, 10)}.xlsx`;
        await generateExcel(results, pol, pod, filename);
        
        res.json({
            success: true,
            count: results.length,
            file: `/exports/${filename}`,
            data: results
        });
    } catch (error) {
        console.error('❌ Erro:', error);
        res.status(500).json({ success: false, error: error.message });
    }
});

// ============================================
// SCRAPER MSC
// ============================================
async function scrapeMSC(pol, pod, service = null) {
    const sailings = [];
    let browser = null;
    let page = null;
    
    const isConnection = isConnectionRoute(pol);
    console.log(`🚢 MSC: ${pol} → ${pod} | Conexão: ${isConnection}`);
    
    try {
        browser = await puppeteer.launch({
            headless: 'new',
            executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || '/usr/bin/google-chrome-stable',
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--disable-web-security',
                '--disable-blink-features=AutomationControlled',
                '--window-size=1920,1080'
            ]
        });
        
        page = await browser.newPage();
        await page.setViewport({ width: 1920, height: 1080 });
        await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36');
        
        // Acessar MSC
        console.log('1. Acessando MSC...');
        await page.goto('https://www.msc.com/en/search-a-schedule', { waitUntil: 'networkidle2', timeout: 60000 });
        await new Promise(r => setTimeout(r, 3000));
        
        // Fechar cookies
        try {
            const cookieBtn = await page.$('#onetrust-accept-btn-handler');
            if (cookieBtn) await cookieBtn.click();
            await new Promise(r => setTimeout(r, 1000));
        } catch (e) {}
        
        // POL
        console.log('2. Selecionando POL...');
        await page.click('#placeOfLoadingInput');
        await new Promise(r => setTimeout(r, 500));
        await page.type('#placeOfLoadingInput', pol, { delay: 50 });
        await new Promise(r => setTimeout(r, 2000));
        await page.keyboard.press('ArrowDown');
        await page.keyboard.press('Enter');
        await new Promise(r => setTimeout(r, 1000));
        
        // POD
        console.log('3. Selecionando POD...');
        await page.click('#placeOfDischargeInput');
        await new Promise(r => setTimeout(r, 500));
        await page.type('#placeOfDischargeInput', pod, { delay: 50 });
        await new Promise(r => setTimeout(r, 2000));
        await page.keyboard.press('ArrowDown');
        await page.keyboard.press('Enter');
        await new Promise(r => setTimeout(r, 1000));
        
        // Buscar
        console.log('4. Clicando em Search...');
        const searchBtn = await page.$('button[type="submit"], button:has-text("Search"), .search-button, [class*="search"]');
        if (searchBtn) {
            await searchBtn.click();
        } else {
            await page.evaluate(() => {
                const btns = Array.from(document.querySelectorAll('button'));
                const searchBtn = btns.find(b => b.textContent.toLowerCase().includes('search'));
                if (searchBtn) searchBtn.click();
            });
        }
        
        await new Promise(r => setTimeout(r, 8000));
        await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-01-resultados.png') });
        
        // Extrair resultados
        console.log('5. Extraindo resultados...');
        const results = await page.evaluate(() => {
            const items = [];
            const text = document.body.innerText;
            const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
            
            // Padrões de data
            const datePattern = /(\w{3}\s+\d{1,2}(?:st|nd|rd|th)?\s+\w{3}\s+\d{4})/gi;
            const transitPattern = /(\d+)\s*days?/gi;
            
            // Procurar por linhas que parecem resultados
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i];
                
                // Se linha contém "MSC " seguido de nome de navio
                if (line.match(/^MSC\s+[A-Z]/i)) {
                    const vessel = line;
                    let etd = '-', eta = '-', transit = '-', routeType = 'Transbordo';
                    
                    // Procurar datas nas próximas linhas
                    for (let j = i + 1; j <= Math.min(i + 10, lines.length - 1); j++) {
                        const nextLine = lines[j];
                        const dates = nextLine.match(datePattern);
                        if (dates && dates.length >= 1 && etd === '-') {
                            etd = dates[0];
                            if (dates.length >= 2) eta = dates[1];
                        }
                        const transitMatch = nextLine.match(/(\d+)\s*days?/i);
                        if (transitMatch) {
                            transit = transitMatch[1] + ' dias';
                        }
                        if (nextLine.toLowerCase().includes('direct')) {
                            routeType = 'Direto';
                        }
                    }
                    
                    if (etd !== '-') {
                        items.push({ vessel, etd, eta, transit, routeType, service: '-', transbordo: '-', transbordoDate: '' });
                    }
                }
            }
            
            return items;
        });
        
        console.log(`   Encontrados: ${results.length} navios`);
        
        // Para rotas de conexão: extrair detalhes clicando em cada navio
        if (isConnection && results.length > 0) {
            console.log('6. Extraindo detalhes de cada navio...');
            
            for (let i = 0; i < Math.min(results.length, 15); i++) {
                const item = results[i];
                console.log(`   [${i+1}/${results.length}] ${item.vessel}`);
                
                try {
                    // Encontrar e clicar no card do navio
                    const extracted = await page.evaluate(async (vesselName, index) => {
                        // Encontrar todos os elementos com o nome do navio
                        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
                        const nodes = [];
                        let node;
                        while (node = walker.nextNode()) {
                            if (node.textContent.includes(vesselName)) {
                                nodes.push(node);
                            }
                        }
                        
                        if (nodes.length === 0) return { success: false, error: 'Não encontrado' };
                        
                        // Pegar o nó correto baseado no index (evitar repetição)
                        const targetNode = nodes[Math.min(index, nodes.length - 1)];
                        let card = targetNode.parentElement;
                        
                        // Subir na árvore DOM até achar um card clicável
                        for (let i = 0; i < 10; i++) {
                            if (!card.parentElement) break;
                            card = card.parentElement;
                            const rect = card.getBoundingClientRect();
                            if (rect.width > 500 && rect.height > 50 && rect.height < 300) break;
                        }
                        
                        // Scroll e clique
                        card.scrollIntoView({ block: 'center' });
                        await new Promise(r => setTimeout(r, 300));
                        card.click();
                        
                        // Esperar expansão
                        await new Promise(r => setTimeout(r, 2000));
                        
                        // Extrair dados do modal/expansão
                        const fullText = document.body.innerText;
                        const lines = fullText.split('\n').map(l => l.trim()).filter(l => l);
                        
                        let service = '-';
                        let transbordo = '-';
                        let transbordoDate = '';
                        
                        const services = ['Santana', 'Carioca', 'Ipanema', 'Jade'];
                        
                        for (let i = 0; i < lines.length; i++) {
                            const line = lines[i];
                            const lower = line.toLowerCase();
                            
                            // Detectar serviço
                            for (const svc of services) {
                                if (line.includes(svc + ' Service') || line === svc || 
                                    lower.includes(svc.toLowerCase() + ' service')) {
                                    service = svc;
                                    break;
                                }
                            }
                            
                            // Detectar transbordo
                            if (lower.includes('singapore') || line.includes('SGSIN')) {
                                transbordo = 'SIN';
                            } else if (lower.includes('busan') || line.includes('KRPUS')) {
                                transbordo = 'BUS';
                            } else if (lower.includes('tanjung pelepas')) {
                                transbordo = 'TPP';
                            } else if (lower.includes('port klang')) {
                                transbordo = 'PKG';
                            }
                            
                            // Se achou transbordo, procurar data
                            if (transbordo !== '-' && !transbordoDate) {
                                for (let j = Math.max(0, i-3); j <= Math.min(lines.length-1, i+3); j++) {
                                    const dateMatch = lines[j].match(/(\d{1,2}(?:st|nd|rd|th)?\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4})/i);
                                    if (dateMatch) {
                                        transbordoDate = dateMatch[1];
                                        break;
                                    }
                                }
                            }
                        }
                        
                        // Fechar modal
                        document.body.click();
                        
                        return { success: true, service, transbordo, transbordoDate };
                    }, item.vessel, i);
                    
                    if (extracted.success) {
                        item.service = extracted.service;
                        item.transbordo = extracted.transbordo;
                        item.transbordoDate = extracted.transbordoDate;
                        console.log(`      ✓ ${extracted.service} | ${extracted.transbordo} ${extracted.transbordoDate}`);
                    }
                    
                    await page.keyboard.press('Escape');
                    await new Promise(r => setTimeout(r, 500));
                    
                } catch (e) {
                    console.log(`      ✗ Erro: ${e.message}`);
                }
            }
        }
        
        // Adicionar aos resultados
        results.forEach(r => {
            sailings.push({
                carrier: 'MSC',
                service: r.service || '-',
                vessel: r.vessel,
                etd: r.etd,
                eta: r.eta,
                transit: r.transit,
                routeType: r.routeType,
                transbordo: r.transbordo || '-',
                transbordoDate: r.transbordoDate || ''
            });
        });
        
    } catch (error) {
        console.log(`❌ Erro MSC: ${error.message}`);
        if (page) await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-error.png') }).catch(() => {});
    } finally {
        if (browser) await browser.close();
    }
    
    console.log(`✅ Total: ${sailings.length} schedules`);
    return sailings;
}

// ============================================
// GERAÇÃO DE EXCEL - Usando CSV como fallback seguro
// ============================================
async function generateExcel(sailings, pol, pod, filename) {
    const filepath = path.join(EXPORTS_DIR, filename);
    
    try {
        // Método 1: ExcelJS com writeBuffer
        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'ALLOG';
        workbook.created = new Date();
        
        const sheet = workbook.addWorksheet('Schedules');
        
        // Definir colunas
        sheet.columns = [
            { header: 'CARRIER', key: 'carrier', width: 10 },
            { header: 'SERVICO', key: 'service', width: 12 },
            { header: 'NAVIO', key: 'vessel', width: 25 },
            { header: 'ETD', key: 'etd', width: 20 },
            { header: 'TRANSBORDO', key: 'transbordo', width: 25 },
            { header: 'ETA', key: 'eta', width: 20 },
            { header: 'TRANSIT', key: 'transit', width: 12 },
            { header: 'TIPO', key: 'tipo', width: 14 }
        ];
        
        // Estilo do header
        const headerRow = sheet.getRow(1);
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D32' } };
        headerRow.alignment = { horizontal: 'center' };
        
        // Adicionar dados
        sailings.forEach(s => {
            const transbordoText = s.transbordo && s.transbordo !== '-' 
                ? `${s.transbordo}${s.transbordoDate ? ' (' + s.transbordoDate + ')' : ''}`
                : '-';
            
            sheet.addRow({
                carrier: s.carrier || 'MSC',
                service: s.service || '-',
                vessel: s.vessel || '-',
                etd: s.etd || '-',
                transbordo: transbordoText,
                eta: s.eta || '-',
                transit: s.transit || '-',
                tipo: s.routeType || '-'
            });
        });
        
        // Salvar com buffer
        const buffer = await workbook.xlsx.writeBuffer();
        fs.writeFileSync(filepath, Buffer.from(buffer));
        
        console.log(`📊 Excel: ${filename} (${buffer.length} bytes)`);
        
    } catch (xlsxError) {
        console.error('⚠️ Erro ExcelJS, gerando CSV...', xlsxError.message);
        
        // Fallback: CSV
        const csvFilename = filename.replace('.xlsx', '.csv');
        const csvPath = path.join(EXPORTS_DIR, csvFilename);
        
        let csv = 'CARRIER,SERVICO,NAVIO,ETD,TRANSBORDO,ETA,TRANSIT,TIPO\n';
        sailings.forEach(s => {
            const tr = s.transbordo !== '-' ? `${s.transbordo} (${s.transbordoDate})` : '-';
            csv += `"${s.carrier}","${s.service}","${s.vessel}","${s.etd}","${tr}","${s.eta}","${s.transit}","${s.routeType}"\n`;
        });
        
        fs.writeFileSync(csvPath, csv, 'utf8');
        console.log(`📊 CSV: ${csvFilename}`);
    }
}

// ============================================
// SERVIDOR
// ============================================
app.listen(PORT, () => {
    console.log(`\n🚀 Servidor: http://localhost:${PORT}`);
    console.log(`📁 Exports: ${EXPORTS_DIR}\n`);
});
