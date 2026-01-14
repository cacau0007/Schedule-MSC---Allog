const express = require('express');
const cors = require('cors');
const puppeteer = require('puppeteer-core');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../frontend')));

const EXPORTS_DIR = path.join(__dirname, '../exports');
if (!fs.existsSync(EXPORTS_DIR)) fs.mkdirSync(EXPORTS_DIR, { recursive: true });

// ============================================
// PORTOS
// ============================================
const PORTS = {
    // China
    'NINGBO': { name: 'Ningbo', search: 'Ningbo' },
    'SHANGHAI': { name: 'Shanghai', search: 'Shanghai' },
    'QINGDAO': { name: 'Qingdao', search: 'Qingdao' },
    'YANTIAN': { name: 'Yantian', search: 'Yantian' },
    'XIAMEN': { name: 'Xiamen', search: 'Xiamen' },
    'SHEKOU': { name: 'Shekou', search: 'Shekou' },
    'SHENZHEN': { name: 'Shenzhen', search: 'Shenzhen' },
    'TIANJIN': { name: 'Tianjin', search: 'Tianjin' },
    'DALIAN': { name: 'Dalian', search: 'Dalian' },
    'GUANGZHOU': { name: 'Guangzhou', search: 'Guangzhou' },
    'HONG KONG': { name: 'Hong Kong', search: 'Hong Kong' },
    
    // Sudeste Asiático
    'SINGAPORE': { name: 'Singapore', search: 'Singapore' },
    'JAKARTA': { name: 'Jakarta', search: 'Jakarta' },
    'PORT KLANG': { name: 'Port Klang', search: 'Port Klang' },
    'HO CHI MINH': { name: 'Ho Chi Minh', search: 'Ho Chi Minh' },
    'BANGKOK': { name: 'Bangkok', search: 'Bangkok' },
    'LAEM CHABANG': { name: 'Laem Chabang', search: 'Laem Chabang' },
    'HAIPHONG': { name: 'Haiphong', search: 'Haiphong' },
    'MANILA': { name: 'Manila', search: 'Manila' },
    'SURABAYA': { name: 'Surabaya', search: 'Surabaya' },
    
    // Ásia Oriental
    'BUSAN': { name: 'Busan', search: 'Busan' },
    'KAOHSIUNG': { name: 'Kaohsiung', search: 'Kaohsiung' },
    'TOKYO': { name: 'Tokyo', search: 'Tokyo' },
    'YOKOHAMA': { name: 'Yokohama', search: 'Yokohama' },
    
    // Índia
    'NHAVA SHEVA': { name: 'Nhava Sheva', search: 'Nhava Sheva' },
    'MUNDRA': { name: 'Mundra', search: 'Mundra' },
    'CHENNAI': { name: 'Chennai', search: 'Chennai' },
    
    // Brasil - Sul
    'NAVEGANTES': { name: 'Navegantes', search: 'Navegantes' },
    'ITAJAI': { name: 'Itajai', search: 'Itajai' },
    'ITAPOA': { name: 'Itapoa', search: 'Itapoa' },
    'PARANAGUA': { name: 'Paranagua', search: 'Paranagua' },
    'RIO GRANDE': { name: 'Rio Grande', search: 'Rio Grande' },
    
    // Brasil - Sudeste
    'SANTOS': { name: 'Santos', search: 'Santos' },
    'RIO DE JANEIRO': { name: 'Rio de Janeiro', search: 'Rio de Janeiro' },
    'VITORIA': { name: 'Vitoria', search: 'Vitoria' },
    'SEPETIBA': { name: 'Sepetiba', search: 'Sepetiba' },
    
    // Brasil - Nordeste
    'SALVADOR': { name: 'Salvador', search: 'Salvador' },
    'SUAPE': { name: 'Suape', search: 'Suape' },
    'PECEM': { name: 'Pecem', search: 'Pecem' },
    'FORTALEZA': { name: 'Fortaleza', search: 'Fortaleza' },
    
    // Brasil - Norte
    'MANAUS': { name: 'Manaus', search: 'Manaus' },
    'BELEM': { name: 'Belem', search: 'Belem' },
    'VILA DO CONDE': { name: 'Vila do Conde', search: 'Vila do Conde' },
};

// ============================================
// SCRAPER MSC
// ============================================
async function scrapeMSC(pol, pod, service) {
    console.log('\n' + '='.repeat(50));
    console.log(`🟡 MSC - ${pol} → ${pod}`);
    if (service) console.log(`   Serviço: ${service}`);
    console.log('='.repeat(50));
    
    const browser = await puppeteer.launch({
        executablePath: process.env.CHROME_PATH || '/usr/bin/chromium-browser',
        headless: true, // INVISÍVEL - não abre janela
        args: [
            '--no-sandbox', 
            '--disable-setuid-sandbox', 
            '--disable-dev-shm-usage',
            '--disable-gpu',
            '--single-process',
            '--disable-blink-features=AutomationControlled',
            '--window-size=1920,1080'
        ]
    });
    
    const page = await browser.newPage();
    await page.setViewport({ width: 1920, height: 1080 });
    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36');
    
    // Esconder que é automação
    await page.evaluateOnNewDocument(() => {
        Object.defineProperty(navigator, 'webdriver', { get: () => false });
    });
    
    const sailings = [];
    const polData = PORTS[pol];
    const podData = PORTS[pod];
    
    try {
        // 1. Acessar MSC Schedules
        console.log('1. Acessando MSC...');
        await page.goto('https://www.msc.com/en/search-a-schedule', { 
            waitUntil: 'domcontentloaded', // MAIS RÁPIDO - não espera imagens
            timeout: 60000 
        });
        
        // 2. FECHAR POPUP DE COOKIES IMEDIATAMENTE
        console.log('2. Fechando popup de cookies...');
        try {
            // Esperar só 1 segundo pelo botão de cookies
            await page.waitForSelector('button#onetrust-accept-btn-handler', { timeout: 3000 });
            await page.click('button#onetrust-accept-btn-handler');
            console.log('   ✅ Cookies aceitos');
        } catch (e) {
            console.log('   Popup de cookies não encontrado ou já fechado');
        }
        await new Promise(r => setTimeout(r, 500));
        
        await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-01-inicio.png') });
        
        await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-02-cookies-fechado.png') });
        
        // Verificar se os portos existem no mapeamento
        if (!polData) {
            throw new Error(`Porto de origem "${pol}" não encontrado no sistema`);
        }
        if (!podData) {
            throw new Error(`Porto de destino "${pod}" não encontrado no sistema`);
        }
        
        // 3. Preencher ORIGEM (From)
        console.log(`3. Preenchendo origem: ${polData.search}`);
        
        // Clicar no campo From
        const fromInput = await page.$('input[placeholder*="From"], input[aria-label*="From"], input[name*="from"], input[id*="from"]');
        if (fromInput) {
            await fromInput.click();
            await new Promise(r => setTimeout(r, 300));
            await fromInput.type(polData.search, { delay: 30 }); // MAIS RÁPIDO
            await new Promise(r => setTimeout(r, 1000)); // Esperar autocomplete
            
            // Selecionar primeira opção do dropdown
            await page.keyboard.press('ArrowDown');
            await new Promise(r => setTimeout(r, 150));
            await page.keyboard.press('Enter');
            await new Promise(r => setTimeout(r, 500));
        } else {
            // Tentar pelo placeholder em português
            await page.evaluate((searchText) => {
                const inputs = document.querySelectorAll('input');
                for (const input of inputs) {
                    const placeholder = (input.placeholder || '').toLowerCase();
                    if (placeholder.includes('from') || placeholder.includes('de') || placeholder.includes('origin')) {
                        input.click();
                        input.value = searchText;
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        return;
                    }
                }
            }, polData.search);
            await new Promise(r => setTimeout(r, 1000));
            await page.keyboard.press('ArrowDown');
            await page.keyboard.press('Enter');
        }
        
        await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-03-origem.png') });
        await new Promise(r => setTimeout(r, 300));
        
        // 4. Preencher DESTINO (To)
        console.log(`4. Preenchendo destino: ${podData.search}`);
        
        const toInput = await page.$('input[placeholder*="To"], input[aria-label*="To"], input[name*="to"], input[id*="to"]');
        if (toInput) {
            await toInput.click();
            await new Promise(r => setTimeout(r, 300));
            await toInput.type(podData.search, { delay: 30 }); // MAIS RÁPIDO
            await new Promise(r => setTimeout(r, 1000)); // Esperar autocomplete
            await page.keyboard.press('ArrowDown');
            await new Promise(r => setTimeout(r, 150));
            await page.keyboard.press('Enter');
            await new Promise(r => setTimeout(r, 500));
        } else {
            await page.evaluate((searchText) => {
                const inputs = document.querySelectorAll('input');
                for (const input of inputs) {
                    const placeholder = (input.placeholder || '').toLowerCase();
                    if (placeholder.includes('to') || placeholder.includes('para') || placeholder.includes('dest')) {
                        input.click();
                        input.value = searchText;
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        return;
                    }
                }
            }, podData.search);
            await new Promise(r => setTimeout(r, 1000));
            await page.keyboard.press('ArrowDown');
            await page.keyboard.press('Enter');
        }
        
        await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-04-destino.png') });
        await new Promise(r => setTimeout(r, 300));
        
        // 5. Clicar em SEARCH (botão amarelo "> Search")
        console.log('5. Clicando no botão Search...');
        
        // MÉTODO DIRETO E RÁPIDO: Encontrar botão pelo texto e clicar
        const searchClicked = await page.evaluate(() => {
            const elements = document.querySelectorAll('button, a, div');
            for (const el of elements) {
                const text = (el.innerText || '').trim();
                const rect = el.getBoundingClientRect();
                
                // Botão Search: texto "Search" ou "> Search", abaixo do header (y > 300)
                if ((text === 'Search' || text === '> Search' || text.includes('Search')) && 
                    rect.width > 80 && rect.width < 200 &&
                    rect.height > 30 && rect.height < 60 && 
                    rect.y > 300 && rect.y < 500) {
                    el.click();
                    return { success: true, y: rect.y };
                }
            }
            return { success: false };
        });
        
        if (searchClicked.success) {
            console.log(`   ✅ Search clicado (y=${searchClicked.y})`);
        } else {
            // Fallback: clicar por coordenadas
            console.log('   Tentando coordenadas...');
            await page.mouse.click(1027, 430);
        }
        
        await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-04b-antes-search.png') });
        
        // Esperar carregamento dos resultados (mínimo necessário)
        console.log('6. Aguardando resultados (3s)...');
        await new Promise(r => setTimeout(r, 3000));
        
        // Scroll para o topo IMEDIATAMENTE (antes do screenshot)
        await page.evaluate(() => window.scrollTo(0, 0));
        
        // Screenshot dos resultados
        await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-05-resultados.png') });
        
        // 6.5 Filtrar por serviço se especificado
        if (service) {
            console.log(`6.5. Filtrando por serviço: ${service}...`);
            try {
                
                // Encontrar posição do filtro
                const filterPos = await page.evaluate(() => {
                    const allElements = document.querySelectorAll('*');
                    
                    for (const el of allElements) {
                        const text = (el.innerText || '').trim();
                        const rect = el.getBoundingClientRect();
                        
                        if (text === 'Filter by: All Services' &&
                            rect.y > 0 && rect.y < 800 &&
                            rect.width > 100 && rect.width < 300 &&
                            rect.height > 30 && rect.height < 70) {
                            return { 
                                found: true, 
                                x: rect.x + rect.width / 2, 
                                y: rect.y + rect.height / 2 
                            };
                        }
                    }
                    return { found: false };
                });
                
                if (filterPos.found) {
                    console.log(`   ✅ Filtro em (${Math.round(filterPos.x)}, ${Math.round(filterPos.y)})`);
                    await page.mouse.click(filterPos.x, filterPos.y); // MOUSE CLICK
                    
                    // Esperar dropdown abrir (aumentado para 2s)
                    await new Promise(r => setTimeout(r, 2000));
                    
                    await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-06-filtro-aberto.png') });
                    
                    // Encontrar posição da opção do serviço
                    const serviceText = service + ' Service';
                    const servicePos = await page.evaluate((targetText, serviceName) => {
                        const allElements = document.querySelectorAll('*');
                        let bestMatch = null;
                        let smallestArea = Infinity;
                        
                        for (const el of allElements) {
                            const text = (el.innerText || '').trim();
                            const rect = el.getBoundingClientRect();
                            
                            // Busca mais flexível: texto exato OU contém o nome do serviço
                            const isMatch = (text === targetText) || 
                                           (text.toLowerCase().includes(serviceName.toLowerCase()) && 
                                            text.toLowerCase().includes('service'));
                            
                            if (isMatch && rect.y > 0 && rect.width > 0 && rect.height > 0 && rect.width < 300) {
                                const area = rect.width * rect.height;
                                if (area < smallestArea) {
                                    smallestArea = area;
                                    bestMatch = { 
                                        found: true, 
                                        x: rect.x + rect.width / 2, 
                                        y: rect.y + rect.height / 2,
                                        text: text
                                    };
                                }
                            }
                        }
                        return bestMatch || { found: false };
                    }, serviceText, service);
                    
                    if (servicePos.found) {
                        console.log(`   ✅ "${servicePos.text}" em (${Math.round(servicePos.x)}, ${Math.round(servicePos.y)})`);
                        await page.mouse.click(servicePos.x, servicePos.y); // MOUSE CLICK
                        console.log(`   ✅ Selecionado!`);
                        await new Promise(r => setTimeout(r, 1500));
                    } else {
                        console.log(`   ⚠️ ${serviceText} não encontrado`);
                    }
                } else {
                    console.log('   ⚠️ Filtro não encontrado');
                }
                
                await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-07-filtro-aplicado.png') });
                
            } catch (e) {
                console.log('   Erro ao filtrar:', e.message);
            }
        }
        
        // 7. Extrair dados da tabela de resultados
        console.log('7. Extraindo dados...');
        
        const data = await page.evaluate((filterService) => {
            const results = [];
            const seenVessels = new Set();
            
            // A MSC mostra resultados em cards/divs com estrutura:
            // Departure | Arrival | Vessel / Voyage No. | Estimated Transit Time | Routing Type
            
            // Pegar o texto de toda a área de resultados
            const resultsArea = document.body.innerText;
            
            // Padrão para extrair linhas de schedule:
            // "Sun 11th Jan 2026 Sat 14th Feb 2026 SANTA CATARINA EXPRESS / 2552W 35 Days Direct"
            const lines = resultsArea.split('\n');
            
            let currentDeparture = null;
            let currentArrival = null;
            let currentVessel = null;
            let currentTransit = null;
            let currentRouting = null;
            
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i].trim();
                
                // Procurar data de departure (formato: "Sun 11th Jan 2026")
                const dateMatch = line.match(/^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+\d{1,2}(?:st|nd|rd|th)?\s+\w{3}\s+\d{4}$/i);
                if (dateMatch) {
                    if (!currentDeparture) {
                        currentDeparture = line;
                    } else if (!currentArrival) {
                        currentArrival = line;
                    }
                    continue;
                }
                
                // Procurar navio (formato: "NOME / CODIGO" ou só "NOME")
                const vesselMatch = line.match(/^([A-Z][A-Z\s]+)\s*\/\s*[A-Z0-9]+W?$/i) ||
                                   line.match(/^([A-Z][A-Z\s]+)\s*$/);
                if (vesselMatch && line.length > 5 && line.length < 50) {
                    // Verificar se não é uma palavra comum
                    const excluded = ['DEPARTURE', 'ARRIVAL', 'VESSEL', 'VOYAGE', 'DIRECT', 'TRANSHIPMENT', 'FILTER', 'RESULTS', 'POINT', 'SERVICES'];
                    const possibleVessel = vesselMatch[1] || line;
                    if (!excluded.some(ex => possibleVessel.toUpperCase().includes(ex))) {
                        currentVessel = possibleVessel.replace(/\s*\/.*/, '').trim();
                    }
                    continue;
                }
                
                // Procurar transit time
                const transitMatch = line.match(/^(\d+)\s*Days?$/i);
                if (transitMatch) {
                    currentTransit = transitMatch[1] + ' dias';
                    continue;
                }
                
                // Procurar routing type
                if (line === 'Direct' || line === 'Transhipment') {
                    currentRouting = line === 'Transhipment' ? 'Transbordo' : line;
                    
                    // Temos um registro completo!
                    if (currentVessel && !seenVessels.has(currentVessel)) {
                        seenVessels.add(currentVessel);
                        results.push({
                            service: filterService || '-',
                            vessel: currentVessel,
                            etd: currentDeparture || '-',
                            eta: currentArrival || '-',
                            transit: currentTransit || '-',
                            routeType: currentRouting || '-'
                        });
                    }
                    
                    // Reset para próximo registro
                    currentDeparture = null;
                    currentArrival = null;
                    currentVessel = null;
                    currentTransit = null;
                    currentRouting = null;
                }
            }
            
            return {
                results,
                count: results.length
            };
        }, service);
        
        console.log(`   Encontrados: ${data.results.length} navios únicos`);
        
        // Processar resultados
        data.results.forEach(r => {
            sailings.push({
                carrier: 'MSC',
                service: r.service,
                vessel: r.vessel,
                pol,
                pod,
                etd: r.etd,
                eta: r.eta,
                transit: r.transit,
                routeType: r.routeType,
                source: 'MSC Website'
            });
        });
        
    } catch (error) {
        console.log(`❌ Erro: ${error.message}`);
        await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-error.png') });
    } finally {
        await browser.close();
    }
    
    console.log(`✅ Total: ${sailings.length} schedules\n`);
    return sailings;
}

// ============================================
// EXCEL
// ============================================
async function generateExcel(sailings, pol, pod) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('SCHEDULES');
    
    // Header
    sheet.mergeCells('A1:F1');
    sheet.getCell('A1').value = `ALLOG - MSC Schedules: ${pol} → ${pod} (${new Date().toLocaleDateString('pt-BR')})`;
    sheet.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFF' } };
    sheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD700' } };
    sheet.getCell('A1').alignment = { horizontal: 'center' };
    
    // Colunas
    const headers = ['SERVIÇO', 'NAVIO', 'ETD', 'ETA', 'TRANSIT', 'TIPO'];
    sheet.addRow(headers);
    sheet.getRow(2).font = { bold: true, color: { argb: 'FFFFFF' } };
    sheet.getRow(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '333333' } };
    
    // Dados
    sailings.forEach(s => {
        const row = sheet.addRow([s.service, s.vessel, s.etd, s.eta, s.transit || s.transitTime, s.routeType || '-']);
        row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDE7' } };
    });
    
    sheet.columns = [
        { width: 12 }, { width: 25 }, { width: 18 }, { width: 18 }, { width: 12 }, { width: 12 }
    ];
    
    const filename = `MSC_${pol}_${pod}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    const filepath = path.join(EXPORTS_DIR, filename);
    await workbook.xlsx.writeFile(filepath);
    return { filepath, filename };
}

// ============================================
// ENDPOINTS
// ============================================
app.get('/api/ports', (req, res) => {
    res.json(Object.entries(PORTS).map(([k, v]) => ({ key: k, ...v })));
});

app.get('/api/services', (req, res) => {
    res.json(['Santana', 'Carioca', 'Ipanema', 'Jade']);
});

app.post('/api/schedules', async (req, res) => {
    const { pol, pod, service } = req.body;
    if (!pol || !pod) return res.status(400).json({ error: 'POL e POD obrigatórios' });
    
    try {
        const sailings = await scrapeMSC(pol, pod, service);
        res.json({ success: true, count: sailings.length, sailings });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/export', async (req, res) => {
    const { pol, pod, sailings } = req.body;
    try {
        const { filepath, filename } = await generateExcel(sailings || [], pol, pod);
        res.download(filepath, filename);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.get('/api/screenshots', (req, res) => {
    const files = fs.readdirSync(EXPORTS_DIR).filter(f => f.endsWith('.png')).sort();
    res.json(files);
});

app.get('/api/screenshot/:name', (req, res) => {
    const fp = path.join(EXPORTS_DIR, req.params.name);
    if (fs.existsSync(fp)) res.sendFile(fp);
    else res.status(404).send('Not found');
});

app.get('/', (req, res) => res.sendFile(path.join(__dirname, '../frontend/index.html')));

// ============================================
// START
// ============================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`
╔═══════════════════════════════════════════════════════╗
║                                                       ║
║   🟡 ALLOG - MSC Schedules v12                        ║
║                                                       ║
║   Acesse: http://localhost:${PORT}                       ║
║                                                       ║
║   Serviços: Santana, Carioca, Ipanema, Jade           ║
║                                                       ║
╚═══════════════════════════════════════════════════════╝
    `);
});
