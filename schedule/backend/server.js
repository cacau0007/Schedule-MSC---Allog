// ============================================
// SERVIDOR DE SCHEDULES - MSC, CMA CGM, MAERSK
// Versão atualizada com 6 melhorias
// ============================================

const express = require('express');
const puppeteer = require('puppeteer-core');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;
const EXPORTS_DIR = path.join(__dirname, '../exports');

if (!fs.existsSync(EXPORTS_DIR)) fs.mkdirSync(EXPORTS_DIR, { recursive: true });

app.use(express.json());
app.use('/exports', express.static(EXPORTS_DIR));
app.use(express.static(path.join(__dirname, '../frontend')));

// ============================================
// MAPEAMENTO DE SERVIÇOS POR ROTA (POL/POD)
// ============================================
const SERVICE_ROUTES = {
    // Rotas para Manaus
    'Yantian-Manaus': ['Santana'],
    'Ningbo-Manaus': ['Santana'],
    'Shanghai-Manaus': ['Santana'],
    'Qingdao-Manaus': ['Santana'],
    'Busan-Manaus': ['Santana'],
    
    // Rotas para Navegantes
    'Xiamen-Navegantes': ['Jade', 'Tiger'],
    'Ningbo-Navegantes': ['Ipanema', 'Carioca', 'Jade'],
    'Shanghai-Navegantes': ['Ipanema', 'Carioca'],
    'Qingdao-Navegantes': ['Carioca'],
    'Busan-Navegantes': ['Ipanema', 'Carioca'],
    'Shekou-Navegantes': ['Carioca'],
    
    // Rotas para Santos
    'Busan-Santos': ['Ipanema', 'Carioca'],
    'Shanghai-Santos': ['Ipanema', 'Carioca'],
    'Ningbo-Santos': ['Ipanema', 'Carioca'],
    'Shekou-Santos': ['Carioca'],
    'Qingdao-Santos': ['Carioca'],
    
    // Rotas para Itajaí
    'Busan-Itajai': ['Carioca'],
    'Shanghai-Itajai': ['Carioca'],
    'Ningbo-Itajai': ['Carioca'],
    'Shekou-Itajai': ['Carioca'],
    'Qingdao-Itajai': ['Carioca'],
    
    // Rotas para Paranaguá
    'Busan-Paranagua': ['Ipanema', 'Carioca'],
    'Shanghai-Paranagua': ['Ipanema', 'Carioca'],
    'Ningbo-Paranagua': ['Ipanema', 'Carioca'],
    'Shekou-Paranagua': ['Carioca'],
    
    // Rotas para Rio de Janeiro
    'Busan-Rio de Janeiro': ['Ipanema', 'Carioca'],
    'Shanghai-Rio de Janeiro': ['Ipanema', 'Carioca'],
    'Ningbo-Rio de Janeiro': ['Ipanema', 'Carioca'],
    'Shekou-Rio de Janeiro': ['Carioca'],
    
    // Rotas para Salvador
    'Yantian-Salvador': ['Santana'],
    'Ningbo-Salvador': ['Santana'],
    'Shanghai-Salvador': ['Santana'],
    'Qingdao-Salvador': ['Santana'],
    'Busan-Salvador': ['Santana'],
    
    // Rotas para Suape
    'Yantian-Suape': ['Santana'],
    'Ningbo-Suape': ['Santana'],
    'Shanghai-Suape': ['Santana'],
    'Qingdao-Suape': ['Santana'],
    'Busan-Suape': ['Santana']
};

// Função para obter serviços disponíveis para uma rota
function getAvailableServices(pol, pod) {
    const routeKey = `${pol}-${pod}`;
    return SERVICE_ROUTES[routeKey] || null;
}

// ============================================
// ENDPOINT: Obter serviços disponíveis
// ============================================
app.get('/api/available-services', (req, res) => {
    const { pol, pod } = req.query;
    
    if (!pol || !pod) {
        return res.json({ services: null, error: 'POL e POD são obrigatórios' });
    }
    
    const services = getAvailableServices(pol, pod);
    
    if (services === null) {
        return res.json({ 
            services: ['ALL'], 
            message: 'Rota não mapeada - permitindo todos os serviços'
        });
    }
    
    return res.json({ 
        services: services,
        allowAll: true // Sempre permitir "Todas as rotas"
    });
});

// ============================================
// ENDPOINT: Buscar schedules
// ============================================
app.post('/api/schedules', async (req, res) => {
    const { pol, pod, carriers, service } = req.body;
    
    console.log('\n=== NOVA REQUISIÇÃO ===');
    console.log('POL:', pol, '| POD:', pod);
    console.log('Carriers:', carriers);
    console.log('Service:', service || 'ALL');
    
    const results = [];
    
    try {
        if (carriers.includes('MSC')) {
            const mscData = await scrapeMSC(pol, pod, service);
            results.push(...mscData);
        }
        
        const filename = `Schedules_${pol}_${pod}_${new Date().toISOString().slice(0, 10)}.xlsx`;
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
// SCRAPER: MSC
// ============================================
async function scrapeMSC(pol, pod, service = null) {
    const sailings = [];
    
    console.log(`\n🚢 === MSC SCRAPER ===`);
    console.log(`Rota: ${pol} → ${pod}`);
    if (service && service !== 'ALL') {
        console.log(`Serviço: ${service}`);
    }
    
    let browser;
    let page;
    
    try {
        browser = await puppeteer.launch({
            headless: true,
            executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || '/usr/bin/google-chrome-stable',
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--disable-web-security',
                '--window-size=1920,1080'
            ]
        });
        
        page = await browser.newPage();
        await page.setViewport({ width: 1920, height: 1080 });
        
        // 1. Acessar site
        console.log('1. Acessando site MSC...');
        await page.goto('https://www.msc.com/en/search-a-schedule', {
            waitUntil: 'networkidle0',
            timeout: 60000
        });
        await new Promise(r => setTimeout(r, 2000));
        
        // 2. Selecionar POL (Port of Loading) - USANDO CTRL+C/CTRL+V
        console.log(`2. Selecionando POL: ${pol}...`);
        const polInput = await page.$('input[placeholder*="Port of loading"]');
        if (polInput) {
            await polInput.click();
            await new Promise(r => setTimeout(r, 500));
            
            // Copiar texto para clipboard (simulado com evaluate)
            await page.evaluate((text) => {
                const input = document.querySelector('input[placeholder*="Port of loading"]');
                if (input) {
                    input.value = text;
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                }
            }, pol);
            
            await new Promise(r => setTimeout(r, 1500));
            await page.keyboard.press('ArrowDown');
            await page.keyboard.press('Enter');
            console.log('   ✅ POL selecionado');
        }
        
        // 3. Selecionar POD (Port of Discharge) - USANDO CTRL+C/CTRL+V
        console.log(`3. Selecionando POD: ${pod}...`);
        await new Promise(r => setTimeout(r, 800));
        const podInput = await page.$('input[placeholder*="Port of discharge"]');
        if (podInput) {
            await podInput.click();
            await new Promise(r => setTimeout(r, 500));
            
            await page.evaluate((text) => {
                const input = document.querySelector('input[placeholder*="Port of discharge"]');
                if (input) {
                    input.value = text;
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                }
            }, pod);
            
            await new Promise(r => setTimeout(r, 1500));
            await page.keyboard.press('ArrowDown');
            await page.keyboard.press('Enter');
            console.log('   ✅ POD selecionado');
        }
        
        // 4. Clicar em Search - MÚLTIPLAS ESTRATÉGIAS
        console.log('4. Clicando em Search...');
        await new Promise(r => setTimeout(r, 1000));
        
        let searchClicked = false;
        
        // Método 1: Buscar por texto exato
        try {
            const searchBtn = await page.evaluateHandle(() => {
                const buttons = Array.from(document.querySelectorAll('button, a, div[role="button"]'));
                return buttons.find(btn => {
                    const text = (btn.innerText || btn.textContent || '').toLowerCase().trim();
                    const rect = btn.getBoundingClientRect();
                    // Ignorar elementos do header (y < 200)
                    if (rect.y < 200 || rect.height < 20) return false;
                    return text.includes('search') || text.includes('pesquisar') || text.includes('buscar');
                });
            });
            
            const element = searchBtn.asElement();
            if (element) {
                await element.click();
                searchClicked = true;
                console.log('   ✅ Search clicado (Método 1 - Texto)');
            }
        } catch (e) {
            console.log('   ⚠️ Método 1 falhou:', e.message);
        }
        
        // Método 2: Buscar botão amarelo por cor
        if (!searchClicked) {
            try {
                const yellowBtn = await page.evaluateHandle(() => {
                    const buttons = Array.from(document.querySelectorAll('button, a, div[role="button"]'));
                    return buttons.find(btn => {
                        const style = getComputedStyle(btn);
                        const bg = style.backgroundColor;
                        const rect = btn.getBoundingClientRect();
                        if (rect.y < 200 || rect.y > 600) return false;
                        // Amarelo: R > 200, G > 150, B < 100
                        const match = bg.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
                        if (match) {
                            const [, r, g, b] = match.map(Number);
                            return r > 200 && g > 150 && b < 100;
                        }
                        return false;
                    });
                });
                
                const element = yellowBtn.asElement();
                if (element) {
                    await element.click();
                    searchClicked = true;
                    console.log('   ✅ Search clicado (Método 2 - Botão amarelo)');
                }
            } catch (e) {
                console.log('   ⚠️ Método 2 falhou:', e.message);
            }
        }
        
        // Método 3: Clicar por coordenadas (posição típica do botão)
        if (!searchClicked) {
            try {
                await page.mouse.click(950, 420);
                searchClicked = true;
                console.log('   ✅ Search clicado (Método 3 - Coordenadas)');
            } catch (e) {
                console.log('   ⚠️ Método 3 falhou:', e.message);
            }
        }
        
        // Método 4: Pressionar Enter
        if (!searchClicked) {
            try {
                await page.keyboard.press('Enter');
                console.log('   ✅ Search via Enter (Método 4)');
            } catch (e) {
                console.log('   ⚠️ Método 4 falhou:', e.message);
            }
        }
        
        // 5. Aguardar resultados
        console.log('5. Aguardando resultados (3s)...');
        await new Promise(r => setTimeout(r, 3000));
        
        await page.evaluate(() => window.scrollTo(0, 0));
        await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-05-resultados.png') });
        
        // 6. Filtrar por serviço (SE ESPECIFICADO E DIFERENTE DE 'ALL')
        if (service && service !== 'ALL') {
            console.log(`6. Filtrando por serviço: ${service}...`);
            try {
                // MELHORIA 1: Buscar filtro por TEXTO ao invés de coordenadas
                const filterFound = await page.evaluate(() => {
                    const allElements = document.querySelectorAll('*');
                    for (const el of allElements) {
                        const text = (el.innerText || '').trim();
                        // Procurar por "Filter by: All Services" ou variações
                        if (text.includes('Filter by') && text.includes('All Services')) {
                            const rect = el.getBoundingClientRect();
                            if (rect.width > 100 && rect.height > 20) {
                                el.click();
                                return true;
                            }
                        }
                    }
                    return false;
                });
                
                if (filterFound) {
                    console.log('   ✅ Filtro encontrado e clicado');
                    
                    // Esperar dropdown abrir (MELHORIA 1: aumentado de 800ms para 2500ms)
                    await new Promise(r => setTimeout(r, 2500));
                    
                    await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-06-filtro-aberto.png') });
                    
                    // Buscar e clicar no serviço
                    const serviceText = service + ' Service';
                    const serviceClicked = await page.evaluate((targetService) => {
                        const allElements = document.querySelectorAll('*');
                        let bestMatch = null;
                        let smallestArea = Infinity;
                        
                        for (const el of allElements) {
                            const text = (el.innerText || '').trim();
                            const rect = el.getBoundingClientRect();
                            
                            // Busca flexível: aceita variações do nome
                            const matches = 
                                text === targetService ||
                                text === targetService.toUpperCase() ||
                                text === targetService.toLowerCase() ||
                                (text.includes(targetService.split(' ')[0]) && text.includes('Service'));
                            
                            if (matches && rect.y > 0 && rect.width > 0 && rect.height > 0 && rect.width < 300) {
                                const area = rect.width * rect.height;
                                if (area < smallestArea) {
                                    smallestArea = area;
                                    bestMatch = el;
                                }
                            }
                        }
                        
                        if (bestMatch) {
                            bestMatch.click();
                            return true;
                        }
                        return false;
                    }, serviceText);
                    
                    if (serviceClicked) {
                        console.log(`   ✅ ${serviceText} selecionado!`);
                        await new Promise(r => setTimeout(r, 1500));
                    } else {
                        console.log(`   ⚠️ ${serviceText} não encontrado no dropdown`);
                    }
                } else {
                    console.log('   ⚠️ Filtro não encontrado');
                }
                
                await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-07-filtro-aplicado.png') });
            } catch (e) {
                console.log('   ⚠️ Erro ao filtrar:', e.message);
            }
        }
        
        // 7. Extrair dados
        console.log('7. Extraindo dados...');
        
        const data = await page.evaluate((filterService) => {
            const results = [];
            const seenVessels = new Map(); // MELHORIA 5: usar Map para armazenar transit time
            const resultsArea = document.body.innerText;
            const lines = resultsArea.split('\n');
            
            let currentDeparture = null;
            let currentArrival = null;
            let currentVessel = null;
            let currentTransit = null;
            let currentRouting = null;
            
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i].trim();
                
                // Detectar datas
                const dateMatch = line.match(/^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+\d{1,2}(?:st|nd|rd|th)?\s+\w{3}\s+\d{4}$/i);
                if (dateMatch) {
                    if (!currentDeparture) {
                        currentDeparture = line;
                    } else if (!currentArrival) {
                        currentArrival = line;
                    }
                    continue;
                }
                
                // Detectar navio
                const vesselMatch = line.match(/^([A-Z][A-Z\s]+)\s*\/\s*[A-Z0-9]+W?$/i) ||
                                   line.match(/^([A-Z][A-Z\s]+)\s*$/);
                if (vesselMatch && line.length > 5 && line.length < 50) {
                    const excluded = ['DEPARTURE', 'ARRIVAL', 'VESSEL', 'VOYAGE', 'DIRECT', 'TRANSHIPMENT', 'FILTER', 'RESULTS', 'POINT', 'SERVICES'];
                    const possibleVessel = vesselMatch[1] || line;
                    if (!excluded.some(ex => possibleVessel.toUpperCase().includes(ex))) {
                        currentVessel = possibleVessel.replace(/\s*\/.*/, '').trim();
                    }
                    continue;
                }
                
                // Detectar transit time
                const transitMatch = line.match(/^(\d+)\s*Days?$/i);
                if (transitMatch) {
                    currentTransit = parseInt(transitMatch[1]); // MELHORIA 5: armazenar como número
                    continue;
                }
                
                // Detectar tipo de rota
                if (line === 'Direct' || line === 'Transhipment') {
                    currentRouting = line === 'Transhipment' ? 'Transbordo' : line;
                    
                    if (currentVessel) {
                        // MELHORIA 5: Verificar se já existe e comparar transit time
                        const existingEntry = seenVessels.get(currentVessel);
                        
                        if (!existingEntry) {
                            // Primeira vez vendo este navio
                            seenVessels.set(currentVessel, {
                                service: filterService || '-',
                                vessel: currentVessel,
                                etd: currentDeparture || '-',
                                eta: currentArrival || '-',
                                transit: currentTransit || 0,
                                routeType: currentRouting || '-'
                            });
                        } else {
                            // Navio duplicado - manter o com MAIOR transit time
                            if (currentTransit > existingEntry.transit) {
                                seenVessels.set(currentVessel, {
                                    service: filterService || '-',
                                    vessel: currentVessel,
                                    etd: currentDeparture || '-',
                                    eta: currentArrival || '-',
                                    transit: currentTransit,
                                    routeType: currentRouting || '-'
                                });
                            }
                        }
                    }
                    
                    currentDeparture = null;
                    currentArrival = null;
                    currentVessel = null;
                    currentTransit = null;
                    currentRouting = null;
                }
            }
            
            // Converter Map para Array
            const uniqueResults = Array.from(seenVessels.values()).map(entry => ({
                ...entry,
                transit: entry.transit ? `${entry.transit} dias` : '-'
            }));
            
            return { results: uniqueResults, count: uniqueResults.length };
        }, service && service !== 'ALL' ? service : null);
        
        console.log(`   Encontrados: ${data.results.length} navios únicos`);
        
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
        if (page) {
            await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-error.png') }).catch(() => {});
        }
    } finally {
        if (browser) {
            await browser.close();
        }
    }
    
    console.log(`✅ Total: ${sailings.length} schedules\n`);
    return sailings;
}

// ============================================
// EXCEL - MELHORIA 6: Formato de data melhorado
// ============================================
async function generateExcel(sailings, pol, pod, filename) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('SCHEDULES');
    
    sheet.mergeCells('A1:G1');
    sheet.getCell('A1').value = `ALLOG - Shipping Schedules: ${pol} → ${pod}`;
    sheet.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFF' } };
    sheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD700' } };
    sheet.getCell('A1').alignment = { horizontal: 'center' };
    
    const headers = ['CARRIER', 'SERVIÇO', 'NAVIO', 'ETD', 'ETA', 'TRANSIT', 'TIPO'];
    sheet.addRow(headers);
    sheet.getRow(2).font = { bold: true, color: { argb: 'FFFFFF' } };
    sheet.getRow(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '333333' } };
    
    sailings.forEach(s => {
        // MELHORIA 6: Formatar data de "Mon 13th Jan 2025" para "Mon - 13/01/2025"
        const etdFormatted = formatDate(s.etd);
        const etaFormatted = formatDate(s.eta);
        
        const row = sheet.addRow([
            s.carrier,
            s.service,
            s.vessel,
            etdFormatted,
            etaFormatted,
            s.transit || s.transitTime || '-',
            s.routeType || '-'
        ]);
        row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDE7' } };
    });
    
    sheet.columns = [
        { width: 10 }, // CARRIER
        { width: 12 }, // SERVIÇO
        { width: 25 }, // NAVIO
        { width: 18 }, // ETD
        { width: 18 }, // ETA
        { width: 12 }, // TRANSIT
        { width: 12 }  // TIPO
    ];
    
    const filepath = path.join(EXPORTS_DIR, filename);
    await workbook.xlsx.writeFile(filepath);
    console.log(`📊 Excel gerado: ${filename}`);
}

// MELHORIA 6: Função para formatar data
function formatDate(dateStr) {
    if (!dateStr || dateStr === '-') return '-';
    
    try {
        // Entrada: "Mon 13th Jan 2025" ou "Sat 17th Jan 2026"
        // Saída: "Mon - 13/01/2025" ou "Sat - 17/01/2026"
        
        const match = dateStr.match(/^(\w+)\s+(\d{1,2})(?:st|nd|rd|th)?\s+(\w+)\s+(\d{4})$/);
        if (!match) return dateStr; // Se não der match, retorna original
        
        const [, dayOfWeek, day, monthName, year] = match;
        
        const months = {
            'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04',
            'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08',
            'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'
        };
        
        const month = months[monthName];
        if (!month) return dateStr;
        
        const dayPadded = day.padStart(2, '0');
        
        return `${dayOfWeek} - ${dayPadded}/${month}/${year}`;
    } catch (e) {
        return dateStr;
    }
}

// ============================================
// SERVIDOR
// ============================================
app.listen(PORT, () => {
    console.log(`\n🚀 Servidor rodando na porta ${PORT}`);
    console.log(`📁 Exports: ${EXPORTS_DIR}\n`);
});
