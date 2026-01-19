// ============================================
// SERVIDOR DE SCHEDULES - MSC, CMA CGM, MAERSK
// Versão atualizada com 6 melhorias + Anti-bloqueio
// ============================================

const express = require('express');
const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// Ativar plugin stealth para evitar detecção
puppeteer.use(StealthPlugin());

const app = express();
const PORT = process.env.PORT || 3000;
const EXPORTS_DIR = path.join(__dirname, '../exports');

if (!fs.existsSync(EXPORTS_DIR)) fs.mkdirSync(EXPORTS_DIR, { recursive: true });

app.use(express.json());
app.use('/exports', express.static(EXPORTS_DIR));
app.use(express.static(path.join(__dirname, '../frontend')));

// ============================================
// MAPEAMENTO DE SERVIÇOS POR ROTA (POL/POD)
// Baseado nas rotas oficiais MSC 2024/2025:
// - SANTANA: Yantian-Ningbo-Shanghai-Qingdao-Busan-Cristobal-Suape-Salvador (Nordeste)
// - CARIOCA: Qingdao-Busan-Ningbo-Shanghai-Shekou-Singapore-Rio-Santos-Paranagua-Imbituba-Itajai (Sul/Sudeste)
// - IPANEMA: Busan-Shanghai-Ningbo-Shekou-HongKong-Singapore-Santos-Paranagua-Navegantes-Montevideo-BuenosAires-RioGrande
// ============================================
const SERVICE_ROUTES = {
    // ========== SANTANA - Nordeste Brasil ==========
    // POL: Yantian, Ningbo, Shanghai, Qingdao, Busan
    // POD: Suape, Salvador (+ Manaus via feeder)
    'Yantian-Suape': ['Santana'],
    'Ningbo-Suape': ['Santana'],
    'Shanghai-Suape': ['Santana'],
    'Qingdao-Suape': ['Santana'],
    'Busan-Suape': ['Santana'],
    'Yantian-Salvador': ['Santana'],
    'Ningbo-Salvador': ['Santana'],
    'Shanghai-Salvador': ['Santana'],
    'Qingdao-Salvador': ['Santana'],
    'Busan-Salvador': ['Santana'],
    'Yantian-Manaus': ['Santana'],
    'Ningbo-Manaus': ['Santana'],
    'Shanghai-Manaus': ['Santana'],
    'Qingdao-Manaus': ['Santana'],
    'Busan-Manaus': ['Santana'],
    
    // ========== CARIOCA - Sul/Sudeste Brasil ==========
    // POL: Qingdao, Busan, Ningbo, Shanghai, Shekou
    // POD: Rio de Janeiro, Santos, Paranagua, Imbituba, Itajai, Itaguai
    'Qingdao-Rio de Janeiro': ['Carioca'],
    'Busan-Rio de Janeiro': ['Carioca'],
    'Ningbo-Rio de Janeiro': ['Carioca'],
    'Shanghai-Rio de Janeiro': ['Carioca'],
    'Shekou-Rio de Janeiro': ['Carioca'],
    'Qingdao-Santos': ['Carioca'],
    'Busan-Santos': ['Carioca', 'Ipanema'],
    'Ningbo-Santos': ['Carioca', 'Ipanema'],
    'Shanghai-Santos': ['Carioca', 'Ipanema'],
    'Shekou-Santos': ['Carioca', 'Ipanema'],
    'Qingdao-Paranagua': ['Carioca'],
    'Busan-Paranagua': ['Carioca', 'Ipanema'],
    'Ningbo-Paranagua': ['Carioca', 'Ipanema'],
    'Shanghai-Paranagua': ['Carioca', 'Ipanema'],
    'Shekou-Paranagua': ['Carioca', 'Ipanema'],
    'Qingdao-Itajai': ['Carioca'],
    'Busan-Itajai': ['Carioca'],
    'Ningbo-Itajai': ['Carioca'],
    'Shanghai-Itajai': ['Carioca'],
    'Shekou-Itajai': ['Carioca'],
    'Qingdao-Imbituba': ['Carioca'],
    'Busan-Imbituba': ['Carioca'],
    'Ningbo-Imbituba': ['Carioca'],
    'Shanghai-Imbituba': ['Carioca'],
    'Shekou-Imbituba': ['Carioca'],
    'Qingdao-Itaguai': ['Carioca'],
    'Busan-Itaguai': ['Carioca'],
    'Ningbo-Itaguai': ['Carioca'],
    'Shanghai-Itaguai': ['Carioca'],
    'Shekou-Itaguai': ['Carioca'],
    
    // ========== IPANEMA - Sul Brasil + Argentina/Uruguai ==========
    // POL: Busan, Shanghai, Ningbo, Shekou, Hong Kong, Singapore
    // POD: Santos, Paranagua, Navegantes, Montevideo, Buenos Aires, Rio Grande
    'Busan-Navegantes': ['Ipanema'],
    'Shanghai-Navegantes': ['Ipanema'],
    'Ningbo-Navegantes': ['Ipanema'],
    'Shekou-Navegantes': ['Ipanema'],
    'Hong Kong-Navegantes': ['Ipanema'],
    'Singapore-Navegantes': ['Ipanema'],
    'Busan-Rio Grande': ['Ipanema'],
    'Shanghai-Rio Grande': ['Ipanema'],
    'Ningbo-Rio Grande': ['Ipanema'],
    'Shekou-Rio Grande': ['Ipanema'],
    'Hong Kong-Rio Grande': ['Ipanema'],
    'Busan-Montevideo': ['Ipanema'],
    'Shanghai-Montevideo': ['Ipanema'],
    'Ningbo-Montevideo': ['Ipanema'],
    'Shekou-Montevideo': ['Ipanema'],
    'Busan-Buenos Aires': ['Ipanema'],
    'Shanghai-Buenos Aires': ['Ipanema'],
    'Ningbo-Buenos Aires': ['Ipanema'],
    'Shekou-Buenos Aires': ['Ipanema']
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
            headless: 'new',
            executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || '/usr/bin/google-chrome-stable',
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--disable-web-security',
                '--disable-blink-features=AutomationControlled',
                '--window-size=1920,1080',
                '--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            ]
        });
        
        page = await browser.newPage();
        
        // Anti-detecção: Remover indicadores de automação
        await page.evaluateOnNewDocument(() => {
            // Remover webdriver
            Object.defineProperty(navigator, 'webdriver', { get: () => false });
            
            // Adicionar plugins falsos
            Object.defineProperty(navigator, 'plugins', {
                get: () => [1, 2, 3, 4, 5]
            });
            
            // Adicionar linguagens
            Object.defineProperty(navigator, 'languages', {
                get: () => ['en-US', 'en', 'pt-BR', 'pt']
            });
            
            // Chrome runtime
            window.chrome = { runtime: {} };
            
            // Permissões
            const originalQuery = window.navigator.permissions.query;
            window.navigator.permissions.query = (parameters) => (
                parameters.name === 'notifications' ?
                    Promise.resolve({ state: Notification.permission }) :
                    originalQuery(parameters)
            );
        });
        
        // Headers realistas
        await page.setExtraHTTPHeaders({
            'Accept-Language': 'en-US,en;q=0.9,pt-BR;q=0.8,pt;q=0.7',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'none',
            'Sec-Fetch-User': '?1',
            'Cache-Control': 'max-age=0'
        });
        
        await page.setViewport({ width: 1920, height: 1080 });
        
        // 1. Acessar site
        console.log('1. Acessando site MSC...');
        await page.goto('https://www.msc.com/en/search-a-schedule', {
            waitUntil: 'networkidle2',
            timeout: 60000
        });
        
        // Screenshot para debug
        await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-01-acesso.png') });
        
        // Verificar se foi bloqueado
        const pageContent = await page.content();
        if (pageContent.includes('Access Denied') || pageContent.includes('blocked')) {
            console.log('❌ BLOQUEADO pelo site! Tentando novamente...');
            
            // Tentar recarregar com delay
            await new Promise(r => setTimeout(r, 3000));
            await page.reload({ waitUntil: 'networkidle2' });
            await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-01b-reload.png') });
        }
        
        await new Promise(r => setTimeout(r, 3000));
        
        // Screenshot inicial para debug
        await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-02-antes-preencher.png'), fullPage: true });
        
        // Listar todos os inputs na página para debug
        const inputs = await page.evaluate(() => {
            const allInputs = document.querySelectorAll('input');
            return Array.from(allInputs).map(inp => ({
                placeholder: inp.placeholder,
                id: inp.id,
                name: inp.name,
                className: inp.className,
                type: inp.type
            }));
        });
        console.log('   Inputs encontrados:', inputs.length);
        inputs.forEach((inp, i) => {
            if (inp.placeholder) console.log(`   [${i}] placeholder: "${inp.placeholder}"`);
        });
        
        // 2. Selecionar POL (Port of Loading) - MÚLTIPLAS ESTRATÉGIAS
        console.log(`2. Selecionando POL: ${pol}...`);
        
        // Estratégia 1: Buscar por placeholder (várias variações)
        let polInput = await page.$('input[placeholder*="loading" i]') ||
                       await page.$('input[placeholder*="origin" i]') ||
                       await page.$('input[placeholder*="departure" i]') ||
                       await page.$('input[placeholder*="from" i]') ||
                       await page.$('input[placeholder*="pol" i]');
        
        // Estratégia 2: Buscar pelo primeiro input de texto visível
        if (!polInput) {
            polInput = await page.evaluateHandle(() => {
                const inputs = document.querySelectorAll('input[type="text"], input:not([type])');
                for (const inp of inputs) {
                    const rect = inp.getBoundingClientRect();
                    if (rect.width > 100 && rect.height > 20 && rect.y > 100 && rect.y < 400) {
                        return inp;
                    }
                }
                return null;
            });
            polInput = polInput.asElement();
        }
        
        if (polInput) {
            await polInput.click();
            await new Promise(r => setTimeout(r, 500));
            
            // Limpar e digitar caractere por caractere (mais confiável para autocomplete)
            await polInput.click({ clickCount: 3 }); // Selecionar tudo
            await page.keyboard.type(pol, { delay: 100 });
            
            console.log(`   Digitado: ${pol}`);
            await new Promise(r => setTimeout(r, 2000)); // Esperar autocomplete
            
            // Screenshot após digitar POL
            await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-02b-apos-pol.png') });
            
            // Selecionar primeira opção do autocomplete
            await page.keyboard.press('ArrowDown');
            await new Promise(r => setTimeout(r, 300));
            await page.keyboard.press('Enter');
            console.log('   ✅ POL selecionado');
        } else {
            console.log('   ❌ Input POL não encontrado!');
        }
        
        await new Promise(r => setTimeout(r, 1000));
        
        // 3. Selecionar POD (Port of Discharge) - MÚLTIPLAS ESTRATÉGIAS
        console.log(`3. Selecionando POD: ${pod}...`);
        
        // Estratégia 1: Buscar por placeholder
        let podInput = await page.$('input[placeholder*="discharge" i]') ||
                       await page.$('input[placeholder*="destination" i]') ||
                       await page.$('input[placeholder*="arrival" i]') ||
                       await page.$('input[placeholder*="to" i]') ||
                       await page.$('input[placeholder*="pod" i]');
        
        // Estratégia 2: Segundo input de texto visível
        if (!podInput) {
            podInput = await page.evaluateHandle(() => {
                const inputs = document.querySelectorAll('input[type="text"], input:not([type])');
                let count = 0;
                for (const inp of inputs) {
                    const rect = inp.getBoundingClientRect();
                    if (rect.width > 100 && rect.height > 20 && rect.y > 100 && rect.y < 400) {
                        count++;
                        if (count === 2) return inp; // Pegar o segundo
                    }
                }
                return null;
            });
            podInput = podInput.asElement();
        }
        
        if (podInput) {
            await podInput.click();
            await new Promise(r => setTimeout(r, 500));
            
            await podInput.click({ clickCount: 3 }); // Selecionar tudo
            await page.keyboard.type(pod, { delay: 100 });
            
            console.log(`   Digitado: ${pod}`);
            await new Promise(r => setTimeout(r, 2000)); // Esperar autocomplete
            
            // Screenshot após digitar POD
            await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-03-apos-pod.png') });
            
            await page.keyboard.press('ArrowDown');
            await new Promise(r => setTimeout(r, 300));
            await page.keyboard.press('Enter');
            console.log('   ✅ POD selecionado');
        } else {
            console.log('   ❌ Input POD não encontrado!');
        }
        
        // Screenshot antes de clicar em Search
        await new Promise(r => setTimeout(r, 1000));
        await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-04-antes-search.png') });
        
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
                // Primeiro, encontrar a posição exata do botão "Filter by"
                const filterInfo = await page.evaluate(() => {
                    const allElements = document.querySelectorAll('*');
                    for (const el of allElements) {
                        const text = (el.innerText || el.textContent || '').trim();
                        if (text === 'Filter by: All Services' || 
                            (text.startsWith('Filter by:') && text.includes('All'))) {
                            const rect = el.getBoundingClientRect();
                            if (rect.width > 100 && rect.width < 350 && rect.height > 20 && rect.height < 80) {
                                return { 
                                    found: true, 
                                    text: text, 
                                    x: rect.x + rect.width / 2,
                                    y: rect.y + rect.height / 2,
                                    width: rect.width,
                                    height: rect.height
                                };
                            }
                        }
                    }
                    return { found: false };
                });
                
                if (filterInfo.found) {
                    console.log(`   📍 Botão encontrado: "${filterInfo.text}"`);
                    console.log(`      Posição: x=${Math.round(filterInfo.x)}, y=${Math.round(filterInfo.y)}`);
                    console.log(`      Tamanho: ${Math.round(filterInfo.width)}x${Math.round(filterInfo.height)}`);
                    
                    // Clicar usando mouse.click nas coordenadas exatas
                    await page.mouse.click(filterInfo.x, filterInfo.y);
                    console.log(`   ✅ Clique realizado em (${Math.round(filterInfo.x)}, ${Math.round(filterInfo.y)})`);
                    
                    // Esperar dropdown abrir (tempo maior)
                    await new Promise(r => setTimeout(r, 2500));
                    
                    await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-06-dropdown-aberto.png') });
                    
                    // Agora buscar as opções do dropdown que devem ter aparecido
                    // Dropdown geralmente aparece logo abaixo do botão (y > filterInfo.y)
                    const dropdownY = filterInfo.y + filterInfo.height;
                    
                    const options = await page.evaluate((minY) => {
                        const items = [];
                        // Buscar elementos que parecem ser opções de dropdown
                        const elements = document.querySelectorAll('li, [role="option"], [role="menuitem"], [role="listbox"] li, ul[role="listbox"] li, div[role="listbox"] div');
                        
                        elements.forEach(el => {
                            const text = (el.innerText || el.textContent || '').trim();
                            const rect = el.getBoundingClientRect();
                            
                            // Opções do dropdown: abaixo do botão, tamanho de item de menu
                            if (rect.y >= minY && rect.y < minY + 400 &&
                                rect.height > 20 && rect.height < 60 &&
                                rect.width > 100 &&
                                text.length > 2 && text.length < 50 &&
                                !text.includes('\n')) {
                                items.push({ 
                                    text: text, 
                                    y: Math.round(rect.y),
                                    x: Math.round(rect.x + rect.width/2),
                                    centerY: Math.round(rect.y + rect.height/2)
                                });
                            }
                        });
                        
                        // Se não encontrou com role, tentar buscar por posição
                        if (items.length === 0) {
                            document.querySelectorAll('div, span, li, a').forEach(el => {
                                const text = (el.innerText || el.textContent || '').trim();
                                const rect = el.getBoundingClientRect();
                                
                                // Verificar se parece ser uma opção de serviço
                                const isServiceOption = 
                                    text.toLowerCase().includes('service') ||
                                    text.toLowerCase() === 'all services' ||
                                    text.toLowerCase() === 'carioca' ||
                                    text.toLowerCase() === 'ipanema' ||
                                    text.toLowerCase() === 'santana' ||
                                    text.toLowerCase() === 'tiger' ||
                                    text.toLowerCase() === 'jade';
                                
                                if (isServiceOption && rect.y >= minY && rect.y < minY + 400 &&
                                    rect.height > 15 && rect.height < 60) {
                                    items.push({ 
                                        text: text, 
                                        y: Math.round(rect.y),
                                        x: Math.round(rect.x + rect.width/2),
                                        centerY: Math.round(rect.y + rect.height/2)
                                    });
                                }
                            });
                        }
                        
                        // Remover duplicatas
                        const unique = [...new Map(items.map(i => [i.text, i])).values()];
                        return unique.sort((a, b) => a.y - b.y);
                    }, dropdownY);
                    
                    console.log(`   📋 Opções do dropdown (${options.length}):`);
                    options.forEach((opt, i) => console.log(`      [${i}] "${opt.text}" (y=${opt.y})`));
                    
                    // Clicar no serviço desejado
                    if (options.length > 0) {
                        const targetService = service.toLowerCase();
                        const match = options.find(opt => 
                            opt.text.toLowerCase() === targetService ||
                            opt.text.toLowerCase() === targetService + ' service' ||
                            opt.text.toLowerCase().includes(targetService)
                        );
                        
                        if (match) {
                            console.log(`   🎯 Match encontrado: "${match.text}" em y=${match.centerY}`);
                            await page.mouse.click(match.x, match.centerY);
                            console.log(`   ✅ Serviço "${match.text}" selecionado!`);
                            await new Promise(r => setTimeout(r, 2000));
                        } else {
                            console.log(`   ⚠️ Serviço "${service}" não encontrado nas opções`);
                        }
                    } else {
                        console.log('   ⚠️ Nenhuma opção de serviço encontrada no dropdown');
                        
                        // Debug: verificar se dropdown abriu
                        const pageAfterClick = await page.evaluate(() => {
                            const elements = [];
                            document.querySelectorAll('[role="listbox"], [role="menu"], ul.dropdown, div.dropdown-menu').forEach(el => {
                                const rect = el.getBoundingClientRect();
                                elements.push({ tag: el.tagName, role: el.getAttribute('role'), y: rect.y, h: rect.height });
                            });
                            return elements;
                        });
                        console.log('   DEBUG dropdowns/menus:', pageAfterClick);
                    }
                } else {
                    console.log('   ⚠️ Botão "Filter by: All Services" não encontrado na página');
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
