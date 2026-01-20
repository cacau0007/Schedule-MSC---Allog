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
// Baseado na Matriz MSC Network 2025/2026
// Fonte: Documento oficial MSC Ásia ↔ ECSA
// ============================================

// Lista dos 4 serviços MSC Asia-Brasil que usamos
const ALL_SERVICES = ['Carioca', 'Ipanema', 'Santana', 'Jade'];

// Abreviações usadas no documento:
// Rio = Rio de Janeiro, Pgua = Paranaguá, Nav = Navegantes
// Mvd = Montevideo, B.A = Buenos Aires, RG = Rio Grande

const SERVICE_ROUTES = {
    // ========================================================
    // CARIOCA - Santos/Rio/Paranaguá/Itapoá/Itaguaí
    // Rota: Busan-Shanghai-Ningbo-Shekou-Singapore-Rio-Paranaguá-Itapoá-Santos-Itaguaí
    // ========================================================
    
    // POLs Diretos do Carioca
    'Shanghai-Santos': ['Carioca', 'Ipanema', 'Santana'],
    'Shanghai-Rio de Janeiro': ['Carioca', 'Santana'],
    'Shanghai-Paranagua': ['Carioca', 'Ipanema', 'Santana'],
    'Shanghai-Itapoa': ['Carioca'],
    'Shanghai-Itaguai': ['Carioca', 'Santana'],
    
    'Ningbo-Santos': ['Carioca', 'Ipanema', 'Santana'],
    'Ningbo-Rio de Janeiro': ['Carioca', 'Santana'],
    'Ningbo-Paranagua': ['Carioca', 'Ipanema', 'Santana'],
    'Ningbo-Itapoa': ['Carioca'],
    'Ningbo-Itaguai': ['Carioca', 'Santana'],
    
    'Shekou-Santos': ['Carioca', 'Ipanema', 'Santana'],
    'Shekou-Rio de Janeiro': ['Carioca', 'Santana'],
    'Shekou-Paranagua': ['Carioca', 'Ipanema', 'Santana'],
    'Shekou-Itapoa': ['Carioca'],
    'Shekou-Itaguai': ['Carioca', 'Santana'],
    
    'Busan-Santos': ['Carioca', 'Ipanema', 'Santana'],
    'Busan-Rio de Janeiro': ['Carioca', 'Santana'],
    'Busan-Paranagua': ['Carioca', 'Ipanema', 'Santana'],
    'Busan-Itapoa': ['Carioca'],
    'Busan-Itaguai': ['Carioca', 'Santana'],
    
    'Singapore-Santos': ['Carioca', 'Ipanema', 'Santana'],
    'Singapore-Rio de Janeiro': ['Carioca', 'Santana'],
    'Singapore-Paranagua': ['Carioca', 'Ipanema', 'Santana'],
    'Singapore-Itapoa': ['Carioca'],
    'Singapore-Itaguai': ['Carioca', 'Santana'],
    
    'Qingdao-Santos': ['Carioca', 'Santana'],
    'Qingdao-Rio de Janeiro': ['Carioca', 'Santana'],
    'Qingdao-Paranagua': ['Carioca', 'Santana'],
    'Qingdao-Itapoa': ['Carioca'],
    'Qingdao-Itaguai': ['Carioca', 'Santana'],
    
    // ========================================================
    // IPANEMA - Santos/Paranaguá/Navegantes/Montevideo/Buenos Aires/Rio Grande
    // Rota: Busan-Shanghai-Ningbo-Shekou-HongKong-Singapore-Santos-Paranaguá-Navegantes-Montevideo-BA-RioGrande
    // ========================================================
    
    'Shanghai-Navegantes': ['Ipanema', 'Santana'],
    'Shanghai-Montevideo': ['Ipanema'],
    'Shanghai-Buenos Aires': ['Ipanema'],
    'Shanghai-Rio Grande': ['Ipanema'],
    
    'Ningbo-Navegantes': ['Ipanema', 'Santana'],
    'Ningbo-Montevideo': ['Ipanema'],
    'Ningbo-Buenos Aires': ['Ipanema'],
    'Ningbo-Rio Grande': ['Ipanema'],
    
    'Shekou-Navegantes': ['Ipanema', 'Santana'],
    'Shekou-Montevideo': ['Ipanema'],
    'Shekou-Buenos Aires': ['Ipanema'],
    'Shekou-Rio Grande': ['Ipanema'],
    
    'Yantian-Santos': ['Ipanema'],
    'Yantian-Paranagua': ['Ipanema'],
    'Yantian-Navegantes': ['Ipanema'],
    'Yantian-Montevideo': ['Ipanema'],
    'Yantian-Buenos Aires': ['Ipanema'],
    'Yantian-Rio Grande': ['Ipanema'],
    
    'Hong Kong-Santos': ['Ipanema'],
    'Hong Kong-Paranagua': ['Ipanema'],
    'Hong Kong-Navegantes': ['Ipanema'],
    'Hong Kong-Montevideo': ['Ipanema'],
    'Hong Kong-Buenos Aires': ['Ipanema'],
    'Hong Kong-Rio Grande': ['Ipanema'],
    
    'Singapore-Navegantes': ['Ipanema', 'Santana'],
    'Singapore-Montevideo': ['Ipanema'],
    'Singapore-Buenos Aires': ['Ipanema'],
    'Singapore-Rio Grande': ['Ipanema'],
    
    // Busan também tem Ipanema para Navegantes e Sul
    'Busan-Navegantes': ['Ipanema', 'Santana'],
    'Busan-Montevideo': ['Ipanema'],
    'Busan-Buenos Aires': ['Ipanema'],
    'Busan-Rio Grande': ['Ipanema'],
    
    // ========================================================
    // SANTANA - Rio/Santos/Paranaguá/Navegantes/Imbituba/Itajaí/Itaguaí
    // Rota: Qingdao-Busan-Ningbo-Shanghai-Shekou-Singapore-Colombo-Rio-Santos-Paranaguá-Navegantes-Imbituba-Itajaí-Itaguaí
    // ========================================================
    
    'Qingdao-Navegantes': ['Santana'],
    'Qingdao-Imbituba': ['Santana'],
    'Qingdao-Itajai': ['Santana'],
    
    'Busan-Imbituba': ['Santana'],
    'Busan-Itajai': ['Santana'],
    
    'Shanghai-Imbituba': ['Santana'],
    'Shanghai-Itajai': ['Santana'],
    
    'Ningbo-Imbituba': ['Santana'],
    'Ningbo-Itajai': ['Santana'],
    
    'Shekou-Imbituba': ['Santana'],
    'Shekou-Itajai': ['Santana'],
    
    'Singapore-Imbituba': ['Santana'],
    'Singapore-Itajai': ['Santana'],
    
    // ========================================================
    // JADE - Via Transbordo (Sines/Valência)
    // Apenas para: Xiamen, Nansha, Fuzhou, Kaohsiung
    // ========================================================
    
    'Xiamen-Santos': ['Jade'],
    'Xiamen-Navegantes': ['Jade'],
    'Xiamen-Paranagua': ['Jade'],
    'Xiamen-Rio de Janeiro': ['Jade'],
    
    'Nansha-Santos': ['Jade'],
    'Nansha-Navegantes': ['Jade'],
    'Nansha-Paranagua': ['Jade'],
    'Nansha-Rio de Janeiro': ['Jade'],
    
    'Fuzhou-Santos': ['Jade'],
    'Fuzhou-Navegantes': ['Jade'],
    'Fuzhou-Paranagua': ['Jade'],
    'Fuzhou-Rio de Janeiro': ['Jade'],
    
    'Kaohsiung-Santos': ['Jade'],
    'Kaohsiung-Navegantes': ['Jade'],
    'Kaohsiung-Paranagua': ['Jade'],
    'Kaohsiung-Rio de Janeiro': ['Jade'],
    
    // Jade também pode ir para Suape/Salvador via T/S Sines
    'Shanghai-Suape': ['Jade'],
    'Shanghai-Salvador': ['Jade'],
    'Ningbo-Suape': ['Jade'],
    'Ningbo-Salvador': ['Jade'],
    
    // ========================================================
    // CONEXÕES VIA HUB (Singapore/Busan)
    // Outports da Indonésia, Vietnã, Tailândia, Malásia, Japão
    // ========================================================
    
    // Jakarta (via Singapore) - tem Carioca, Ipanema e Santana
    'Jakarta-Santos': ['Carioca', 'Ipanema', 'Santana'],
    'Jakarta-Rio de Janeiro': ['Carioca', 'Santana'],
    'Jakarta-Paranagua': ['Carioca', 'Ipanema', 'Santana'],
    'Jakarta-Navegantes': ['Ipanema', 'Santana'],
    'Jakarta-Itapoa': ['Carioca'],
    'Jakarta-Itaguai': ['Carioca', 'Santana'],
    
    // Surabaya (via Singapore)
    'Surabaya-Santos': ['Carioca', 'Ipanema', 'Santana'],
    'Surabaya-Rio de Janeiro': ['Carioca', 'Santana'],
    'Surabaya-Paranagua': ['Carioca', 'Ipanema', 'Santana'],
    'Surabaya-Navegantes': ['Ipanema', 'Santana'],
    
    // Laem Chabang (via Singapore)
    'Laem Chabang-Santos': ['Carioca', 'Ipanema', 'Santana'],
    'Laem Chabang-Rio de Janeiro': ['Carioca', 'Santana'],
    'Laem Chabang-Paranagua': ['Carioca', 'Ipanema', 'Santana'],
    'Laem Chabang-Navegantes': ['Ipanema', 'Santana'],
    
    // Ho Chi Minh (via Singapore)
    'Ho Chi Minh-Santos': ['Carioca', 'Ipanema', 'Santana'],
    'Ho Chi Minh-Rio de Janeiro': ['Carioca', 'Santana'],
    'Ho Chi Minh-Paranagua': ['Carioca', 'Ipanema', 'Santana'],
    'Ho Chi Minh-Navegantes': ['Ipanema', 'Santana'],
    
    // Haiphong (via Singapore)
    'Haiphong-Santos': ['Carioca', 'Ipanema', 'Santana'],
    'Haiphong-Rio de Janeiro': ['Carioca', 'Santana'],
    'Haiphong-Paranagua': ['Carioca', 'Ipanema', 'Santana'],
    'Haiphong-Navegantes': ['Ipanema', 'Santana'],
    
    // Port Klang (via Singapore)
    'Port Klang-Santos': ['Carioca', 'Ipanema', 'Santana'],
    'Port Klang-Rio de Janeiro': ['Carioca', 'Santana'],
    'Port Klang-Paranagua': ['Carioca', 'Ipanema', 'Santana'],
    'Port Klang-Navegantes': ['Ipanema', 'Santana'],
    
    // Xingang/Tianjin (via Busan)
    'Tianjin-Santos': ['Carioca', 'Ipanema', 'Santana'],
    'Tianjin-Rio de Janeiro': ['Carioca', 'Santana'],
    'Tianjin-Paranagua': ['Carioca', 'Ipanema', 'Santana'],
    'Tianjin-Navegantes': ['Ipanema', 'Santana'],
    'Xingang-Santos': ['Carioca', 'Ipanema', 'Santana'],
    'Xingang-Rio de Janeiro': ['Carioca', 'Santana'],
    'Xingang-Paranagua': ['Carioca', 'Ipanema', 'Santana'],
    'Xingang-Navegantes': ['Ipanema', 'Santana'],
    
    // Dalian (via Busan)
    'Dalian-Santos': ['Carioca', 'Ipanema', 'Santana'],
    'Dalian-Rio de Janeiro': ['Carioca', 'Santana'],
    'Dalian-Paranagua': ['Carioca', 'Ipanema', 'Santana'],
    'Dalian-Navegantes': ['Ipanema', 'Santana'],
    
    // Yokohama (via Busan)
    'Yokohama-Santos': ['Carioca', 'Ipanema'],
    'Yokohama-Rio de Janeiro': ['Carioca'],
    'Yokohama-Paranagua': ['Carioca', 'Ipanema'],
    'Yokohama-Navegantes': ['Ipanema'],
    
    // Tokyo (via Busan)
    'Tokyo-Santos': ['Carioca', 'Ipanema'],
    'Tokyo-Rio de Janeiro': ['Carioca'],
    'Tokyo-Paranagua': ['Carioca', 'Ipanema'],
    'Tokyo-Navegantes': ['Ipanema']
};

// Função para obter serviços disponíveis para uma rota
// Retorna null se a rota não estiver mapeada (não filtrar)
function getAvailableServices(pol, pod) {
    const routeKey = `${pol}-${pod}`;
    return SERVICE_ROUTES[routeKey] || null; // Se não mapeado, retorna null (não filtrar)
}

// ============================================
// ENDPOINT: Obter serviços disponíveis
// ============================================
app.get('/api/available-services', (req, res) => {
    const { pol, pod } = req.query;
    
    if (!pol || !pod) {
        return res.json({ services: ALL_SERVICES, message: 'Selecione POL e POD' });
    }
    
    const services = getAvailableServices(pol, pod);
    const routeKey = `${pol}-${pod}`;
    const isMapped = SERVICE_ROUTES.hasOwnProperty(routeKey);
    
    if (!isMapped) {
        // Rota não mapeada - não mostrar filtro de serviço
        return res.json({ 
            services: null,
            mapped: false,
            message: `Rota ${pol}-${pod} não mapeada - busca sem filtro de serviço`
        });
    }
    
    return res.json({ 
        services: services,
        mapped: true,
        message: `Serviços disponíveis para ${pol}-${pod}: ${services.join(', ')}`
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
    
    // Verificar se a rota está mapeada e se o serviço é válido
    const availableServices = getAvailableServices(pol, pod);
    let shouldFilter = false;
    let filterService = null;
    
    if (service && service !== 'ALL') {
        if (availableServices === null) {
            // Rota não mapeada - não filtrar
            console.log(`Rota não mapeada - buscando sem filtro de serviço`);
            shouldFilter = false;
        } else if (availableServices.map(s => s.toLowerCase()).includes(service.toLowerCase())) {
            // Serviço válido para esta rota - filtrar
            console.log(`Serviço solicitado: ${service} (válido para esta rota)`);
            shouldFilter = true;
            filterService = service;
        } else {
            // Serviço não disponível para esta rota - não filtrar
            console.log(`⚠️ Serviço "${service}" não disponível para ${pol}-${pod}`);
            console.log(`   Serviços disponíveis: ${availableServices.join(', ')}`);
            console.log(`   Buscando sem filtro de serviço`);
            shouldFilter = false;
        }
    } else {
        console.log(`Buscando todos os serviços`);
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
        
        // 6. Filtrar por serviço (SE shouldFilter = true)
        if (shouldFilter && filterService) {
            console.log(`6. Filtrando por serviço: ${filterService}...`);
            
            // Nome do serviço com "Service" (como aparece no site)
            const serviceWithSuffix = filterService + ' Service';
            console.log(`   Buscando: "${serviceWithSuffix}"`);
            
            try {
                // ============================================
                // ESTRATÉGIA 1: Encontrar e clicar no dropdown
                // ============================================
                
                // Primeiro, encontrar o elemento clicável do filtro
                const filterBtn = await page.evaluateHandle(() => {
                    // Buscar por texto "Filter by" ou "All Services"
                    const allElements = Array.from(document.querySelectorAll('*'));
                    
                    for (const el of allElements) {
                        const text = (el.innerText || '').trim();
                        
                        // Deve conter "Filter by" e "Service"
                        if (text.includes('Filter by') && text.includes('Service')) {
                            const rect = el.getBoundingClientRect();
                            // Tamanho de botão típico
                            if (rect.width > 100 && rect.width < 350 && 
                                rect.height > 20 && rect.height < 70 &&
                                rect.y > 300 && rect.y < 700) {
                                return el;
                            }
                        }
                    }
                    return null;
                });
                
                const filterElement = filterBtn.asElement();
                
                if (filterElement) {
                    // Obter posição do botão
                    const btnBox = await filterElement.boundingBox();
                    console.log(`   📍 Botão encontrado em x=${Math.round(btnBox.x)}, y=${Math.round(btnBox.y)}`);
                    
                    // ============================================
                    // MÚLTIPLAS TENTATIVAS DE ABRIR O DROPDOWN
                    // ============================================
                    
                    let dropdownOpened = false;
                    
                    // Tentativa 1: Clicar diretamente no elemento
                    console.log('   🔄 Tentativa 1: Clique direto no elemento...');
                    await filterElement.click();
                    await new Promise(r => setTimeout(r, 1500));
                    
                    // Verificar se abriu
                    let optionsCount = await page.evaluate((svc) => {
                        const elements = document.querySelectorAll('*');
                        let count = 0;
                        elements.forEach(el => {
                            const text = (el.innerText || '').trim().toLowerCase();
                            if (text === svc.toLowerCase() || text === (svc + ' service').toLowerCase()) {
                                const rect = el.getBoundingClientRect();
                                if (rect.height > 10 && rect.height < 60 && rect.y > 400) count++;
                            }
                        });
                        return count;
                    }, filterService);
                    
                    if (optionsCount === 0) {
                        // Tentativa 2: Clicar com mouse.click nas coordenadas
                        console.log('   🔄 Tentativa 2: Clique por coordenadas...');
                        await page.mouse.click(btnBox.x + btnBox.width / 2, btnBox.y + btnBox.height / 2);
                        await new Promise(r => setTimeout(r, 1500));
                    }
                    
                    // Tentativa 3: Clicar no lado direito (onde geralmente fica a seta)
                    optionsCount = await page.evaluate((svc) => {
                        const elements = document.querySelectorAll('*');
                        let count = 0;
                        elements.forEach(el => {
                            const text = (el.innerText || '').trim().toLowerCase();
                            if (text === svc.toLowerCase() || text === (svc + ' service').toLowerCase()) {
                                const rect = el.getBoundingClientRect();
                                if (rect.height > 10 && rect.height < 60 && rect.y > 400) count++;
                            }
                        });
                        return count;
                    }, filterService);
                    
                    if (optionsCount === 0) {
                        console.log('   🔄 Tentativa 3: Clique na seta (lado direito)...');
                        await page.mouse.click(btnBox.x + btnBox.width - 15, btnBox.y + btnBox.height / 2);
                        await new Promise(r => setTimeout(r, 1500));
                    }
                    
                    // Tentativa 4: Duplo clique
                    optionsCount = await page.evaluate((svc) => {
                        const elements = document.querySelectorAll('*');
                        let count = 0;
                        elements.forEach(el => {
                            const text = (el.innerText || '').trim().toLowerCase();
                            if (text === svc.toLowerCase() || text === (svc + ' service').toLowerCase()) {
                                const rect = el.getBoundingClientRect();
                                if (rect.height > 10 && rect.height < 60 && rect.y > 400) count++;
                            }
                        });
                        return count;
                    }, filterService);
                    
                    if (optionsCount === 0) {
                        console.log('   🔄 Tentativa 4: Duplo clique...');
                        await page.mouse.click(btnBox.x + btnBox.width / 2, btnBox.y + btnBox.height / 2, { clickCount: 2 });
                        await new Promise(r => setTimeout(r, 1500));
                    }
                    
                    await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-06-dropdown.png') });
                    
                    // ============================================
                    // BUSCAR E CLICAR NA OPÇÃO DO SERVIÇO
                    // ============================================
                    
                    console.log(`   🔍 Buscando opção "${serviceWithSuffix}"...`);
                    
                    // Buscar o elemento do serviço desejado
                    const serviceOption = await page.evaluateHandle((svcName, svcWithSuffix) => {
                        const elements = Array.from(document.querySelectorAll('*'));
                        const targetTexts = [
                            svcName.toLowerCase(),
                            svcWithSuffix.toLowerCase(),
                            svcName.toLowerCase() + ' service'
                        ];
                        
                        // Buscar elemento que contenha o nome do serviço
                        for (const el of elements) {
                            const text = (el.innerText || '').trim().toLowerCase();
                            const rect = el.getBoundingClientRect();
                            
                            // Deve ser um item de menu (tamanho apropriado)
                            if (rect.height < 15 || rect.height > 55) continue;
                            if (rect.width < 80 || rect.width > 350) continue;
                            if (rect.y < 450) continue; // Deve estar abaixo do botão
                            
                            // Verificar se é o serviço que queremos
                            for (const target of targetTexts) {
                                if (text === target || text.includes(target)) {
                                    // Verificar se não é o botão (que também contém o texto)
                                    if (!text.includes('filter by')) {
                                        return el;
                                    }
                                }
                            }
                        }
                        return null;
                    }, filterService, serviceWithSuffix);
                    
                    const serviceElement = serviceOption.asElement();
                    
                    if (serviceElement) {
                        const svcBox = await serviceElement.boundingBox();
                        console.log(`   ✅ Opção encontrada em y=${Math.round(svcBox.y)}`);
                        
                        // Clicar na opção
                        await serviceElement.click();
                        console.log(`   ✅ Serviço "${filterService}" selecionado!`);
                        
                        await new Promise(r => setTimeout(r, 2000));
                        dropdownOpened = true;
                    } else {
                        // ============================================
                        // FALLBACK: Listar todas as opções disponíveis
                        // ============================================
                        console.log('   ⚠️ Opção não encontrada diretamente. Listando opções disponíveis...');
                        
                        const availableOptions = await page.evaluate(() => {
                            const options = [];
                            const knownServices = ['santana', 'carioca', 'ipanema', 'jade', 'tiger', 'dragon', 'lion', 'all services'];
                            
                            document.querySelectorAll('*').forEach(el => {
                                const text = (el.innerText || '').trim();
                                const textLower = text.toLowerCase();
                                const rect = el.getBoundingClientRect();
                                
                                // Filtrar por tamanho e posição
                                if (rect.height < 15 || rect.height > 55) return;
                                if (rect.y < 400 || rect.y > 900) return;
                                
                                // Verificar se contém nome de serviço conhecido
                                for (const svc of knownServices) {
                                    if (textLower.includes(svc) && !textLower.includes('filter')) {
                                        if (!options.some(o => o.text === text)) {
                                            options.push({
                                                text: text,
                                                y: Math.round(rect.y),
                                                x: Math.round(rect.x + rect.width / 2),
                                                centerY: Math.round(rect.y + rect.height / 2)
                                            });
                                        }
                                        break;
                                    }
                                }
                            });
                            
                            return options.sort((a, b) => a.y - b.y);
                        });
                        
                        console.log(`   📋 Opções disponíveis (${availableOptions.length}):`);
                        availableOptions.forEach((opt, i) => console.log(`      [${i}] "${opt.text}" (y=${opt.y})`));
                        
                        // Tentar encontrar match e clicar
                        const targetLower = filterService.toLowerCase();
                        const match = availableOptions.find(opt => 
                            opt.text.toLowerCase().includes(targetLower)
                        );
                        
                        if (match) {
                            console.log(`   🎯 Match encontrado: "${match.text}"`);
                            await page.mouse.click(match.x, match.centerY);
                            console.log(`   ✅ Clicado em (${match.x}, ${match.centerY})`);
                            await new Promise(r => setTimeout(r, 2000));
                            dropdownOpened = true;
                        } else {
                            console.log(`   ⚠️ Serviço "${filterService}" não encontrado nas opções`);
                        }
                    }
                    
                    await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-07-filtro-aplicado.png') });
                    
                } else {
                    console.log('   ⚠️ Botão "Filter by" não encontrado na página');
                }
                
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
