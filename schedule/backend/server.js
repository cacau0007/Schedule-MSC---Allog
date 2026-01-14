// ============================================
// SERVIDOR DE SCHEDULES - MSC
// Versão 3.0 - Mapeamento COMPLETO de rotas
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
// MAPEAMENTO COMPLETO DE SERVIÇOS POR ROTA
// Baseado nas rotas oficiais MSC 2024/2025:
// 
// SANTANA: Yantian → Ningbo → Shanghai → Qingdao → Busan → Cristobal → Suape → Salvador
//          (Conecta Manaus/Pecem via feeder de Salvador)
//
// CARIOCA: Qingdao → Busan → Ningbo → Shanghai → Shekou → Singapore → Colombo → 
//          Rio de Janeiro → Santos → Paranagua → Navegantes → Imbituba → Itajai → Santos → Itaguai
//
// IPANEMA: Busan → Shanghai → Ningbo → Shekou → Yantian → Hong Kong → Singapore → 
//          Santos → Paranagua → Navegantes → Montevideo → Buenos Aires → Rio Grande
//
// TIGER/JADE/PHOENIX/DRAGON: Serviços Asia-Mediterrâneo que podem conectar ao Brasil via transbordo
// ============================================

// Lista de todos os POLs (Portos de Origem na Asia)
const ALL_POLS = [
    'Yantian', 'Ningbo', 'Shanghai', 'Qingdao', 'Busan', 'Shekou', 'Xiamen',
    'Hong Kong', 'Singapore', 'Kaohsiung', 'Dalian', 'Xingang', 'Nansha',
    'Tianjin', 'Qinzhou', 'Fuzhou', 'Lianyungang', 'Taicang', 'Zhanjiang',
    'Haiphong', 'Ho Chi Minh', 'Laem Chabang', 'Port Klang', 'Tanjung Pelepas',
    'Jakarta', 'Surabaya', 'Colombo', 'Mundra', 'Nhava Sheva', 'Chennai'
];

// Lista de todos os PODs (Portos de Destino no Brasil e região)
const ALL_PODS = [
    'Santos', 'Paranagua', 'Navegantes', 'Itajai', 'Imbituba', 'Rio Grande',
    'Rio de Janeiro', 'Itaguai', 'Vitoria', 'Salvador', 'Suape', 'Pecem',
    'Manaus', 'Belem', 'Fortaleza', 'Montevideo', 'Buenos Aires', 'Itapoa'
];

// ============================================
// MAPEAMENTO DE SERVIÇOS POR ROTA (POL → POD)
// ============================================
const SERVICE_ROUTES = {
    // =====================================================
    // SANTANA SERVICE - POLs: Yantian, Ningbo, Shanghai, Qingdao, Busan
    // PODs: Suape, Salvador, Pecem, Manaus (via feeder), Cristobal, Caucedo
    // =====================================================
    
    // Yantian → Brasil Norte/Nordeste (Santana)
    'Yantian-Suape': ['Santana'],
    'Yantian-Salvador': ['Santana'],
    'Yantian-Pecem': ['Santana'],
    'Yantian-Manaus': ['Santana'],
    'Yantian-Fortaleza': ['Santana'],
    
    // Ningbo → Brasil Norte/Nordeste (Santana)
    'Ningbo-Suape': ['Santana'],
    'Ningbo-Salvador': ['Santana'],
    'Ningbo-Pecem': ['Santana'],
    'Ningbo-Manaus': ['Santana'],
    'Ningbo-Fortaleza': ['Santana'],
    
    // Shanghai → Brasil Norte/Nordeste (Santana)
    'Shanghai-Suape': ['Santana'],
    'Shanghai-Salvador': ['Santana'],
    'Shanghai-Pecem': ['Santana'],
    'Shanghai-Manaus': ['Santana'],
    'Shanghai-Fortaleza': ['Santana'],
    
    // Qingdao → Brasil Norte/Nordeste (Santana)
    'Qingdao-Suape': ['Santana'],
    'Qingdao-Salvador': ['Santana'],
    'Qingdao-Pecem': ['Santana'],
    'Qingdao-Manaus': ['Santana'],
    'Qingdao-Fortaleza': ['Santana'],
    
    // Busan → Brasil Norte/Nordeste (Santana)
    'Busan-Suape': ['Santana'],
    'Busan-Salvador': ['Santana'],
    'Busan-Pecem': ['Santana'],
    'Busan-Manaus': ['Santana'],
    'Busan-Fortaleza': ['Santana'],
    
    // =====================================================
    // CARIOCA SERVICE - POLs: Qingdao, Busan, Ningbo, Shanghai, Shekou, Singapore
    // PODs: Rio de Janeiro, Santos, Paranagua, Navegantes, Imbituba, Itajai, Itaguai
    // =====================================================
    
    // Qingdao → Brasil Sul/Sudeste (Carioca)
    'Qingdao-Rio de Janeiro': ['Carioca'],
    'Qingdao-Santos': ['Carioca'],
    'Qingdao-Paranagua': ['Carioca'],
    'Qingdao-Navegantes': ['Carioca'],
    'Qingdao-Imbituba': ['Carioca'],
    'Qingdao-Itajai': ['Carioca'],
    'Qingdao-Itaguai': ['Carioca'],
    'Qingdao-Itapoa': ['Carioca'],
    
    // Busan → Brasil Sul/Sudeste (Carioca + Ipanema)
    'Busan-Rio de Janeiro': ['Carioca'],
    'Busan-Santos': ['Ipanema', 'Carioca'],
    'Busan-Paranagua': ['Ipanema', 'Carioca'],
    'Busan-Navegantes': ['Ipanema', 'Carioca'],
    'Busan-Imbituba': ['Carioca'],
    'Busan-Itajai': ['Carioca'],
    'Busan-Itaguai': ['Carioca'],
    'Busan-Itapoa': ['Carioca'],
    'Busan-Rio Grande': ['Ipanema'],
    'Busan-Montevideo': ['Ipanema'],
    'Busan-Buenos Aires': ['Ipanema'],
    
    // Ningbo → Brasil Sul/Sudeste (Carioca + Ipanema)
    'Ningbo-Rio de Janeiro': ['Carioca'],
    'Ningbo-Santos': ['Ipanema', 'Carioca'],
    'Ningbo-Paranagua': ['Ipanema', 'Carioca'],
    'Ningbo-Navegantes': ['Ipanema', 'Carioca'],
    'Ningbo-Imbituba': ['Carioca'],
    'Ningbo-Itajai': ['Carioca'],
    'Ningbo-Itaguai': ['Carioca'],
    'Ningbo-Itapoa': ['Carioca'],
    'Ningbo-Rio Grande': ['Ipanema'],
    'Ningbo-Montevideo': ['Ipanema'],
    'Ningbo-Buenos Aires': ['Ipanema'],
    
    // Shanghai → Brasil Sul/Sudeste (Carioca + Ipanema)
    'Shanghai-Rio de Janeiro': ['Carioca'],
    'Shanghai-Santos': ['Ipanema', 'Carioca'],
    'Shanghai-Paranagua': ['Ipanema', 'Carioca'],
    'Shanghai-Navegantes': ['Ipanema', 'Carioca'],
    'Shanghai-Imbituba': ['Carioca'],
    'Shanghai-Itajai': ['Carioca'],
    'Shanghai-Itaguai': ['Carioca'],
    'Shanghai-Itapoa': ['Carioca'],
    'Shanghai-Rio Grande': ['Ipanema'],
    'Shanghai-Montevideo': ['Ipanema'],
    'Shanghai-Buenos Aires': ['Ipanema'],
    
    // Shekou → Brasil Sul/Sudeste (Carioca + Ipanema)
    'Shekou-Rio de Janeiro': ['Carioca'],
    'Shekou-Santos': ['Ipanema', 'Carioca'],
    'Shekou-Paranagua': ['Ipanema', 'Carioca'],
    'Shekou-Navegantes': ['Ipanema', 'Carioca'],
    'Shekou-Imbituba': ['Carioca'],
    'Shekou-Itajai': ['Carioca'],
    'Shekou-Itaguai': ['Carioca'],
    'Shekou-Itapoa': ['Carioca'],
    'Shekou-Rio Grande': ['Ipanema'],
    'Shekou-Montevideo': ['Ipanema'],
    'Shekou-Buenos Aires': ['Ipanema'],
    
    // =====================================================
    // IPANEMA SERVICE - POLs: Busan, Shanghai, Ningbo, Shekou, Yantian, Hong Kong, Singapore
    // PODs: Santos, Paranagua, Navegantes, Montevideo, Buenos Aires, Rio Grande
    // =====================================================
    
    // Yantian → Brasil Sul + Argentina/Uruguai (Ipanema)
    'Yantian-Santos': ['Ipanema'],
    'Yantian-Paranagua': ['Ipanema'],
    'Yantian-Navegantes': ['Ipanema'],
    'Yantian-Rio Grande': ['Ipanema'],
    'Yantian-Montevideo': ['Ipanema'],
    'Yantian-Buenos Aires': ['Ipanema'],
    
    // Hong Kong → Brasil Sul + Argentina/Uruguai (Ipanema)
    'Hong Kong-Santos': ['Ipanema'],
    'Hong Kong-Paranagua': ['Ipanema'],
    'Hong Kong-Navegantes': ['Ipanema'],
    'Hong Kong-Rio Grande': ['Ipanema'],
    'Hong Kong-Montevideo': ['Ipanema'],
    'Hong Kong-Buenos Aires': ['Ipanema'],
    
    // Singapore → Brasil Sul/Sudeste (Ipanema + Carioca)
    'Singapore-Rio de Janeiro': ['Carioca'],
    'Singapore-Santos': ['Ipanema', 'Carioca'],
    'Singapore-Paranagua': ['Ipanema', 'Carioca'],
    'Singapore-Navegantes': ['Ipanema', 'Carioca'],
    'Singapore-Imbituba': ['Carioca'],
    'Singapore-Itajai': ['Carioca'],
    'Singapore-Itaguai': ['Carioca'],
    'Singapore-Rio Grande': ['Ipanema'],
    'Singapore-Montevideo': ['Ipanema'],
    'Singapore-Buenos Aires': ['Ipanema'],
    
    // =====================================================
    // XIAMEN - Pode ter Tiger/Jade via transbordo no Mediterrâneo
    // =====================================================
    'Xiamen-Santos': ['Ipanema', 'Carioca', 'Jade', 'Tiger'],
    'Xiamen-Paranagua': ['Ipanema', 'Carioca', 'Jade', 'Tiger'],
    'Xiamen-Navegantes': ['Ipanema', 'Carioca', 'Jade', 'Tiger'],
    'Xiamen-Itajai': ['Carioca', 'Jade'],
    'Xiamen-Imbituba': ['Carioca'],
    'Xiamen-Rio de Janeiro': ['Carioca'],
    'Xiamen-Suape': ['Santana'],
    'Xiamen-Salvador': ['Santana'],
    'Xiamen-Manaus': ['Santana'],
    
    // =====================================================
    // KAOHSIUNG - Tiger Service (via transbordo)
    // =====================================================
    'Kaohsiung-Santos': ['Tiger'],
    'Kaohsiung-Paranagua': ['Tiger'],
    'Kaohsiung-Navegantes': ['Tiger'],
    'Kaohsiung-Itajai': ['Tiger'],
    
    // =====================================================
    // DALIAN / XINGANG - Tiger Service (via transbordo)
    // =====================================================
    'Dalian-Santos': ['Tiger'],
    'Dalian-Paranagua': ['Tiger'],
    'Dalian-Navegantes': ['Tiger'],
    'Dalian-Itajai': ['Tiger'],
    
    'Xingang-Santos': ['Tiger'],
    'Xingang-Paranagua': ['Tiger'],
    'Xingang-Navegantes': ['Tiger'],
    'Xingang-Itajai': ['Tiger'],
    
    // =====================================================
    // NANSHA - Dragon/Lion Services (via transbordo)
    // =====================================================
    'Nansha-Santos': ['Dragon', 'Lion'],
    'Nansha-Paranagua': ['Dragon'],
    'Nansha-Navegantes': ['Dragon'],
    
    // =====================================================
    // COLOMBO - Hub de transbordo (Carioca passa por lá)
    // =====================================================
    'Colombo-Rio de Janeiro': ['Carioca'],
    'Colombo-Santos': ['Carioca'],
    'Colombo-Paranagua': ['Carioca'],
    'Colombo-Navegantes': ['Carioca'],
    'Colombo-Itajai': ['Carioca'],
    'Colombo-Itaguai': ['Carioca'],
    
    // =====================================================
    // PORTOS DO SUDESTE ASIÁTICO (via Singapore transbordo)
    // =====================================================
    
    // Laem Chabang (Tailândia)
    'Laem Chabang-Santos': ['Ipanema', 'Carioca'],
    'Laem Chabang-Paranagua': ['Ipanema', 'Carioca'],
    'Laem Chabang-Navegantes': ['Ipanema', 'Carioca'],
    
    // Ho Chi Minh / Haiphong (Vietnã)
    'Ho Chi Minh-Santos': ['Ipanema', 'Carioca'],
    'Ho Chi Minh-Paranagua': ['Ipanema', 'Carioca'],
    'Ho Chi Minh-Navegantes': ['Ipanema', 'Carioca'],
    
    'Haiphong-Santos': ['Ipanema', 'Carioca'],
    'Haiphong-Paranagua': ['Ipanema', 'Carioca'],
    'Haiphong-Navegantes': ['Ipanema', 'Carioca'],
    
    // Port Klang / Tanjung Pelepas (Malásia)
    'Port Klang-Santos': ['Ipanema', 'Carioca'],
    'Port Klang-Paranagua': ['Ipanema', 'Carioca'],
    'Port Klang-Navegantes': ['Ipanema', 'Carioca'],
    
    'Tanjung Pelepas-Santos': ['Ipanema', 'Carioca'],
    'Tanjung Pelepas-Paranagua': ['Ipanema', 'Carioca'],
    'Tanjung Pelepas-Navegantes': ['Ipanema', 'Carioca'],
    
    // Jakarta / Surabaya (Indonésia)
    'Jakarta-Santos': ['Ipanema', 'Carioca'],
    'Jakarta-Paranagua': ['Ipanema', 'Carioca'],
    'Jakarta-Navegantes': ['Ipanema', 'Carioca'],
    
    'Surabaya-Santos': ['Ipanema', 'Carioca'],
    'Surabaya-Paranagua': ['Ipanema', 'Carioca'],
    'Surabaya-Navegantes': ['Ipanema', 'Carioca'],
    
    // =====================================================
    // ÍNDIA (via Colombo transbordo)
    // =====================================================
    'Mundra-Santos': ['Carioca'],
    'Mundra-Paranagua': ['Carioca'],
    'Mundra-Navegantes': ['Carioca'],
    
    'Nhava Sheva-Santos': ['Carioca'],
    'Nhava Sheva-Paranagua': ['Carioca'],
    'Nhava Sheva-Navegantes': ['Carioca'],
    
    'Chennai-Santos': ['Carioca'],
    'Chennai-Paranagua': ['Carioca'],
    'Chennai-Navegantes': ['Carioca'],
    
    // =====================================================
    // PORTOS CHINESES SECUNDÁRIOS (via hubs principais)
    // =====================================================
    
    // Tianjin
    'Tianjin-Santos': ['Carioca', 'Tiger'],
    'Tianjin-Paranagua': ['Carioca'],
    'Tianjin-Navegantes': ['Carioca'],
    'Tianjin-Suape': ['Santana'],
    'Tianjin-Salvador': ['Santana'],
    
    // Fuzhou
    'Fuzhou-Santos': ['Ipanema', 'Carioca'],
    'Fuzhou-Paranagua': ['Ipanema', 'Carioca'],
    'Fuzhou-Navegantes': ['Ipanema', 'Carioca'],
    
    // Lianyungang
    'Lianyungang-Santos': ['Carioca'],
    'Lianyungang-Paranagua': ['Carioca'],
    'Lianyungang-Navegantes': ['Carioca'],
    
    // Taicang
    'Taicang-Santos': ['Ipanema', 'Carioca'],
    'Taicang-Paranagua': ['Ipanema', 'Carioca'],
    'Taicang-Navegantes': ['Ipanema', 'Carioca'],
    
    // Qinzhou
    'Qinzhou-Santos': ['Carioca'],
    'Qinzhou-Paranagua': ['Carioca'],
    'Qinzhou-Navegantes': ['Carioca'],
    
    // Zhanjiang
    'Zhanjiang-Santos': ['Carioca'],
    'Zhanjiang-Paranagua': ['Carioca'],
    'Zhanjiang-Navegantes': ['Carioca']
};

// ============================================
// FUNÇÃO: Obter serviços disponíveis para rota
// ============================================
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
        // Rota não mapeada - retorna ALL (permite qualquer serviço)
        return res.json({ 
            services: ['ALL'], 
            message: 'Todas as rotas disponíveis',
            allowAll: true
        });
    }
    
    // Rota mapeada - retorna serviços específicos + opção ALL
    return res.json({ 
        services: services,
        allowAll: true  // SEMPRE permitir "Todos os serviços" como opção
    });
});

// ============================================
// ENDPOINT: Listar todos os POLs disponíveis
// ============================================
app.get('/api/pols', (req, res) => {
    res.json({ pols: ALL_POLS });
});

// ============================================
// ENDPOINT: Listar todos os PODs disponíveis
// ============================================
app.get('/api/pods', (req, res) => {
    res.json({ pods: ALL_PODS });
});

// ============================================
// ENDPOINT: Buscar schedules
// ============================================
app.post('/api/schedules', async (req, res) => {
    const { pol, pod, carriers, service } = req.body;
    
    console.log('\n' + '='.repeat(50));
    console.log('🟡 MSC - ' + pol + ' → ' + pod);
    if (service && service !== 'ALL') {
        console.log('   Serviço: ' + service);
    }
    console.log('='.repeat(50));
    
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
        console.error('❌ Erro:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// ============================================
// SCRAPER: MSC
// ============================================
async function scrapeMSC(pol, pod, service = null) {
    const sailings = [];
    let browser;
    let page;
    
    try {
        console.log('1. Acessando MSC...');
        
        browser = await puppeteer.launch({
            headless: 'new',
            executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || process.env.CHROME_PATH || '/usr/bin/google-chrome-stable',
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--disable-web-security',
                '--disable-features=VizDisplayCompositor',
                '--window-size=1920,1080'
            ]
        });
        
        page = await browser.newPage();
        await page.setViewport({ width: 1920, height: 1080 });
        
        await page.goto('https://www.msc.com/en/search-a-schedule', {
            waitUntil: 'networkidle2',
            timeout: 60000
        });
        
        // Esperar página carregar completamente
        await new Promise(r => setTimeout(r, 5000));
        
        // Screenshot de debug
        await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-01-pagina-carregada.png') });
        
        // 2. Fechar popup de cookies
        console.log('2. Fechando popup de cookies...');
        try {
            const cookieBtn = await page.$('button#onetrust-accept-btn-handler');
            if (cookieBtn) {
                console.log('   Encontrado: button#onetrust-accept-btn-handler');
                await cookieBtn.click();
                await new Promise(r => setTimeout(r, 500));
            }
        } catch (e) {}
        
        // 3. Preencher POL
        console.log(`3. Preenchendo origem: ${pol}`);
        const polFilled = await fillInput(page, 'Port of loading', pol);
        if (!polFilled) {
            throw new Error('Não conseguiu preencher POL');
        }
        
        // 4. Preencher POD
        console.log(`4. Preenchendo destino: ${pod}`);
        const podFilled = await fillInput(page, 'Port of discharge', pod);
        if (!podFilled) {
            throw new Error('Não conseguiu preencher POD');
        }
        
        // 5. Clicar em Search
        console.log('5. Clicando no botão Search...');
        await clickSearchButton(page);
        
        // 6. Aguardar resultados
        console.log('6. Aguardando resultados (10s)...');
        await new Promise(r => setTimeout(r, 10000));
        
        // 6.5. Filtrar por serviço
        if (service && service !== 'ALL') {
            console.log(`6.5. Filtrando por serviço: ${service}...`);
            await filterByService(page, service);
        }
        
        // 7. Extrair dados
        console.log('7. Extraindo dados...');
        const data = await extractScheduleData(page, service);
        
        console.log(`   Encontrados: ${data.length} navios únicos`);
        
        data.forEach(r => {
            sailings.push({
                carrier: 'MSC',
                service: r.service || service || '-',
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
            try {
                await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-error.png') });
            } catch (e) {}
        }
    } finally {
        if (browser) {
            await browser.close();
        }
    }
    
    console.log(`✅ Total: ${sailings.length} schedules`);
    return sailings;
}

// ============================================
// Preenchimento rápido de input
// ============================================
async function fillInput(page, placeholder, value) {
    try {
        // Tentar múltiplos seletores
        const selectors = [
            `input[placeholder*="${placeholder}"]`,
            `input[placeholder*="${placeholder.toLowerCase()}"]`,
            `input[aria-label*="${placeholder}"]`,
            `input[name*="origin" i]`,
            `input[name*="destination" i]`,
            `input[id*="origin" i]`,
            `input[id*="destination" i]`
        ];
        
        let input = null;
        for (const sel of selectors) {
            try {
                await page.waitForSelector(sel, { timeout: 3000 });
                input = await page.$(sel);
                if (input) {
                    console.log(`   Encontrado: ${sel}`);
                    break;
                }
            } catch (e) {
                continue;
            }
        }
        
        // Se não encontrou por seletor, buscar por texto na página
        if (!input) {
            console.log(`   Buscando input por evaluate...`);
            input = await page.evaluateHandle((ph) => {
                const inputs = document.querySelectorAll('input');
                for (const inp of inputs) {
                    const placeholder = (inp.placeholder || '').toLowerCase();
                    const ariaLabel = (inp.getAttribute('aria-label') || '').toLowerCase();
                    if (placeholder.includes(ph.toLowerCase()) || ariaLabel.includes(ph.toLowerCase())) {
                        return inp;
                    }
                }
                // Tentar pegar inputs visíveis
                const allInputs = Array.from(inputs).filter(i => {
                    const rect = i.getBoundingClientRect();
                    return rect.width > 0 && rect.height > 0;
                });
                if (allInputs.length >= 2) {
                    // Primeiro input = origem, segundo = destino
                    return ph.includes('loading') ? allInputs[0] : allInputs[1];
                }
                return null;
            }, placeholder);
        }
        
        if (!input) {
            console.log(`   ⚠️ Input não encontrado para: ${placeholder}`);
            return false;
        }
        
        await input.click();
        await new Promise(r => setTimeout(r, 500));
        
        // Limpar e preencher
        await page.evaluate((el, text) => {
            if (el && el.nodeName === 'INPUT') {
                el.value = '';
                el.value = text;
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true }));
            }
        }, input, value);
        
        await new Promise(r => setTimeout(r, 2000));
        
        await page.keyboard.press('ArrowDown');
        await new Promise(r => setTimeout(r, 300));
        await page.keyboard.press('Enter');
        
        console.log(`   ✅ ${placeholder} preenchido`);
        return true;
        
    } catch (e) {
        console.log(`   ⚠️ Erro ao preencher ${placeholder}: ${e.message}`);
        return false;
    }
}

// ============================================
// Clicar no botão Search
// ============================================
async function clickSearchButton(page) {
    try {
        const clicked = await page.evaluate(() => {
            const buttons = Array.from(document.querySelectorAll('button, a, [role="button"]'));
            for (const btn of buttons) {
                const text = (btn.innerText || btn.textContent || '').trim().toUpperCase();
                if (text === 'SEARCH' || text === 'BUSCAR') {
                    btn.click();
                    return true;
                }
            }
            return false;
        });
        
        if (clicked) {
            console.log('   ✅ Clicado!');
            return;
        }
        
        const submitBtn = await page.$('button[type="submit"]');
        if (submitBtn) {
            await submitBtn.click();
            console.log('   ✅ Clicado (submit)!');
            return;
        }
        
        console.log('   ⚠️ Botão Search não encontrado');
        
    } catch (e) {
        console.log(`   ⚠️ Erro ao clicar Search: ${e.message}`);
    }
}

// ============================================
// Filtro robusto por serviço
// ============================================
async function filterByService(page, serviceName) {
    try {
        await new Promise(r => setTimeout(r, 2000));
        
        // Buscar dropdown por texto
        const filterOpened = await page.evaluate(() => {
            const allElements = document.querySelectorAll('*');
            
            for (const el of allElements) {
                const text = (el.innerText || '').trim();
                
                if ((text.includes('Filter by') || text === 'All Services' || text.includes('All Services')) 
                    && !text.includes('\n')
                    && text.length < 50) {
                    
                    const rect = el.getBoundingClientRect();
                    if (rect.width > 50 && rect.height > 10 && rect.width < 400) {
                        el.click();
                        return { success: true, text: text };
                    }
                }
            }
            
            const selects = document.querySelectorAll('select, [role="listbox"], [role="combobox"]');
            for (const sel of selects) {
                const rect = sel.getBoundingClientRect();
                if (rect.y > 300 && rect.width > 100) {
                    sel.click();
                    return { success: true, text: 'select/dropdown' };
                }
            }
            
            return { success: false };
        });
        
        if (filterOpened.success) {
            console.log(`   ✅ Filtro aberto: ${filterOpened.text}`);
            
            await new Promise(r => setTimeout(r, 2000));
            
            await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-06-filtro-aberto.png') });
            
            // Selecionar o serviço
            const serviceSelected = await page.evaluate((targetService) => {
                const searchTexts = [
                    targetService + ' Service',
                    targetService,
                    targetService.toUpperCase() + ' SERVICE',
                    targetService.toLowerCase() + ' service'
                ];
                
                const allElements = document.querySelectorAll('*');
                let bestMatch = null;
                let smallestArea = Infinity;
                
                for (const el of allElements) {
                    const text = (el.innerText || '').trim();
                    const rect = el.getBoundingClientRect();
                    
                    const matches = searchTexts.some(search => 
                        text === search || 
                        text.toLowerCase() === search.toLowerCase()
                    );
                    
                    if (matches && rect.y > 0 && rect.width > 0 && rect.height > 0 && rect.width < 300) {
                        const area = rect.width * rect.height;
                        if (area < smallestArea && area > 100) {
                            smallestArea = area;
                            bestMatch = el;
                        }
                    }
                }
                
                if (bestMatch) {
                    bestMatch.click();
                    return { success: true, text: bestMatch.innerText.trim() };
                }
                
                return { success: false };
            }, serviceName);
            
            if (serviceSelected.success) {
                console.log(`   ✅ ${serviceSelected.text} selecionado!`);
                await new Promise(r => setTimeout(r, 2000));
            } else {
                console.log(`   ⚠️ ${serviceName} Service não encontrado no dropdown`);
            }
            
        } else {
            console.log('   ⚠️ Filtro não encontrado');
        }
        
        await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-07-filtro-aplicado.png') });
        
    } catch (e) {
        console.log(`   ⚠️ Erro ao filtrar: ${e.message}`);
    }
}

// ============================================
// Extração com duplicatas inteligentes
// ============================================
async function extractScheduleData(page, filterService) {
    const data = await page.evaluate((serviceName) => {
        const seenVessels = new Map();
        const text = document.body.innerText;
        const lines = text.split('\n');
        
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
            const vesselMatch = line.match(/^([A-Z][A-Z\s\-]+)\s*(?:\/\s*[A-Z0-9]+W?)?$/i);
            if (vesselMatch && line.length > 5 && line.length < 50) {
                const excluded = ['DEPARTURE', 'ARRIVAL', 'VESSEL', 'VOYAGE', 'DIRECT', 
                                 'TRANSHIPMENT', 'FILTER', 'RESULTS', 'POINT', 'SERVICES',
                                 'PORT OF', 'SEARCH', 'ALL SERVICES', 'MONDAY', 'TUESDAY',
                                 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY'];
                const possibleVessel = (vesselMatch[1] || line).replace(/\s*\/.*/, '').trim();
                
                if (!excluded.some(ex => possibleVessel.toUpperCase().includes(ex)) && possibleVessel.length > 3) {
                    currentVessel = possibleVessel;
                }
                continue;
            }
            
            // Detectar transit time
            const transitMatch = line.match(/^(\d+)\s*Days?$/i);
            if (transitMatch) {
                currentTransit = parseInt(transitMatch[1]);
                continue;
            }
            
            // Detectar tipo de rota
            if (line === 'Direct' || line === 'Transhipment') {
                currentRouting = line === 'Transhipment' ? 'Transbordo' : line;
                
                if (currentVessel && currentDeparture) {
                    const existingEntry = seenVessels.get(currentVessel);
                    
                    if (!existingEntry) {
                        seenVessels.set(currentVessel, {
                            service: serviceName || '-',
                            vessel: currentVessel,
                            etd: currentDeparture || '-',
                            eta: currentArrival || '-',
                            transit: currentTransit || 0,
                            routeType: currentRouting || '-'
                        });
                    } else {
                        // Duplicado - manter o com MAIOR transit time
                        if ((currentTransit || 0) > existingEntry.transit) {
                            seenVessels.set(currentVessel, {
                                service: serviceName || '-',
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
        
        return Array.from(seenVessels.values()).map(entry => ({
            ...entry,
            transit: entry.transit ? `${entry.transit} dias` : '-'
        }));
        
    }, filterService && filterService !== 'ALL' ? filterService : null);
    
    return data;
}

// ============================================
// Excel com formato de data melhorado
// ============================================
async function generateExcel(sailings, pol, pod, filename) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('SCHEDULES');
    
    // Cabeçalho principal
    sheet.mergeCells('A1:G1');
    sheet.getCell('A1').value = `ALLOG - Shipping Schedules: ${pol} → ${pod}`;
    sheet.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFF' } };
    sheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E3A5F' } };
    sheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
    sheet.getRow(1).height = 30;
    
    // Cabeçalhos das colunas
    const headers = ['CARRIER', 'SERVIÇO', 'NAVIO', 'ETD', 'ETA', 'TRANSIT', 'TIPO'];
    const headerRow = sheet.addRow(headers);
    headerRow.font = { bold: true, color: { argb: 'FFFFFF' } };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '2E7D32' } };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
    headerRow.height = 25;
    
    // Dados
    sailings.forEach((s, index) => {
        const etdFormatted = formatDate(s.etd);
        const etaFormatted = formatDate(s.eta);
        
        const row = sheet.addRow([
            s.carrier,
            s.service,
            s.vessel,
            etdFormatted,
            etaFormatted,
            s.transit || '-',
            s.routeType || '-'
        ]);
        
        const bgColor = index % 2 === 0 ? 'F5F5F5' : 'FFFFFF';
        row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
        row.alignment = { vertical: 'middle' };
    });
    
    // Largura das colunas
    sheet.columns = [
        { width: 10 },
        { width: 12 },
        { width: 28 },
        { width: 18 },
        { width: 18 },
        { width: 12 },
        { width: 12 }
    ];
    
    // Bordas
    sheet.eachRow((row) => {
        row.eachCell((cell) => {
            cell.border = {
                top: { style: 'thin', color: { argb: 'CCCCCC' } },
                left: { style: 'thin', color: { argb: 'CCCCCC' } },
                bottom: { style: 'thin', color: { argb: 'CCCCCC' } },
                right: { style: 'thin', color: { argb: 'CCCCCC' } }
            };
        });
    });
    
    const filepath = path.join(EXPORTS_DIR, filename);
    await workbook.xlsx.writeFile(filepath);
    console.log(`📊 Excel gerado: ${filename}`);
}

// ============================================
// Função para formatar data
// ============================================
function formatDate(dateStr) {
    if (!dateStr || dateStr === '-') return '-';
    
    try {
        const match = dateStr.match(/^(\w+)\s+(\d{1,2})(?:st|nd|rd|th)?\s+(\w+)\s+(\d{4})$/i);
        if (!match) return dateStr;
        
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
    console.log('\n' + '='.repeat(50));
    console.log('🚀 SERVIDOR DE SCHEDULES v3.0');
    console.log('   Mapeamento COMPLETO de rotas MSC');
    console.log('='.repeat(50));
    console.log(`📍 Porta: ${PORT}`);
    console.log(`📁 Exports: ${EXPORTS_DIR}`);
    console.log(`🗺️  Rotas mapeadas: ${Object.keys(SERVICE_ROUTES).length}`);
    console.log('='.repeat(50) + '\n');
});
