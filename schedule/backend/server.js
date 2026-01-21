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
// APENAS: Carioca, Ipanema, Santana, Jade
// ============================================

const ALL_SERVICES = ['Carioca', 'Ipanema', 'Santana', 'Jade'];

// ============================================
// LEGENDA:
// - Rotas DIRETAS: têm serviços específicos, filtrar normalmente
// - Rotas CONEXÃO: fazem transbordo, NÃO filtrar (buscar todos)
// ============================================

// POLs que são CONEXÃO (via Singapore ou Busan) - NÃO filtrar
const CONNECTION_POLS = [
    'Jakarta', 'Surabaya', 'Panjang', 'Belawan', 'Semarang',
    'Laem Chabang', 'Bangkok', 'Haiphong', 'Ho Chi Minh', 'Phnom Penh',
    'Port Klang', 'Penang', 'Tanjung Pelepas',
    'Xingang', 'Tianjin', 'Dalian', 'Incheon',
    'Yokohama', 'Tokyo', 'Kobe', 'Osaka', 'Nagoya',
    'Kaohsiung', 'Keelung'
];

// Função para verificar se é rota de conexão
function isConnectionRoute(pol) {
    return CONNECTION_POLS.includes(pol);
}

const SERVICE_ROUTES = {
    // ========================================================
    // SANTOS - Carioca, Ipanema, Santana + Jade (todos)
    // ========================================================
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

    // ========================================================
    // RIO DE JANEIRO - Carioca, Santana + Jade (todos)
    // ========================================================
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

    // ========================================================
    // PARANAGUÁ - Carioca, Ipanema, Santana + Jade (todos)
    // ========================================================
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

    // ========================================================
    // NAVEGANTES - Ipanema, Santana + Jade (todos)
    // (Carioca NÃO atende Navegantes!)
    // ========================================================
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

    // ========================================================
    // ITAPOÁ - Carioca + Jade
    // ========================================================
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

    // ========================================================
    // ITAGUAÍ - Carioca, Santana + Jade
    // ========================================================
    'Shanghai-Itaguai': ['Carioca', 'Santana', 'Jade'],
    'Ningbo-Itaguai': ['Carioca', 'Santana', 'Jade'],
    'Shekou-Itaguai': ['Carioca', 'Santana', 'Jade'],
    'Busan-Itaguai': ['Carioca', 'Santana', 'Jade'],
    'Singapore-Itaguai': ['Carioca', 'Santana', 'Jade'],
    'Qingdao-Itaguai': ['Carioca', 'Santana', 'Jade'],

    // ========================================================
    // IMBITUBA - Santana + Jade
    // ========================================================
    'Shanghai-Imbituba': ['Santana', 'Jade'],
    'Ningbo-Imbituba': ['Santana', 'Jade'],
    'Shekou-Imbituba': ['Santana', 'Jade'],
    'Busan-Imbituba': ['Santana', 'Jade'],
    'Singapore-Imbituba': ['Santana', 'Jade'],
    'Qingdao-Imbituba': ['Santana', 'Jade'],

    // ========================================================
    // ITAJAÍ - Santana + Jade
    // ========================================================
    'Shanghai-Itajai': ['Santana', 'Jade'],
    'Ningbo-Itajai': ['Santana', 'Jade'],
    'Shekou-Itajai': ['Santana', 'Jade'],
    'Busan-Itajai': ['Santana', 'Jade'],
    'Singapore-Itajai': ['Santana', 'Jade'],
    'Qingdao-Itajai': ['Santana', 'Jade'],

    // ========================================================
    // SUAPE - Jade (+ Santana conforme PDF)
    // ========================================================
    'Shanghai-Suape': ['Santana', 'Jade'],
    'Ningbo-Suape': ['Santana', 'Jade'],
    'Shekou-Suape': ['Santana', 'Jade'],
    'Busan-Suape': ['Santana', 'Jade'],
    'Singapore-Suape': ['Santana', 'Jade'],
    'Qingdao-Suape': ['Santana', 'Jade'],
    'Yantian-Suape': ['Jade'],
    'Hong Kong-Suape': ['Jade'],
    'Xiamen-Suape': ['Jade'],
    'Nansha-Suape': ['Jade'],
    'Fuzhou-Suape': ['Jade'],

    // ========================================================
    // SALVADOR - Jade (+ Santana conforme PDF)
    // ========================================================
    'Shanghai-Salvador': ['Santana', 'Jade'],
    'Ningbo-Salvador': ['Santana', 'Jade'],
    'Shekou-Salvador': ['Santana', 'Jade'],
    'Busan-Salvador': ['Santana', 'Jade'],
    'Singapore-Salvador': ['Santana', 'Jade'],
    'Qingdao-Salvador': ['Santana', 'Jade'],
    'Yantian-Salvador': ['Jade'],
    'Hong Kong-Salvador': ['Jade'],
    'Xiamen-Salvador': ['Jade'],
    'Nansha-Salvador': ['Jade'],
    'Fuzhou-Salvador': ['Jade'],

    // ========================================================
    // MONTEVIDEO - Ipanema
    // ========================================================
    'Shanghai-Montevideo': ['Ipanema'],
    'Ningbo-Montevideo': ['Ipanema'],
    'Shekou-Montevideo': ['Ipanema'],
    'Busan-Montevideo': ['Ipanema'],
    'Singapore-Montevideo': ['Ipanema'],
    'Yantian-Montevideo': ['Ipanema'],
    'Hong Kong-Montevideo': ['Ipanema'],

    // ========================================================
    // BUENOS AIRES - Ipanema
    // ========================================================
    'Shanghai-Buenos Aires': ['Ipanema'],
    'Ningbo-Buenos Aires': ['Ipanema'],
    'Shekou-Buenos Aires': ['Ipanema'],
    'Busan-Buenos Aires': ['Ipanema'],
    'Singapore-Buenos Aires': ['Ipanema'],
    'Yantian-Buenos Aires': ['Ipanema'],
    'Hong Kong-Buenos Aires': ['Ipanema'],

    // ========================================================
    // RIO GRANDE - Ipanema
    // ========================================================
    'Shanghai-Rio Grande': ['Ipanema'],
    'Ningbo-Rio Grande': ['Ipanema'],
    'Shekou-Rio Grande': ['Ipanema'],
    'Busan-Rio Grande': ['Ipanema'],
    'Singapore-Rio Grande': ['Ipanema'],
    'Yantian-Rio Grande': ['Ipanema'],
    'Hong Kong-Rio Grande': ['Ipanema'],

    // ========================================================
    // MANAUS - Santana (via T/S Santos)
    // ========================================================
    'Shanghai-Manaus': ['Santana'],
    'Ningbo-Manaus': ['Santana'],
    'Qingdao-Manaus': ['Santana'],
    'Busan-Manaus': ['Santana'],
    'Shekou-Manaus': ['Santana'],
    'Singapore-Manaus': ['Santana'],
    'Yantian-Manaus': ['Santana'],

    // ========================================================
    // VITÓRIA - Santana, Carioca (via T/S Rio de Janeiro)
    // ========================================================
    'Shanghai-Vitoria': ['Santana', 'Carioca'],
    'Ningbo-Vitoria': ['Santana', 'Carioca'],
    'Qingdao-Vitoria': ['Santana', 'Carioca'],
    'Busan-Vitoria': ['Santana', 'Carioca'],
    'Shekou-Vitoria': ['Santana', 'Carioca'],
    'Singapore-Vitoria': ['Santana', 'Carioca'],

    // ========================================================
    // PECÉM - Santana (via T/S Salvador ou Santos)
    // ========================================================
    'Shanghai-Pecem': ['Santana'],
    'Ningbo-Pecem': ['Santana'],
    'Qingdao-Pecem': ['Santana'],
    'Busan-Pecem': ['Santana'],
    'Shekou-Pecem': ['Santana'],
    'Singapore-Pecem': ['Santana'],

    // ========================================================
    // FORTALEZA - Santana (via T/S Salvador ou Santos)
    // ========================================================
    'Shanghai-Fortaleza': ['Santana'],
    'Ningbo-Fortaleza': ['Santana'],
    'Qingdao-Fortaleza': ['Santana'],
    'Busan-Fortaleza': ['Santana'],
    'Shekou-Fortaleza': ['Santana'],
    'Singapore-Fortaleza': ['Santana'],

    // ========================================================
    // BELÉM - Santana (via T/S Santos)
    // ========================================================
    'Shanghai-Belem': ['Santana'],
    'Ningbo-Belem': ['Santana'],
    'Qingdao-Belem': ['Santana'],
    'Busan-Belem': ['Santana'],
    'Shekou-Belem': ['Santana'],
    'Singapore-Belem': ['Santana']
};

// Função para obter serviços disponíveis para uma rota
// Retorna null se:
// - Rota não está mapeada
// - POL é de conexão (não filtrar)
function getAvailableServices(pol, pod) {
    // Se for rota de conexão, retornar null (não filtrar)
    if (isConnectionRoute(pol)) {
        return null;
    }
    
    const routeKey = `${pol}-${pod}`;
    return SERVICE_ROUTES[routeKey] || null;
}

// ============================================
// ENDPOINT: Obter serviços disponíveis
// ============================================
app.get('/api/available-services', (req, res) => {
    const { pol, pod } = req.query;
    
    if (!pol || !pod) {
        return res.json({ services: ALL_SERVICES, message: 'Selecione POL e POD' });
    }
    
    // Verificar se é rota de conexão
    if (isConnectionRoute(pol)) {
        return res.json({ 
            services: ALL_SERVICES,
            mapped: false,
            isConnection: true,
            message: `Rota de conexão (${pol}) - busca sem filtro`
        });
    }
    
    const services = getAvailableServices(pol, pod);
    const routeKey = `${pol}-${pod}`;
    const isMapped = SERVICE_ROUTES.hasOwnProperty(routeKey);
    
    if (!isMapped) {
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
// ENDPOINT: Exportar Excel
// ============================================
app.post('/api/export', async (req, res) => {
    const { pol, pod, sailings } = req.body;
    
    console.log('\n=== EXPORTANDO EXCEL ===');
    console.log('POL:', pol, '| POD:', pod);
    console.log('Sailings:', sailings?.length || 0);
    
    try {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'ALLOG Schedules';
        workbook.created = new Date();
        
        const sheet = workbook.addWorksheet('SCHEDULES');
        
        // Agrupar por serviço para mostrar no cabeçalho
        const services = [...new Set((sailings || []).map(s => s.service).filter(s => s && s !== '-'))];
        const servicesText = services.length > 0 ? `MSC (${services.join(', ')})` : 'MSC';
        
        // LINHA 1: Header com CARRIERS/SERVICES
        sheet.mergeCells('A1:F1');
        sheet.getCell('A1').value = 'CARRIERS / SERVICES';
        sheet.getCell('A1').font = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
        sheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
        sheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
        sheet.getCell('A1').border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
        
        sheet.getCell('G1').value = servicesText;
        sheet.getCell('G1').font = { bold: true, size: 10 };
        sheet.getCell('G1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
        sheet.getCell('G1').alignment = { horizontal: 'center', vertical: 'middle' };
        sheet.getCell('G1').border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
        sheet.mergeCells('G1:I1');
        
        // LINHA 2: Headers das colunas
        const headers = ['VESSEL', `ETD ${pol}`, 'T/S PORT', 'ETA T/S', `ETA ${pod}`, 'TRANSIT', 'SERVICO'];
        const headerRow = sheet.addRow(headers);
        headerRow.font = { bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
        headerRow.height = 20;
        headerRow.eachCell((cell, colNumber) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
        });
        
        // LINHAS DE DADOS
        (sailings || []).forEach((s, index) => {
            const etdFormatted = formatDateShort(s.etd);
            const etaFormatted = formatDateShort(s.eta);
            const transbordoDateFormatted = formatDateShort(s.transbordoDate);
            
            const row = sheet.addRow([
                s.vessel || '-',
                etdFormatted,
                s.transbordo && s.transbordo !== '-' ? s.transbordo : '-',
                s.transbordo && s.transbordo !== '-' ? transbordoDateFormatted : '-',
                etaFormatted,
                s.transit || '-',
                s.service || '-'
            ]);
            
            // Estilo alternado para linhas
            const bgColor = index % 2 === 0 ? 'FFFFFFFF' : 'FFF2F2F2';
            row.height = 18;
            row.eachCell((cell, colNumber) => {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
                cell.border = { top: {style:'thin', color: {argb:'FFD0D0D0'}}, left: {style:'thin', color: {argb:'FFD0D0D0'}}, bottom: {style:'thin', color: {argb:'FFD0D0D0'}}, right: {style:'thin', color: {argb:'FFD0D0D0'}} };
                cell.font = { size: 9 };
            });
            
            // Vessel em negrito e alinhado à esquerda
            row.getCell(1).font = { bold: true, size: 9 };
            row.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
        });
        
        // Larguras das colunas
        sheet.columns = [
            { width: 28 }, // VESSEL
            { width: 14 }, // ETD POL
            { width: 10 }, // T/S PORT
            { width: 14 }, // ETA T/S
            { width: 14 }, // ETA POD
            { width: 10 }, // TRANSIT
            { width: 12 }  // SERVICO
        ];
        
        // Gerar buffer e enviar
        const buffer = await workbook.xlsx.writeBuffer();
        
        const filename = `MSC_${pol}_${pod}.xlsx`;
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Length', buffer.length);
        res.send(buffer);
        
        console.log(`📊 Excel exportado: ${filename} (${buffer.length} bytes)`);
        
    } catch (error) {
        console.error('❌ Erro ao exportar:', error);
        res.status(500).json({ error: error.message });
    }
});

// Função de formatação para export - formato curto dd-mmm
function formatDateShort(dateStr) {
    if (!dateStr || dateStr === '-') return '-';
    try {
        const match = dateStr.match(/(\w{3})?\s*(\d{1,2})(?:st|nd|rd|th)?\s+(\w{3})\s+(\d{4})/);
        if (!match) return dateStr;
        const [, , day, monthName, year] = match;
        const monthsShort = { 'Jan': 'jan', 'Feb': 'fev', 'Mar': 'mar', 'Apr': 'abr', 'May': 'mai', 'Jun': 'jun', 'Jul': 'jul', 'Aug': 'ago', 'Sep': 'set', 'Oct': 'out', 'Nov': 'nov', 'Dec': 'dez' };
        const month = monthsShort[monthName] || monthName.toLowerCase();
        return `${day.padStart(2, '0')}-${month}`;
    } catch (e) {
        return dateStr;
    }
}

// Função de formatação antiga (mantida para compatibilidade)
function formatDateExport(dateStr) {
    if (!dateStr || dateStr === '-') return '-';
    try {
        const match = dateStr.match(/(\w{3})\s+(\d{1,2})(?:st|nd|rd|th)?\s+(\w{3})\s+(\d{4})/);
        if (!match) {
            const match2 = dateStr.match(/(\d{1,2})(?:st|nd|rd|th)?\s+(\w{3})\s+(\d{4})/);
            if (!match2) return dateStr;
            const [, day, monthName, year] = match2;
            const months = { 'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12' };
            const month = months[monthName];
            if (!month) return dateStr;
            return `${day.padStart(2, '0')}/${month}/${year}`;
        }
        const [, weekDay, day, monthName, year] = match;
        const months = { 'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12' };
        const month = months[monthName];
        if (!month) return dateStr;
        return `${weekDay} ${day.padStart(2, '0')}/${month}/${year}`;
    } catch (e) {
        return dateStr;
    }
}

// ============================================
// SCRAPER: MSC
// ============================================
async function scrapeMSC(pol, pod, service = null) {
    const sailings = [];
    
    console.log(`\n🚢 === MSC SCRAPER ===`);
    console.log(`Rota: ${pol} → ${pod}`);
    
    // Verificar se é rota de conexão
    const isConnection = isConnectionRoute(pol);
    if (isConnection) {
        console.log(`📍 Rota de CONEXÃO - não filtrar por serviço`);
    }
    
    // Verificar se a rota está mapeada e se o serviço é válido
    const availableServices = getAvailableServices(pol, pod);
    let shouldFilter = false;
    let filterService = null;
    
    if (isConnection) {
        // Rota de conexão - NUNCA filtrar
        shouldFilter = false;
        console.log(`Buscando todos os serviços (conexão)`);
    } else if (service && service !== 'ALL') {
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
        
        // Verificar se é rota de conexão
        const isConnection = isConnectionRoute(pol);
        
        const data = await page.evaluate((filterService, isConn) => {
            const results = [];
            const seenVessels = new Map();
            const resultsArea = document.body.innerText;
            const lines = resultsArea.split('\n');
            
            let currentDeparture = null;
            let currentArrival = null;
            let currentVessel = null;
            let currentTransit = null;
            let currentRouting = null;
            let resultIndex = 0;
            
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
                    currentTransit = parseInt(transitMatch[1]);
                    continue;
                }
                
                // Detectar tipo de rota
                if (line === 'Direct' || line === 'Transhipment') {
                    currentRouting = line === 'Transhipment' ? 'Transbordo' : line;
                    
                    if (currentVessel) {
                        // Para rotas de CONEXÃO: manter TODOS os resultados (incluindo duplicados)
                        if (isConn) {
                            results.push({
                                index: resultIndex++,
                                service: filterService || '-',
                                vessel: currentVessel,
                                etd: currentDeparture || '-',
                                eta: currentArrival || '-',
                                transit: currentTransit ? `${currentTransit} dias` : '-',
                                routeType: currentRouting || '-'
                            });
                        } else {
                            // Para rotas DIRETAS: manter apenas um por navio (maior transit)
                            const existingEntry = seenVessels.get(currentVessel);
                            
                            if (!existingEntry) {
                                seenVessels.set(currentVessel, {
                                    service: filterService || '-',
                                    vessel: currentVessel,
                                    etd: currentDeparture || '-',
                                    eta: currentArrival || '-',
                                    transit: currentTransit || 0,
                                    routeType: currentRouting || '-'
                                });
                            } else if (currentTransit > existingEntry.transit) {
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
            
            // Para rotas DIRETAS: converter Map para Array
            if (!isConn) {
                return { 
                    results: Array.from(seenVessels.values()).map(entry => ({
                        ...entry,
                        transit: entry.transit ? `${entry.transit} dias` : '-'
                    })), 
                    isConnection: false 
                };
            }
            
            return { results, isConnection: true };
        }, service && service !== 'ALL' ? service : null, isConnection);
        
        console.log(`   Encontrados: ${data.results.length} resultados`);
        
        // 8. Para TODAS as rotas: clicar em cada resultado para obter serviço e transbordo
        let finalResults = data.results;
        
        if (data.results.length > 0) {
            console.log('8. Extraindo detalhes de cada sailing...');
            
            await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-08-antes-cliques.png') });
            
            for (let i = 0; i < Math.min(data.results.length, 6); i++) {
                const item = data.results[i];
                console.log(`   [${i+1}/${Math.min(data.results.length, 6)}] ${item.vessel}...`);
                
                try {
                    // Primeiro, garantir que nenhum painel está aberto
                    // Clicar no body ou pressionar ESC
                    await page.keyboard.press('Escape');
                    await new Promise(r => setTimeout(r, 500));
                    
                    // Scroll para o topo da lista primeiro
                    await page.evaluate(() => {
                        const header = document.querySelector('h1, h2, [class*="header"]');
                        if (header) header.scrollIntoView({ block: 'start' });
                    });
                    await new Promise(r => setTimeout(r, 300));
                    
                    // Agora procurar e clicar na linha correta
                    const clicked = await page.evaluate(async (index, vesselName) => {
                        // Procurar todas as linhas de resultado
                        const rows = [];
                        document.querySelectorAll('div, article, section').forEach(el => {
                            const text = el.innerText || '';
                            const rect = el.getBoundingClientRect();
                            
                            // Linha de resultado válida
                            if (text.includes('MSC') && 
                                text.includes('Days') &&
                                (text.includes('Transhipment') || text.includes('Direct')) &&
                                rect.width > 800 && 
                                rect.height > 50 && 
                                rect.height < 150 &&
                                rect.top > 0) {
                                rows.push({ el, top: rect.top, text: text.substring(0, 100) });
                            }
                        });
                        
                        rows.sort((a, b) => a.top - b.top);
                        
                        // Remover duplicatas
                        const uniqueRows = [];
                        let lastTop = -100;
                        for (const row of rows) {
                            if (row.top - lastTop > 40) {
                                uniqueRows.push(row);
                                lastTop = row.top;
                            }
                        }
                        
                        if (index >= uniqueRows.length) {
                            return { success: false, error: `Índice ${index} >= ${uniqueRows.length} linhas` };
                        }
                        
                        const targetRow = uniqueRows[index];
                        
                        // Scroll até a linha
                        targetRow.el.scrollIntoView({ block: 'center', behavior: 'instant' });
                        await new Promise(r => setTimeout(r, 400));
                        
                        // Procurar botão de expandir dentro da linha
                        const btn = targetRow.el.querySelector('button, [role="button"], svg, [class*="expand"], [class*="chevron"]');
                        if (btn) {
                            btn.click();
                        } else {
                            targetRow.el.click();
                        }
                        
                        return { success: true, rowText: targetRow.text.substring(0, 50) };
                    }, i, item.vessel);
                    
                    if (clicked && clicked.success) {
                        console.log(`      Clicou na linha ${i+1}`);
                        
                        // Esperar expansão do painel
                        await new Promise(r => setTimeout(r, 2000));
                        
                        // Screenshot do primeiro
                        if (i === 0) {
                            await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-09-primeiro-expandido.png') });
                        }
                        
                        // Extrair dados do painel expandido
                        const details = await page.evaluate((vesselName, currentService) => {
                            let service = currentService || '-';
                            let transbordo = '-';
                            let transbordoDate = '';
                            
                            const allText = document.body.innerText;
                            const lines = allText.split('\n').map(l => l.trim()).filter(l => l);
                            
                            // Serviços principais da MSC (NUNCA mudar para outros)
                            const mainServices = ['Santana', 'Carioca', 'Ipanema', 'Jade'];
                            
                            // Portos de transbordo conhecidos
                            const transbordoPorts = {
                                'Singapore': 'SIN',
                                'SGSIN': 'SIN',
                                'Busan': 'BUS',
                                'KRPUS': 'BUS',
                                'Tanjung Pelepas': 'TPP',
                                'MYTPP': 'TPP',
                                'Colombo': 'CMB',
                                'LKCMB': 'CMB',
                                'Port Klang': 'PKG',
                                'MYPKG': 'PKG',
                                'Santos': 'STS',
                                'BRSTS': 'STS'
                            };
                            
                            // Procurar no painel expandido - geralmente tem uma seção com os detalhes da rota
                            // O painel mostra: POL -> Transbordo -> POD com datas
                            
                            for (let i = 0; i < lines.length; i++) {
                                const line = lines[i];
                                
                                // Procurar serviço principal (só se ainda não tem)
                                if (service === '-') {
                                    for (const svc of mainServices) {
                                        if (line.includes(svc) && !line.includes('Filter')) {
                                            service = svc;
                                            break;
                                        }
                                    }
                                }
                                
                                // Detectar porto de transbordo
                                if (transbordo === '-') {
                                    for (const [portName, portCode] of Object.entries(transbordoPorts)) {
                                        if (line.includes(portName)) {
                                            // Verificar se é realmente transbordo (não é POL nem POD)
                                            // Transbordo geralmente aparece no meio da rota
                                            transbordo = portCode;
                                            
                                            // Buscar data próxima (antes ou depois da linha)
                                            for (let j = Math.max(0, i - 3); j <= Math.min(lines.length - 1, i + 3); j++) {
                                                const dateMatch = lines[j].match(/(\w{3}\s+\d{1,2}(?:st|nd|rd|th)?\s+\w{3}\s+\d{4})/);
                                                if (dateMatch) {
                                                    transbordoDate = dateMatch[1];
                                                    break;
                                                }
                                            }
                                            break;
                                        }
                                    }
                                }
                            }
                            
                            return { service, transbordo, transbordoDate };
                        }, item.vessel, item.service);
                        
                        // Atualizar item com os detalhes extraídos
                        if (details.service && details.service !== '-') {
                            item.service = details.service;
                        }
                        if (details.transbordo && details.transbordo !== '-') {
                            item.transbordo = details.transbordo;
                            item.transbordoDate = details.transbordoDate || '';
                        }
                        
                        console.log(`      ✓ Serviço: ${item.service}, Transbordo: ${item.transbordo} ${item.transbordoDate || ''}`);
                        
                        // IMPORTANTE: Fechar o painel antes de ir para o próximo
                        await page.evaluate((index) => {
                            const rows = [];
                            document.querySelectorAll('div, article, section').forEach(el => {
                                const text = el.innerText || '';
                                const rect = el.getBoundingClientRect();
                                if (text.includes('MSC') && text.includes('Days') &&
                                    rect.width > 800 && rect.height > 50 && rect.height < 150 && rect.top > 0) {
                                    rows.push({ el, top: rect.top });
                                }
                            });
                            rows.sort((a, b) => a.top - b.top);
                            const uniqueRows = [];
                            let lastTop = -100;
                            for (const row of rows) {
                                if (row.top - lastTop > 40) {
                                    uniqueRows.push(row);
                                    lastTop = row.top;
                                }
                            }
                            if (index < uniqueRows.length) {
                                const btn = uniqueRows[index].el.querySelector('button, svg');
                                if (btn) btn.click();
                                else uniqueRows[index].el.click();
                            }
                        }, i);
                        
                        await new Promise(r => setTimeout(r, 800));
                        
                    } else {
                        console.log(`      ⚠️ ${clicked ? clicked.error : 'Falhou'}`);
                    }
                    
                } catch (e) {
                    console.log(`      ❌ Erro: ${e.message}`);
                }
                
                // Pausa entre linhas
                await new Promise(r => setTimeout(r, 500));
            }
            
            await page.screenshot({ path: path.join(EXPORTS_DIR, 'msc-10-apos-cliques.png') });
        }
        
        console.log(`   Total final: ${finalResults.length} resultados`);
        
        // Limitar a 6 resultados
        const limitedResults = finalResults.slice(0, 6);
        
        limitedResults.forEach(r => {
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
                transbordo: r.transbordo || '-',
                transbordoDate: r.transbordoDate || '',
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
// EXCEL - MELHORIA 6: Formato de data melhorado + Transbordo
// ============================================
async function generateExcel(sailings, pol, pod, filename) {
    const workbook = new ExcelJS.Workbook();
    
    // Metadados do workbook para evitar corrupção
    workbook.creator = 'ALLOG Schedules';
    workbook.created = new Date();
    workbook.modified = new Date();
    
    const sheet = workbook.addWorksheet('SCHEDULES');
    
    sheet.mergeCells('A1:H1');
    sheet.getCell('A1').value = `ALLOG - Shipping Schedules: ${pol} - ${pod}`;
    sheet.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFF' } };
    sheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD700' } };
    sheet.getCell('A1').alignment = { horizontal: 'center' };
    
    const headers = ['CARRIER', 'SERVICO', 'NAVIO', 'ETD', 'TRANSBORDO', 'ETA', 'TRANSIT', 'TIPO'];
    sheet.addRow(headers);
    sheet.getRow(2).font = { bold: true, color: { argb: 'FFFFFF' } };
    sheet.getRow(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '333333' } };
    
    sailings.forEach(s => {
        const etdFormatted = formatDate(s.etd);
        const etaFormatted = formatDate(s.eta);
        const transbordoDateFormatted = formatDate(s.transbordoDate);
        const transbordoInfo = s.transbordo && s.transbordo !== '-' 
            ? `${s.transbordo}${transbordoDateFormatted && transbordoDateFormatted !== '-' ? ' (' + transbordoDateFormatted + ')' : ''}`
            : '-';
        
        const row = sheet.addRow([
            s.carrier || '-',
            s.service || '-',
            s.vessel || '-',
            etdFormatted,
            transbordoInfo,
            etaFormatted,
            s.transit || s.transitTime || '-',
            s.routeType || '-'
        ]);
        row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDE7' } };
    });
    
    sheet.columns = [
        { width: 10 }, // CARRIER
        { width: 12 }, // SERVICO
        { width: 25 }, // NAVIO
        { width: 18 }, // ETD
        { width: 22 }, // TRANSBORDO
        { width: 18 }, // ETA
        { width: 12 }, // TRANSIT
        { width: 12 }  // TIPO
    ];
    
    const filepath = path.join(EXPORTS_DIR, filename);
    
    // Usar writeFile diretamente (mais confiável para xlsx)
    await workbook.xlsx.writeFile(filepath);
    
    // Verificar se o arquivo foi criado corretamente
    const stats = fs.statSync(filepath);
    console.log(`📊 Excel gerado: ${filename} (${stats.size} bytes)`);
}

// MELHORIA 6: Função para formatar data
// Formato: "Thu 21/01/2026"
function formatDate(dateStr) {
    if (!dateStr || dateStr === '-') return '-';
    
    try {
        // Entrada: "Wed 21st Jan 2026" ou "Sat 17th Jan 2026" ou "Sun 25th Jan 2026"
        // Saída: "Thu 21/01/2026" (dia da semana abreviado + dd/mm/yyyy)
        
        const match = dateStr.match(/(\w{3})\s+(\d{1,2})(?:st|nd|rd|th)?\s+(\w{3})\s+(\d{4})/);
        if (!match) {
            // Tentar formato sem dia da semana
            const match2 = dateStr.match(/(\d{1,2})(?:st|nd|rd|th)?\s+(\w{3})\s+(\d{4})/);
            if (!match2) return dateStr;
            
            const [, day, monthName, year] = match2;
            const months = {
                'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04',
                'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08',
                'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'
            };
            const month = months[monthName];
            if (!month) return dateStr;
            const dayPadded = day.padStart(2, '0');
            return `${dayPadded}/${month}/${year}`;
        }
        
        const [, weekDay, day, monthName, year] = match;
        
        const months = {
            'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04',
            'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08',
            'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'
        };
        
        const month = months[monthName];
        if (!month) return dateStr;
        
        const dayPadded = day.padStart(2, '0');
        
        return `${weekDay} ${dayPadded}/${month}/${year}`;
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
