require('dotenv').config();
const { Telegraf, Markup, session } = require('telegraf');
const Database = require('better-sqlite3');
const PptxGenJS = require('pptxgenjs');
const { exec, execSync, spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const Jimp = require('jimp');
const { PDFDocument } = require('pdf-lib');

// ==================== CONFIG ====================
const BOT_TOKEN = process.env.BOT_TOKEN;
const GROQ_API_KEY = process.env.GROQ_API_KEY;
const ADMIN_ID = parseInt(process.env.ADMIN_ID || '0');
const FREE_SLIDES = 999999; // Hammasi bepul

if (!BOT_TOKEN || !GROQ_API_KEY) {
    console.error('❌ BOT_TOKEN va GROQ_API_KEY kerak!');
    process.exit(1);
}

// ==================== DB ====================
const db = new Database('./users.db');
db.exec(`
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY,
        name TEXT,
        surname TEXT,
        username TEXT,
        lang TEXT DEFAULT 'uz',
        freeUsed INTEGER DEFAULT 0,
        paidUsed INTEGER DEFAULT 0,
        paidUntil TEXT,
        joined TEXT
    );
    CREATE TABLE IF NOT EXISTS payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        userId INTEGER,
        amount INTEGER,
        paket TEXT,
        status TEXT,
        date TEXT
    );
`);

// ==================== LANGUAGES ====================
const T = {
    uz: {
        welcome: "👋 Xush kelibsiz! SlaydTop botiga xush kelibsiz!",
        chooseLang: "🌐 Tilni tanlang:",
        chooseAction: "📋 Nima qilmoqchisiz?",
        chooseSlideCount: "📊 Nechta slayd kerak? (1-50)",
        chooseTemplate: "🎨 Shablon tanlang:",
        enterTopic: "✍️ Mavzuni yozing:",
        generating: "⏳ Slayd tayyorlanmoqda...",
        ready: "✅ Slayd tayyor!",
        error: "❌ Xatolik yuz berdi. Qayta urinib ko'ring.",
        noFree: "🚫 Bepul slaydlar tugadi. Paket sotib oling.",
        buyPacket: "💳 Paket sotib olish",
        myProfile: "👤 Mening profilim",
        support: "📞 Qo'llab-quvvatlash",
        back: "🔙 Orqaga",
        mainMenu: "🏠 Asosiy menyu",
        sinov: "Sinov",
        mini: "Mini",
        standard: "Standard",
        premium: "Premium",
        preparedBy: "Tayyorladi",
        page: "Sahifa",
        of: "dan",
        test: "TEST",
        crossword: "KRASSVORD",
        essay: "Referat",
        independentWork: "Mustaqil Ish",
        infographic: "INFOGRAFIKA",
        collage: "Kollaj",
        pdfDocument: "PDF Hujjat",
        diagram: "Diagrammali",
        unexpectedError: "Kutilmagan xatolik yuz berdi."
    },
    ru: {
        welcome: "👋 Добро пожаловать! Добро пожаловать в бот SlaydTop!",
        chooseLang: "🌐 Выберите язык:",
        chooseAction: "📋 Что вы хотите сделать?",
        chooseSlideCount: "📊 Сколько слайдов нужно? (1-50)",
        chooseTemplate: "🎨 Выберите шаблон:",
        enterTopic: "✍️ Напишите тему:",
        generating: "⏳ Слайд готовится...",
        ready: "✅ Слайд готов!",
        error: "❌ Произошла ошибка. Попробуйте снова.",
        noFree: "🚫 Бесплатные слайды закончились. Купите пакет.",
        buyPacket: "💳 Купить пакет",
        myProfile: "👤 Мой профиль",
        support: "📞 Поддержка",
        back: "🔙 Назад",
        mainMenu: "🏠 Главное меню",
        sinov: "Пробный",
        mini: "Мини",
        standard: "Стандарт",
        premium: "Премиум",
        preparedBy: "Подготовил",
        page: "Страница",
        of: "из",
        test: "ТЕСТ",
        crossword: "КРОССВОРД",
        essay: "Реферат",
        independentWork: "Самостоятельная работа",
        infographic: "ИНФОГРАФИКА",
        collage: "Коллаж",
        pdfDocument: "PDF Документ",
        diagram: "С диаграммой",
        unexpectedError: "Произошла непредвиденная ошибка."
    },
    en: {
        welcome: "👋 Welcome! Welcome to SlaydTop bot!",
        chooseLang: "🌐 Choose language:",
        chooseAction: "📋 What do you want to do?",
        chooseSlideCount: "📊 How many slides do you need? (1-50)",
        chooseTemplate: "🎨 Choose template:",
        enterTopic: "✍️ Write the topic:",
        generating: "⏳ Slide is being prepared...",
        ready: "✅ Slide is ready!",
        error: "❌ An error occurred. Please try again.",
        noFree: "🚫 Free slides are over. Buy a package.",
        buyPacket: "💳 Buy package",
        myProfile: "👤 My profile",
        support: "📞 Support",
        back: "🔙 Back",
        mainMenu: "🏠 Main menu",
        sinov: "Trial",
        mini: "Mini",
        standard: "Standard",
        premium: "Premium",
        preparedBy: "Prepared by",
        page: "Page",
        of: "of",
        test: "TEST",
        crossword: "CROSSWORD",
        essay: "Essay",
        independentWork: "Independent Work",
        infographic: "INFOGRAPHIC",
        collage: "Collage",
        pdfDocument: "PDF Document",
        diagram: "Diagram",
        unexpectedError: "An unexpected error occurred."
    }
};

function t(userId, key) {
    const user = getUser(userId);
    const lang = user?.lang || 'uz';
    return T[lang]?.[key] || T.uz[key] || key;
}

function getLang(userId) {
    const user = getUser(userId);
    return user?.lang || 'uz';
}

function pptxLabels(lang) {
    return T[lang] || T.uz;
}

// ==================== DB FUNCTIONS ====================
function getUser(userId) {
    return db.prepare('SELECT * FROM users WHERE id = ?').get(userId);
}

function addUser(userId, name, surname, username) {
    const existing = getUser(userId);
    if (!existing) {
        db.prepare('INSERT INTO users (id, name, surname, username, joined) VALUES (?, ?, ?, ?, ?)')
            .run(userId, name || '', surname || '', username || '', new Date().toISOString());
    }
}

function updateLang(userId, lang) {
    db.prepare('UPDATE users SET lang = ? WHERE id = ?').run(lang, userId);
}

function hasFreeSlides(userId) {
    const user = getUser(userId);
    return (user?.freeUsed || 0) < FREE_SLIDES;
}

function useFreeSlide(userId) {
    db.prepare('UPDATE users SET freeUsed = freeUsed + 1 WHERE id = ?').run(userId);
}

function getPaket(userId, count) {
    const l = pptxLabels(getLang(userId));
    const isFree = count <= 4 && hasFreeSlides(userId);
    if (isFree) return { nom: l.sinov, emoji: '🎁', narx: 0, min: 1, max: 4 };
    if (count <= 10) return { nom: l.mini, emoji: '⚡', narx: 5000, min: 1, max: 10 };
    if (count <= 25) return { nom: l.standard, emoji: '⭐', narx: 10000, min: 11, max: 25 };
    return { nom: l.premium, emoji: '💎', narx: 20000, min: 26, max: 50 };
}

// ==================== KEYBOARDS ====================
function langKeyboard() {
    return Markup.keyboard([
        ['🇺🇿 O'zbekcha', '🇷🇺 Русский', '🇬🇧 English']
    ]).resize();
}

function mainKeyboard(userId) {
    const l = pptxLabels(getLang(userId));
    return Markup.keyboard([
        [l.buyPacket, l.myProfile],
        [l.support]
    ]).resize();
}

function backKeyboard(userId) {
    const l = pptxLabels(getLang(userId));
    return Markup.keyboard([[l.back]]).resize();
}

function slideCountKeyboard(userId) {
    const l = pptxLabels(getLang(userId));
    const buttons = [];
    for (let i = 1; i <= 50; i += 5) {
        const row = [];
        for (let j = i; j < i + 5 && j <= 50; j++) {
            row.push(`🎨 ${j}`);
        }
        buttons.push(row);
    }
    buttons.push([l.back]);
    return Markup.keyboard(buttons).resize();
}

function templateKeyboard(userId) {
    const l = pptxLabels(getLang(userId));
    return Markup.keyboard([
        ['1. Classic', '2. Modern', '3. Dark'],
        ['4. Minimal', '5. Colorful', '6. Professional'],
        [l.back]
    ]).resize();
}

function actionKeyboard(userId) {
    const l = pptxLabels(getLang(userId));
    return Markup.keyboard([
        ['📊 Slayd', '📋 Test', '🧩 Krassvord'],
        ['📝 Referat', '📈 Infografika', '🖼️ Kollaj'],
        [l.back]
    ]).resize();
}

// ==================== AI FUNCTIONS ====================
async function generateWithGroq(prompt, maxTokens = 2000) {
    try {
        const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${GROQ_API_KEY}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                model: 'llama3-70b-8192',
                messages: [{ role: 'user', content: prompt }],
                max_tokens: maxTokens,
                temperature: 0.7
            })
        });
        const data = await res.json();
        if (!data || !data.choices || !data.choices[0] || !data.choices[0].message || !data.choices[0].message.content) {
            console.error('GROQ xato:', JSON.stringify(data || {}).slice(0, 300));
            return null;
        }
        return data.choices[0].message.content;
    } catch (e) {
        console.error('GROQ error:', e.message);
        return null;
    }
}

async function downloadPollinationsImage(prompt, outputPath) {
    try {
        const safePrompt = encodeURIComponent(prompt.replace(/['"]/g, '').slice(0, 200)).replace(/%20/g, '+');
        const url = `https://image.pollinations.ai/prompt/${safePrompt}?width=1024&height=1024&nologo=true`;
        const res = await fetch(url);
        if (!res.ok) return null;
        const buffer = await res.arrayBuffer();
        fs.writeFileSync(outputPath, Buffer.from(buffer));
        return outputPath;
    } catch (e) {
        console.error('Pollinations error:', e.message);
        return null;
    }
}

// ==================== TEMPLATE PATH FIX ====================
function getTemplatePath(templateId) {
    // Railway uchun process.cwd() ishlatish kerak
    return templateId ? path.join(process.cwd(), `${templateId}.pptx`) : null;
}

// ==================== PYTHON FIX ====================
function runPythonScript(scriptPath, jsonPath, outputPath) {
    return new Promise((resolve, reject) => {
        // Railway da python3 emas, python ishlatiladi
        exec(`python "${scriptPath}" "${jsonPath}" "${outputPath}"`, (error, stdout, stderr) => {
            if (error) {
                console.error('Python error:', stderr);
                reject(error);
            } else {
                resolve(stdout);
            }
        });
    });
}

// ==================== JIMP FIX (v0.16.13) ====================
async function resizeImage(imgPath, maxWidth = 800, maxHeight = 600) {
    try {
        const jimg = await Jimp.read(imgPath);
        const oW = jimg.getWidth();
        const oH = jimg.getHeight();
        let w = oW, h = oH;
        if (oW > maxWidth || oH > maxHeight) {
            const ratio = Math.min(maxWidth / oW, maxHeight / oH);
            w = Math.round(oW * ratio);
            h = Math.round(oH * ratio);
        }
        await jimg.resize(w, h).writeAsync(imgPath);
        return imgPath;
    } catch (e) {
        console.error('Jimp error:', e.message);
        return null;
    }
}

// ==================== BOT INIT ====================
const bot = new Telegraf(BOT_TOKEN);

bot.use(session());
bot.use((ctx, next) => {
    if (!ctx.session) ctx.session = {};
    return next();
});

// ==================== ERROR HANDLER FIX ====================
bot.catch((err, ctx) => {
    console.error('Bot xato:', err.message);
    if (ctx?.from?.id) {
        const userId = ctx.from.id;
        ctx.reply(t(userId, 'unexpectedError')).catch(() => {});
    }
});

// ==================== LANGUAGE SELECTION ====================
bot.start(async (ctx) => {
    const userId = ctx.from.id;
    addUser(userId, ctx.from.first_name, ctx.from.last_name, ctx.from.username);
    await ctx.reply(t(userId, 'welcome'), langKeyboard());
});

bot.hears(/🇺🇿 O'zbekcha/, (ctx) => {
    updateLang(ctx.from.id, 'uz');
    ctx.reply(t(ctx.from.id, 'chooseAction'), mainKeyboard(ctx.from.id));
});

bot.hears(/🇷🇺 Русский/, (ctx) => {
    updateLang(ctx.from.id, 'ru');
    ctx.reply(t(ctx.from.id, 'chooseAction'), mainKeyboard(ctx.from.id));
});

bot.hears(/🇬🇧 English/, (ctx) => {
    updateLang(ctx.from.id, 'en');
    ctx.reply(t(ctx.from.id, 'chooseAction'), mainKeyboard(ctx.from.id));
});

// ==================== MAIN MENU ====================
bot.hears([/💳 .*Paket.*/, /💳 .*пакет.*/, /💳 .*package.*/], async (ctx) => {
    const userId = ctx.from.id;
    const l = pptxLabels(getLang(userId));
    await ctx.reply(`${l.buyPacket}:

🎁 ${l.sinov} - 3 ta bepul
⚡ ${l.mini} - 5,000 so'm (10 ta)
⭐ ${l.standard} - 10,000 so'm (25 ta)
💎 ${l.premium} - 20,000 so'm (50 ta)`, mainKeyboard(userId));
});

bot.hears([/👤 .*profil.*/, /👤 .*профиль.*/, /👤 .*profile.*/], async (ctx) => {
    const userId = ctx.from.id;
    const user = getUser(userId);
    const l = pptxLabels(getLang(userId));
    const freeLeft = Math.max(0, FREE_SLIDES - (user?.freeUsed || 0));
    await ctx.reply(`👤 ${l.myProfile}:

🎁 ${l.sinov}: ${freeLeft} ta qoldi
📊 Jami: ${user?.paidUsed || 0} ta`, mainKeyboard(userId));
});

bot.hears([/📞 .*Qo'llab.*/, /📞 .*Поддержка.*/, /📞 .*Support.*/], async (ctx) => {
    await ctx.reply('📞 @admin_username', mainKeyboard(ctx.from.id));
});

// ==================== SLIDE GENERATION START ====================
bot.hears(/📊 Slayd/, async (ctx) => {
    const userId = ctx.from.id;
    await ctx.reply(t(userId, 'chooseSlideCount'), slideCountKeyboard(userId));
    ctx.session.step = 'choose_count';
    ctx.session.action = 'slide';
});

bot.hears(/📋 Test/, async (ctx) => {
    const userId = ctx.from.id;
    await ctx.reply(t(userId, 'chooseSlideCount'), slideCountKeyboard(userId));
    ctx.session.step = 'choose_count';
    ctx.session.action = 'test';
});

bot.hears(/🧩 Krassvord/, async (ctx) => {
    const userId = ctx.from.id;
    await ctx.reply(t(userId, 'chooseSlideCount'), slideCountKeyboard(userId));
    ctx.session.step = 'choose_count';
    ctx.session.action = 'crossword';
});

bot.hears(/📝 Referat/, async (ctx) => {
    const userId = ctx.from.id;
    await ctx.reply(t(userId, 'chooseSlideCount'), slideCountKeyboard(userId));
    ctx.session.step = 'choose_count';
    ctx.session.action = 'essay';
});

bot.hears(/📈 Infografika/, async (ctx) => {
    const userId = ctx.from.id;
    await ctx.reply(t(userId, 'chooseSlideCount'), slideCountKeyboard(userId));
    ctx.session.step = 'choose_count';
    ctx.session.action = 'infographic';
});

bot.hears(/🖼️ Kollaj/, async (ctx) => {
    const userId = ctx.from.id;
    await ctx.reply(t(userId, 'chooseSlideCount'), slideCountKeyboard(userId));
    ctx.session.step = 'choose_count';
    ctx.session.action = 'collage';
});

// ==================== SLIDE COUNT SELECTION ====================
bot.hears(/🎨 (\d+)/, async (ctx) => {
    const userId = ctx.from.id;
    const count = parseInt(ctx.match[1]);
    if (count < 1 || count > 50) {
        return ctx.reply(t(userId, 'chooseSlideCount'));
    }

    const paket = getPaket(userId, count);
    if (false && paket.narx > 0) {
        // TODO: Payment logic
        await ctx.reply(`💳 ${paket.nom} paket: ${paket.narx} so'm
To'lovni amalga oshiring...`, mainKeyboard(userId));
        return;
    }

    ctx.session.count = count;
    ctx.session.step = 'choose_template';
    await ctx.reply(t(userId, 'chooseTemplate'), templateKeyboard(userId));
});

// ==================== TEMPLATE SELECTION ====================
bot.hears(/\d+\.\s*(.+)/, async (ctx) => {
    const userId = ctx.from.id;
    if (ctx.session.step !== 'choose_template') return;

    const templateName = ctx.match[1].toLowerCase();
    const templateMap = {
        'classic': 'template1', 'modern': 'template2', 'dark': 'template3',
        'minimal': 'template4', 'colorful': 'template5', 'professional': 'template6'
    };
    ctx.session.template = templateMap[templateName] || null;
    ctx.session.step = 'enter_topic';
    await ctx.reply(t(userId, 'enterTopic'), backKeyboard(userId));
});

// ==================== TOPIC INPUT & GENERATION ====================
bot.on('text', async (ctx) => {
    const userId = ctx.from.id;
    const text = ctx.message.text;
    const l = pptxLabels(getLang(userId));

    if (text === l.back || text === l.mainMenu) {
        ctx.session.step = null;
        return ctx.reply(t(userId, 'chooseAction'), mainKeyboard(userId));
    }

    if (ctx.session.step === 'enter_topic') {
        const topic = text;
        const count = ctx.session.count || 5;
        const template = ctx.session.template;
        const action = ctx.session.action || 'slide';

        await ctx.reply(t(userId, 'generating'));

        try {
            let outputPath;

            if (action === 'slide') {
                outputPath = await makeSlidePptx(userId, topic, count, template);
            } else if (action === 'test') {
                outputPath = await makeTestPptx(userId, topic, count);
            } else if (action === 'crossword') {
                outputPath = await makeCrosswordPptx(userId, topic, count);
            } else if (action === 'essay') {
                outputPath = await makeTextPptx(userId, topic, count);
            } else if (action === 'infographic') {
                outputPath = await makeInfoPptx(userId, topic, count);
            } else if (action === 'collage') {
                outputPath = await makeCollagePdf(userId, topic, count);
            }

            if (outputPath && fs.existsSync(outputPath)) {
                await ctx.replyWithDocument({ source: outputPath });
                fs.unlinkSync(outputPath);
                useFreeSlide(userId);
                await ctx.reply(t(userId, 'ready'), mainKeyboard(userId));
            } else {
                await ctx.reply(t(userId, 'error'));
            }
        } catch (e) {
            console.error('Generation error:', e);
            await ctx.reply(t(userId, 'error'));
        }

        ctx.session.step = null;
    }
});

// ==================== SLIDE GENERATION FUNCTIONS ====================
async function makeSlidePptx(userId, topic, count, templateId) {
    const l = pptxLabels(getLang(userId));
    const user = getUser(userId);
    const outputPath = path.join(process.cwd(), `slide_${userId}_${Date.now()}.pptx`);

    // AI content generation
    const prompt = `Create ${count} slide titles and bullet points for a presentation about "${topic}" in ${getLang(userId)} language. Format: JSON array with {title, bullets: []} for each slide.`;
    const aiContent = await generateWithGroq(prompt, 3000);
    let slides = [];
    try {
        slides = JSON.parse(aiContent || '[]');
    } catch (e) {
        slides = Array(count).fill(0).map((_, i) => ({
            title: `${topic} - ${l.page} ${i + 1}`,
            bullets: ['Content coming soon...']
        }));
    }

    const templateFile = getTemplatePath(templateId);

    if (templateFile && fs.existsSync(templateFile)) {
        // Use Python script with template
        const jsonPath = path.join(process.cwd(), `data_${userId}_${Date.now()}.json`);
        fs.writeFileSync(jsonPath, JSON.stringify({ topic, slides, lang: getLang(userId), user }));

        const scriptPath = path.join(process.cwd(), 'make_slide.py');
        await runPythonScript(scriptPath, jsonPath, outputPath);

        if (fs.existsSync(jsonPath)) fs.unlinkSync(jsonPath);
        if (fs.existsSync(outputPath)) return outputPath;
    }

    // Fallback: PptxGenJS
    return makeSlidePptxFallback(userId, topic, slides, outputPath);
}

async function makeSlidePptxFallback(userId, topic, slides, outputPath) {
    const l = pptxLabels(getLang(userId));
    const user = getUser(userId);
    const pptx = new PptxGenJS();

    pptx.layout = 'LAYOUT_16x9';
    pptx.author = 'SlaydTop Bot';
    pptx.title = topic;

    // Cover slide
    const cover = pptx.addSlide();
    cover.background = { color: '1F4E79' };
    cover.addText(topic, { x: 1, y: 2, w: 8, h: 1.5, fontSize: 36, color: 'FFFFFF', bold: true, align: 'center' });
    cover.addText(`${l.preparedBy}: ${user?.name || ''} ${user?.surname || ''}`, { x: 1, y: 4, w: 8, h: 0.5, fontSize: 14, color: 'D9E2F3', align: 'center' });

    // Content slides
    for (const slide of slides) {
        const s = pptx.addSlide();
        s.background = { color: 'FFFFFF' };
        s.addText(slide.title || 'Untitled', { x: 0.5, y: 0.5, w: 9, h: 0.8, fontSize: 24, color: '1F4E79', bold: true });

        if (slide.bullets && slide.bullets.length > 0) {
            s.addText(slide.bullets.map(b => `• ${b}`).join('\n'), { x: 0.5, y: 1.5, w: 9, h: 4, fontSize: 16, color: '333333' });
        }
    }

    await pptx.writeFile({ fileName: outputPath });
    return outputPath;
}

async function makeTestPptx(userId, topic, count) {
    const l = pptxLabels(getLang(userId));
    const outputPath = path.join(process.cwd(), `test_${userId}_${Date.now()}.pptx`);
    const pptx = new PptxGenJS();

    const prompt = `Create ${count} multiple choice questions about "${topic}" in ${getLang(userId)} language. Format: JSON array with {question, options: [], answer} for each question.`;
    const aiContent = await generateWithGroq(prompt, 3000);
    let questions = [];
    try {
        questions = JSON.parse(aiContent || '[]');
    } catch (e) {
        questions = Array(count).fill(0).map((_, i) => ({
            question: `Question ${i + 1} about ${topic}?`,
            options: ['A', 'B', 'C', 'D'],
            answer: 'A'
        }));
    }

    pptx.addSlide().addText(`${l.test}: ${topic}`, { x: 1, y: 2, w: 8, h: 1, fontSize: 32, bold: true, align: 'center' });

    for (const q of questions) {
        const s = pptx.addSlide();
        s.addText(q.question, { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 20, bold: true });
        if (q.options) {
            s.addText(q.options.map((o, i) => `${String.fromCharCode(65 + i)}. ${o}`).join('\n'), { x: 0.5, y: 1.5, w: 9, h: 3, fontSize: 16 });
        }
    }

    await pptx.writeFile({ fileName: outputPath });
    return outputPath;
}

async function makeCrosswordPptx(userId, topic, count) {
    const l = pptxLabels(getLang(userId));
    const outputPath = path.join(process.cwd(), `crossword_${userId}_${Date.now()}.pptx`);
    const pptx = new PptxGenJS();

    pptx.addSlide().addText(`${l.crossword}: ${topic}`, { x: 1, y: 2, w: 8, h: 1, fontSize: 32, bold: true, align: 'center' });

    const prompt = `Create ${count} crossword clues and answers about "${topic}" in ${getLang(userId)} language. Format: JSON array with {clue, answer} for each word.`;
    const aiContent = await generateWithGroq(prompt, 2000);
    let words = [];
    try {
        words = JSON.parse(aiContent || '[]');
    } catch (e) {
        words = Array(count).fill(0).map((_, i) => ({ clue: `Clue ${i + 1}`, answer: `WORD${i}` }));
    }

    const s = pptx.addSlide();
    s.addText(words.map((w, i) => `${i + 1}. ${w.clue}`).join('\n'), { x: 0.5, y: 0.5, w: 9, h: 4, fontSize: 14 });

    await pptx.writeFile({ fileName: outputPath });
    return outputPath;
}

async function makeTextPptx(userId, topic, count) {
    const l = pptxLabels(getLang(userId));
    const outputPath = path.join(process.cwd(), `text_${userId}_${Date.now()}.pptx`);
    const pptx = new PptxGenJS();

    const prompt = `Write a detailed essay or independent work about "${topic}" in ${getLang(userId)} language, divided into ${count} sections. Format: JSON array with {title, content} for each section.`;
    const aiContent = await generateWithGroq(prompt, 4000);
    let sections = [];
    try {
        sections = JSON.parse(aiContent || '[]');
    } catch (e) {
        sections = Array(count).fill(0).map((_, i) => ({ title: `Section ${i + 1}`, content: 'Content...' }));
    }

    pptx.addSlide().addText(`${l.essay}: ${topic}`, { x: 1, y: 2, w: 8, h: 1, fontSize: 32, bold: true, align: 'center' });

    for (const sec of sections) {
        const s = pptx.addSlide();
        s.addText(sec.title, { x: 0.5, y: 0.5, w: 9, h: 0.8, fontSize: 22, bold: true });
        s.addText(sec.content || '', { x: 0.5, y: 1.5, w: 9, h: 4, fontSize: 14 });
    }

    await pptx.writeFile({ fileName: outputPath });
    return outputPath;
}

async function makeInfoPptx(userId, topic, count) {
    const l = pptxLabels(getLang(userId));
    const outputPath = path.join(process.cwd(), `info_${userId}_${Date.now()}.pptx`);
    const pptx = new PptxGenJS();

    const prompt = `Create ${count} infographic data points and statistics about "${topic}" in ${getLang(userId)} language. Format: JSON array with {title, stat, description} for each point.`;
    const aiContent = await generateWithGroq(prompt, 3000);
    let points = [];
    try {
        points = JSON.parse(aiContent || '[]');
    } catch (e) {
        points = Array(count).fill(0).map((_, i) => ({ title: `Point ${i + 1}`, stat: '0%', description: 'Description' }));
    }

    pptx.addSlide().addText(`${l.infographic}: ${topic}`, { x: 1, y: 2, w: 8, h: 1, fontSize: 32, bold: true, align: 'center' });

    for (const p of points) {
        const s = pptx.addSlide();
        s.addText(p.title, { x: 0.5, y: 0.5, w: 9, h: 0.8, fontSize: 24, bold: true, color: 'FF6B35' });
        s.addText(p.stat || '', { x: 0.5, y: 1.5, w: 9, h: 1, fontSize: 48, bold: true, color: '1F4E79', align: 'center' });
        s.addText(p.description || '', { x: 0.5, y: 2.8, w: 9, h: 2, fontSize: 14 });
    }

    await pptx.writeFile({ fileName: outputPath });
    return outputPath;
}

async function makeCollagePdf(userId, topic, count) {
    const l = pptxLabels(getLang(userId));
    const outputPath = path.join(process.cwd(), `collage_${userId}_${Date.now()}.pdf`);

    // Generate images via Pollinations
    const images = [];
    for (let i = 0; i < count; i++) {
        const imgPath = path.join(process.cwd(), `img_${userId}_${i}_${Date.now()}.jpg`);
        const downloaded = await downloadPollinationsImage(`${topic} ${l.page} ${i + 1}`, imgPath);
        if (downloaded) images.push(downloaded);
    }

    // Create PDF with pdf-lib (no LibreOffice needed)
    const pdfDoc = await PDFDocument.create();
    for (const imgPath of images) {
        try {
            const imgBytes = fs.readFileSync(imgPath);
            let img;
            if (imgPath.endsWith('.jpg') || imgPath.endsWith('.jpeg')) {
                img = await pdfDoc.embedJpg(imgBytes);
            } else {
                img = await pdfDoc.embedPng(imgBytes);
            }
            const page = pdfDoc.addPage([600, 400]);
            page.drawImage(img, { x: 0, y: 0, width: 600, height: 400 });
            fs.unlinkSync(imgPath);
        } catch (e) {
            console.error('PDF image error:', e.message);
        }
    }

    const pdfBytes = await pdfDoc.save();
    fs.writeFileSync(outputPath, pdfBytes);
    return outputPath;
}

// ==================== ADMIN PANEL ====================
bot.hears([/📋 .*/, /👥 .*/, /📢 .*/], async (ctx) => {
    if (ctx.from.id !== ADMIN_ID) return;
    await ctx.reply('Admin panel', mainKeyboard(ctx.from.id));
});

// ==================== LAUNCH ====================
bot.launch();
console.log('🤖 Bot ishga tushdi!');

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
