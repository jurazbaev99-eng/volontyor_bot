const { Telegraf, Markup } = require('telegraf');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const http = require('http');

const BOT_TOKEN = '8913247281:AAGUGFSiQyKc4qPipaB2BGU_l4PXamNbDQ0';
const ADMIN_ID = 7470599966;

const EXCEL_PATH = path.join(__dirname, 'volunteers.xlsx');

const NEIGHBORHOODS = [
  'Тахиатош МФЙ',
  'Обод макон МФЙ',
  'Найманкўл ОФЙ',
  'Сарайкўл ОФЙ',
  'Кенегес ОФЙ',
  'Янги макон ОФЙ',
  'Шамчироқ МФЙ',
  'Жайхун МФЙ',
  'Нурли келажак МФЙ',
  'Дўстлик МФЙ',
  'Ҳалқлар дўстлиги МФЙ',
  'Айдин жол МФЙ',
  'Қушчилик МФЙ',
  'Марказий МФЙ'
];

const userState = {};

async function initExcel() {
  try {
    if (!fs.existsSync(EXCEL_PATH)) {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Волонтёрлар');
      worksheet.columns = [
        { header: 'ID', key: 'id', width: 15 },
        { header: 'Исм ва Фамилия', key: 'name', width: 25 },
        { header: 'Ёши', key: 'age', width: 10 },
        { header: 'Маҳалла / ОФЙ', key: 'neighborhood', width: 25 },
        { header: 'Телефон рақам', key: 'phone', width: 20 },
        { header: 'Рўйхатдан ўтган вақти', key: 'date', width: 22 }
      ];
      await workbook.xlsx.writeFile(EXCEL_PATH);
    }
  } catch (err) {
    console.error('Excel яратишда хатолик:', err);
  }
}

async function isAlreadyRegistered(userId) {
  try {
    if (!fs.existsSync(EXCEL_PATH)) return false;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_PATH);
    const worksheet = workbook.getWorksheet('Волонтёрлар');
    if (!worksheet) return false;

    let found = false;
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        const cellValue = row.getCell(1).value;
        if (cellValue && String(cellValue) === String(userId)) {
          found = true;
        }
      }
    });
    return found;
  } catch (err) {
    return false;
  }
}

const bot = new Telegraf(BOT_TOKEN);

const sendExcelFile = async (ctx) => {
  if (ctx.from.id !== ADMIN_ID) {
    return ctx.reply('⛔ Кечирасиз, бу буйруқ фақат админ учун!');
  }

  if (fs.existsSync(EXCEL_PATH)) {
    await ctx.replyWithDocument({
      source: EXCEL_PATH,
      filename: 'Тахиатош_актив_волонтёрлари.xlsx'
    });
  } else {
    await ctx.reply('📭 Ҳозиргача ҳеч ким рўйхатдан ўтмаган.');
  }
};

bot.command('export', sendExcelFile);
bot.hears('📊 Excel базани юклаб олиш', sendExcelFile);

bot.command('admin', async (ctx) => {
  if (ctx.from.id !== ADMIN_ID) {
    return ctx.reply('⛔ Кечирасиз, бу буйруқ фақат админ учун!');
  }
  await ctx.reply(
    '⚙️ Админ панели:\n\n' +
    'Барча йиғилган волонтёрлар рўйхатини .xlsx файлида олиш учун пастдаги тугмани босинг:',
    Markup.keyboard([
      ['📊 Excel базани юклаб олиш']
    ]).resize()
  );
});

bot.start(async (ctx) => {
  await initExcel();
  const userId = ctx.from.id;

  const registered = await isAlreadyRegistered(userId);
  if (registered) {
    return ctx.reply(
      '✅ **Сиз аллақачон волонтёрлар рўйхатидан ўтгансиз!**\n\n' +
      '🤝 Аризангиз қабул қилинган. Тез орада сиз билан боғланамиз!',
      Markup.removeKeyboard()
    );
  }

  userState[userId] = { step: 'NAME', data: {} };

  await ctx.reply(
    '🌟 Ассалому алайкум! Тахиатош тумани волонтёрлари сафига хуш келибсиз!\n\n' +
    '👤 Илтимос, Исм ва Фамилиянгизни киритинг:\n' +
    '(Масалан: Жасур Уразбаев)',
    Markup.removeKeyboard()
  );
});

bot.on('message', async (ctx) => {
  const userId = ctx.from.id;
  const text = ctx.message?.text;

  if (text?.startsWith('/')) return;

  const registered = await isAlreadyRegistered(userId);
  if (registered) {
    delete userState[userId];
    return ctx.reply('✅ Сиз аллақачон рўйхатдан ўтгансиз!');
  }

  if (!userState[userId]) {
    userState[userId] = { step: 'NAME', data: {} };
    return ctx.reply(
      '🌟 Илтимос, сўровномани бошлаш учун Исм ва Фамилиянгизни киритинг:\n(Масалан: Жасур Уразбаев)'
    );
  }

  const state = userState[userId];

  if (state.step === 'NAME') {
    if (!text) {
      return ctx.reply('⚠️ Илтимос, исм ва фамилиянгизни матн шаклида киритинг!\n\nМасалан: Жасур Уразбаев');
    }

    state.data.name = text;
    state.step = 'AGE';

    return ctx.reply(
      `Хурсандмиз, ${text}! 😊\n\n` +
      '🎂 Ёшингизни киритинг (волонтёрлик учун 14 дан 30 ёшгача бўлиши керак):'
    );
  }

  if (state.step === 'AGE') {
    const age = parseInt(text, 10);
    if (isNaN(age) || age < 14 || age > 30) {
      return ctx.reply(
        '🚫 Кечирасиз! Волонтёрлик сўровномасида фақат 14 дан 30 ёшгача бўлган ёшлар қатнаша олади.\n\n' +
        '🔢 Илтимос, ёшингизни рақамларда тўғри киритинг (14-30):'
      );
    }

    state.data.age = age;
    state.step = 'NEIGHBORHOOD';

    const keyboardButtons = [];
    for (let i = 0; i < NEIGHBORHOODS.length; i += 2) {
      const row = [NEIGHBORHOODS[i]];
      if (NEIGHBORHOODS[i + 1]) row.push(NEIGHBORHOODS[i + 1]);
      keyboardButtons.push(row);
    }

    return ctx.reply(
      '🏡 Қайси маҳалла ёки овулда яшайсиз?\n\n' +
      '👇 Қуйидаги тугмалардан бирини босиб танланг:',
      Markup.keyboard(keyboardButtons).resize().oneTime()
    );
  }

  if (state.step === 'NEIGHBORHOOD') {
    if (!text || !NEIGHBORHOODS.includes(text)) {
      return ctx.reply('⚠️ Илтимос, маҳаллангизни пастдаги тугмалардан бирини босиб танланг!');
    }

    state.data.neighborhood = text;
    state.step = 'PHONE';

    return ctx.reply(
      '📱 Сиз билан боғланишимиз учун телефон рақамингизни юборинг:\n\n' +
      'Пастдаги "📱 Телефон рақамни юбориш" тугмасини босинг ёки рақамингизни ёзиб қолдиринг (+9989X1234567):',
      Markup.keyboard([
        [Markup.button.contactRequest('📱 Телефон рақамни юбориш')]
      ]).resize().oneTime()
    );
  }

  if (state.step === 'PHONE') {
    let phone = '';

    if (ctx.message?.contact?.phone_number) {
      phone = ctx.message.contact.phone_number;
    } else if (text) {
      phone = text;
    } else {
      return ctx.reply('⚠️ Илтимос, телефон рақамингизни тугма орқали юборинг ёки матн шаклида ёзиб қолдиринг.');
    }

    state.data.phone = phone;

    try {
      await initExcel();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(EXCEL_PATH);
      const worksheet = workbook.getWorksheet('Волонтёрлар');

      worksheet.addRow([
        userId,
        state.data.name,
        state.data.age,
        state.data.neighborhood,
        state.data.phone,
        new Date().toLocaleString('uz-UZ')
      ]);

      await workbook.xlsx.writeFile(EXCEL_PATH);

      await ctx.reply(
        '🎉 Табриклаймиз! Сўровнома муваффақиятли тўлдирилди!\n\n' +
        '📋 Киритилган маълумотларингиз:\n' +
        `👤 Исм ва Фамилия: ${state.data.name}\n` +
        `🎂 Ёшингиз: ${state.data.age}\n` +
        `🏡 Маҳаллангиз: ${state.data.neighborhood}\n` +
        `📱 Телефонингиз: ${state.data.phone}\n\n` +
        '✨ Сизнинг аризангиз қабул қилинди. Тез орада сиз билан боғланамиз!',
        Markup.removeKeyboard()
      );

      delete userState[userId];
    } catch (err) {
      console.error('Excel сақлашда хатолик:', err);
      await ctx.reply('⚠️ Маълумотларни сақлашда хатолик юз берди. Илтимос, /start ни босиб қайтадан уриниб кўринг.');
    }
  }
});

bot.catch((err) => {
  console.error('Бот хатолиги:', err);
});

initExcel().then(() => {
  bot.launch();
  console.log('Бот муваффақиятли ишга тушди!');
});

const PORT = process.env.PORT || 10000;
http.createServer((req, res) => {
  res.writeHead(200, { 'Content-Type': 'text/plain' });
  res.end('Volontyor bot is running 24/7!');
}).listen(PORT, () => {
  console.log(`Сервер ${PORT}-портда муваффақиятли ишга тушди.`);
});

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));