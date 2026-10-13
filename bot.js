const { Telegraf, Scenes, session, Markup } = require('telegraf');
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
    console.error('Excel файл яратишда хатолик:', err);
  }
}

// Рўйхатдан ўтиш сўровномаси
const registerWizard = new Scenes.WizardScene(
  'REGISTER_WIZARD',
  
  // 1-қадам: Исм-фамилияни қабул қилиш
  async (ctx) => {
    const text = ctx.message?.text;
    
    // Агар буйруқ юборилса, сценарийдан чиқамиз
    if (text?.startsWith('/')) {
      return ctx.scene.leave();
    }

    if (!text) {
      await ctx.reply('Илтимос, исм ва фамилиянгизни тўлиқ матн кўринишида киритинг!\n\nМасалан: Жасур Уразбаев');
      return;
    }

    ctx.wizard.state.data = { name: text };
    await ctx.reply(
      `Хурсандмиз, ${text}!\n\n` +
      'Ёшингизни киритинг (волонтёрлик учун 14 дан 30 ёшгача бўлиши керак):'
    );
    return ctx.wizard.next();
  },

  // 2-қадам: Ёш
  async (ctx) => {
    const text = ctx.message?.text;
    if (text?.startsWith('/')) {
      return ctx.scene.leave();
    }

    const age = parseInt(text, 10);
    if (isNaN(age) || age < 14 || age > 30) {
      await ctx.reply(
        'Кечирасиз! Волонтёрлик сўровномасида фақат 14 дан 30 ёшгача бўлган ёшлар қатнаша олади.\n\n' +
        'Илтимос, ёшингизни рақамларда тўғри киритинг (14-30):'
      );
      return;
    }
    ctx.wizard.state.data.age = age;

    const keyboardButtons = [];
    for (let i = 0; i < NEIGHBORHOODS.length; i += 2) {
      const row = [NEIGHBORHOODS[i]];
      if (NEIGHBORHOODS[i + 1]) row.push(NEIGHBORHOODS[i + 1]);
      keyboardButtons.push(row);
    }

    await ctx.reply(
      'Қайси маҳалла ёки овулда яшайсиз?\n\n' +
      'Қуйидаги тугмалардан бирини босиб танланг:',
      Markup.keyboard(keyboardButtons).resize().oneTime()
    );
    return ctx.wizard.next();
  },

  // 3-қадам: Телефон рақам
  async (ctx) => {
    const selectedNeighborhood = ctx.message?.text;
    if (selectedNeighborhood?.startsWith('/')) {
      return ctx.scene.leave();
    }
    
    if (!selectedNeighborhood || !NEIGHBORHOODS.includes(selectedNeighborhood)) {
      await ctx.reply('Илтимос, маҳаллангизни қуйидаги тугмалардан бирини босиб танланг!');
      return;
    }

    ctx.wizard.state.data.neighborhood = selectedNeighborhood;
    
    await ctx.reply(
      'Сиз билан боғланишимиз учун телефон рақамингизни юборинг:\n\n' +
      'Пастдаги "Телефон рақамни юбориш" тугмасини босинг ёки рақамингизни ёзиб қолдиринг (+9989X1234567):',
      Markup.keyboard([
        [Markup.button.contactRequest('Телефон рақамни юбориш')]
      ]).resize().oneTime()
    );
    return ctx.wizard.next();
  },

  // 4-қадам: Сақлаш
  async (ctx) => {
    let phone = '';
    const text = ctx.message?.text;

    if (text?.startsWith('/')) {
      return ctx.scene.leave();
    }
    
    if (ctx.message?.contact?.phone_number) {
      phone = ctx.message.contact.phone_number;
    } else if (text) {
      phone = text;
    } else {
      await ctx.reply('Илтимос, телефон рақамингизни тугма орқали юборинг ёки матн шаклида ёзиб қолдиринг.');
      return;
    }

    ctx.wizard.state.data.phone = phone;

    try {
      await initExcel();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(EXCEL_PATH);
      const worksheet = workbook.getWorksheet('Волонтёрлар');

      worksheet.addRow([
        ctx.from.id,
        ctx.wizard.state.data.name,
        ctx.wizard.state.data.age,
        ctx.wizard.state.data.neighborhood,
        ctx.wizard.state.data.phone,
        new Date().toLocaleString('uz-UZ')
      ]);

      await workbook.xlsx.writeFile(EXCEL_PATH);

      await ctx.reply(
        'Табриклаймиз! Сўровнома муваффақиятли тўлдирилди!\n\n' +
        'Киритилган маълумотларингиз:\n' +
        `Исм ва Фамилия: ${ctx.wizard.state.data.name}\n` +
        `Ёшингиз: ${ctx.wizard.state.data.age}\n` +
        `Маҳаллангиз: ${ctx.wizard.state.data.neighborhood}\n` +
        `Телефонингиз: ${ctx.wizard.state.data.phone}\n\n` +
        'Сизнинг аризангиз қабул қилинди. Тез орада сиз билан боғланамиз!',
        Markup.removeKeyboard()
      );
    } catch (err) {
      console.error('Excel сақлашда хатолик:', err);
      await ctx.reply('Маълумотларни сақлашда хатолик юз берди. Илтимос, /start ни босиб қайтадан уриниб кўринг.');
    }

    return ctx.scene.leave();
  }
);

const bot = new Telegraf(BOT_TOKEN);
const stage = new Scenes.Stage([registerWizard]);

bot.use(session());

const sendExcelFile = async (ctx) => {
  if (ctx.from.id !== ADMIN_ID) {
    return ctx.reply('Кечирасиз, бу буйруқ фақат админ учун!');
  }

  if (fs.existsSync(EXCEL_PATH)) {
    await ctx.replyWithDocument({
      source: EXCEL_PATH,
      filename: 'Тахиатош_актив_волонтёрлари.xlsx'
    });
  } else {
    await ctx.reply('Ҳозиргача ҳеч ким рўйхатдан ўтмаган.');
  }
};

// Командаларга биринчи ишлов бериш
bot.command('export', async (ctx) => {
  if (ctx.scene) await ctx.scene.leave();
  await sendExcelFile(ctx);
});

bot.command('admin', async (ctx) => {
  if (ctx.from.id !== ADMIN_ID) {
    return ctx.reply('Кечирасиз, бу буйруқ фақат админ учун!');
  }
  if (ctx.scene) await ctx.scene.leave();
  await ctx.reply(
    'Админ панели:\n\n' +
    'Барча йиғилган волонтёрлар рўйхатини .xlsx файлида олиш учун пастдаги тугмани босинг:',
    Markup.keyboard([
      ['Excel базани юклаб олиш']
    ]).resize()
  );
});

bot.hears('Excel базани юклаб олиш', sendExcelFile);

bot.use(stage.middleware());

// /start
bot.start(async (ctx) => {
  try {
    await initExcel();
    await ctx.scene.leave();
    await ctx.reply(
      'Ассалому алайкум! Тахиатош тумани волонтёрлари сафига хуш келибсиз!\n\n' +
      'Илтимос, Исм ва Фамилиянгизни киритинг:\n' +
      '(Масалан: Жасур Уразбаев)'
    );
    await ctx.scene.enter('REGISTER_WIZARD');
  } catch (err) {
    console.error('/start хатолиги:', err);
  }
});

bot.catch((err, ctx) => {
  console.error(`Бот хатолиги [${ctx.updateType}]:`, err);
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