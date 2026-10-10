import 'dotenv/config';
import { Telegraf, Markup } from 'telegraf';
import { google } from 'googleapis';

const required = [
  'BOT_TOKEN',
  'TELEGRAM_USER_ID',
  'SPREADSHEET_ID',
  'GOOGLE_APPLICATION_CREDENTIALS',
] as const;

for (const key of required) {
  if (!process.env[key]) {
    throw new Error(`Не задана переменная ${key} в .env`);
  }
}

const userId = Number(process.env.TELEGRAM_USER_ID);
if (!Number.isSafeInteger(userId) || userId <= 0) {
  throw new Error('TELEGRAM_USER_ID должен быть числовым Telegram ID.');
}

const bot = new Telegraf(process.env.BOT_TOKEN!);
const spreadsheetId = process.env.SPREADSHEET_ID!;
const sheetName = 'Операции';

const auth = new google.auth.GoogleAuth({
  keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS!,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

type Transaction = {
  id: string;
  rowIndex: number;
  date: string;
  type: 'Доход' | 'Расход';
  amount: number;
  category: string;
  comment: string;
};

type EditField = 'amount' | 'comment' | 'category';
type EditState = { id: string; field: EditField };
const editStates = new Map<number, EditState>();
const transactionTypeStates = new Map<number, Transaction['type']>();

function getMadridDate(): string {
  return new Date().toLocaleString('sv-SE', { timeZone: 'Europe/Madrid' });
}

function getMadridMonth(): string {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Europe/Madrid',
    year: 'numeric',
    month: '2-digit',
  }).formatToParts(new Date());
  const year = parts.find((part) => part.type === 'year')!.value;
  const month = parts.find((part) => part.type === 'month')!.value;
  return `${year}-${month}`;
}

function shiftMonth(month: string, offset: number): string {
  const [year, monthNumber] = month.split('-').map(Number);
  const date = new Date(Date.UTC(year, monthNumber - 1 + offset, 1));
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}`;
}

function money(value: number): string {
  return `${value.toLocaleString('ru-RU', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })} €`;
}

function getCategory(comment: string): string {
  const text = comment.toLowerCase();
  const categories: Record<string, string[]> = {
    Продукты: ['продукт', 'супермаркет', 'магазин', 'mercadona', 'lidl', 'aldi'],
    Транспорт: ['бензин', 'топливо', 'такси', 'автобус', 'метро', 'парковка', 'поезд'],
    'Кафе и рестораны': ['кафе', 'ресторан', 'кофе', 'обед', 'ужин', 'бар', 'доставка еды'],
    Жильё: ['аренда', 'квартира', 'коммунал', 'электричество', 'вода'],
    Зарплата: ['зарплата', 'аванс', 'оклад'],
    Развлечения: ['кино', 'игра', 'подписка', 'развлеч', 'spotify', 'netflix'],
    Здоровье: ['аптека', 'врач', 'лекарств', 'стоматолог'],
    Покупки: ['одежда', 'обувь', 'покупка', 'amazon'],
  };

  for (const [category, keywords] of Object.entries(categories)) {
    if (keywords.some((keyword) => text.includes(keyword))) return category;
  }
  return 'Другое';
}

function createId(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

async function getTransactions(): Promise<Transaction[]> {
  const result = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${sheetName}'!A2:F`,
  });

  const rows = result.data.values ?? [];
  const transactions: Transaction[] = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rowIndex = i + 2;
    if (row.length < 3) continue;

    const type = String(row[1]);
    if (type !== 'Доход' && type !== 'Расход') continue;

    const amount = Number(row[2]);
    if (!Number.isFinite(amount) || amount <= 0) continue;

    let id = String(row[5] ?? '').trim();
    if (!id) {
      id = createId();
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${sheetName}'!F${rowIndex}`,
        valueInputOption: 'RAW',
        requestBody: { values: [[id]] },
      });
    }

    transactions.push({
      id,
      rowIndex,
      date: String(row[0] ?? ''),
      type,
      amount,
      category: String(row[3] ?? 'Другое'),
      comment: String(row[4] ?? ''),
    });
  }

  return transactions;
}

async function addTransaction(transaction: Omit<Transaction, 'rowIndex'>): Promise<void> {
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `'${sheetName}'!A:F`,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    requestBody: {
      values: [
        [
          transaction.date,
          transaction.type,
          transaction.amount,
          transaction.category,
          transaction.comment,
          transaction.id,
        ],
      ],
    },
  });
}

function calculateBalance(transactions: Transaction[]): number {
  return transactions.reduce(
    (sum, transaction) =>
      sum + (transaction.type === 'Доход' ? transaction.amount : -transaction.amount),
    0,
  );
}

function summarizeMonth(transactions: Transaction[], month: string) {
  const items = transactions.filter((transaction) => transaction.date.startsWith(month));
  const income = items
    .filter((transaction) => transaction.type === 'Доход')
    .reduce((sum, transaction) => sum + transaction.amount, 0);
  const expenses = items
    .filter((transaction) => transaction.type === 'Расход')
    .reduce((sum, transaction) => sum + transaction.amount, 0);
  const categories = new Map<string, number>();

  for (const transaction of items) {
    if (transaction.type === 'Расход') {
      categories.set(
        transaction.category,
        (categories.get(transaction.category) ?? 0) + transaction.amount,
      );
    }
  }

  return {
    income,
    expenses,
    net: income - expenses,
    categories: [...categories.entries()].sort((a, b) => b[1] - a[1]),
  };
}

function compareValues(current: number, previous: number): string {
  const difference = current - previous;
  const sign = difference > 0 ? '+' : difference < 0 ? '−' : '';
  const percentage =
    previous !== 0
      ? ` (${sign}${Math.abs((difference / Math.abs(previous)) * 100).toFixed(1)}%)`
      : current !== 0
        ? ' (ранее 0 €)'
        : '';
  return `${sign}${money(Math.abs(difference))}${percentage}`;
}

async function buildReport(month: string): Promise<string> {
  const transactions = await getTransactions();
  const current = summarizeMonth(transactions, month);
  const previousMonth = shiftMonth(month, -1);
  const previous = summarizeMonth(transactions, previousMonth);

  const categoryText = current.categories.length
    ? current.categories
        .map(([category, amount]) => {
          const percentage = current.expenses > 0 ? (amount / current.expenses) * 100 : 0;
          return `• ${category}: ${money(amount)} (${percentage.toFixed(1)}%)`;
        })
        .join('\n')
    : 'Пока нет расходов';

  return (
    `📊 Отчёт за ${month}\n\n` +
    `🟢 Доходы: ${money(current.income)}\n` +
    `🔴 Расходы: ${money(current.expenses)}\n` +
    `💰 Разница: ${money(current.net)}\n\n` +
    `📁 Расходы по категориям:\n${categoryText}\n\n` +
    `📈 Сравнение с ${previousMonth}:\n` +
    `Доходы: ${compareValues(current.income, previous.income)}\n` +
    `Расходы: ${compareValues(current.expenses, previous.expenses)}\n` +
    `Разница: ${compareValues(current.net, previous.net)}`
  );
}

async function getSheetId(): Promise<number> {
  const result = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: 'sheets.properties',
  });
  const sheet = result.data.sheets?.find((item) => item.properties?.title === sheetName);
  if (sheet?.properties?.sheetId == null) {
    throw new Error(`Лист "${sheetName}" не найден.`);
  }
  return sheet.properties.sheetId;
}

async function deleteTransaction(transaction: Transaction): Promise<void> {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          deleteDimension: {
            range: {
              sheetId: await getSheetId(),
              dimension: 'ROWS',
              startIndex: transaction.rowIndex - 1,
              endIndex: transaction.rowIndex,
            },
          },
        },
      ],
    },
  });
}

async function updateTransaction(
  transaction: Transaction,
  changes: { type?: Transaction['type']; amount?: number; category?: string; comment?: string },
): Promise<void> {
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${sheetName}'!A${transaction.rowIndex}:E${transaction.rowIndex}`,
    valueInputOption: 'RAW',
    requestBody: {
      values: [
        [
          transaction.date,
          changes.type ?? transaction.type,
          changes.amount ?? transaction.amount,
          changes.category ?? transaction.category,
          changes.comment ?? transaction.comment,
        ],
      ],
    },
  });
}

// Ограничиваем доступ личным Telegram-аккаунтом.
bot.use(async (ctx, next) => {
  if (ctx.from?.id !== userId) {
    if (ctx.chat?.type === 'private') {
      await ctx.reply('У тебя нет доступа к этому боту.');
    }
    return;
  }
  return next();
});

bot.start(async (ctx) => {
  await ctx.reply(
    'Привет! Я твой финансовый помощник 💶\n\n' +
      'Примеры операций:\n+1500 зарплата\n-20 продукты\n-12.50 кофе\n\n' +
      'Команды:\n' +
      '/balance — баланс\n/stats — статистика за месяц\n' +
      '/report — отчёт со сравнением месяцев\n/history — история с кнопками\n/cancel — отменить ввод',
    Markup.keyboard([
      ['➕ Доход', '➖ Расход'],
      ['💰 Баланс', '📊 Статистика'],
      ['🧾 История', '📈 Отчёт'],
    ]).resize(),
  );
});

bot.command('balance', async (ctx) => {
  try {
    const transactions = await getTransactions();
    await ctx.reply(`💰 Текущий баланс: ${money(calculateBalance(transactions))}`);
  } catch (error) {
    console.error(error);
    await ctx.reply('Не удалось прочитать таблицу. Проверь настройки Google Sheets.');
  }
});

bot.command('stats', async (ctx) => {
  try {
    await ctx.reply(await buildReport(getMadridMonth()));
  } catch (error) {
    console.error(error);
    await ctx.reply('Не удалось сформировать статистику.');
  }
});

bot.command('report', async (ctx) => {
  const parts = ctx.message.text.trim().split(/\s+/);
  const month = parts[1] || getMadridMonth();

  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) {
    await ctx.reply('Формат месяца: /report YYYY-MM\nНапример: /report 2026-09');
    return;
  }

  try {
    await ctx.reply(await buildReport(month));
  } catch (error) {
    console.error(error);
    await ctx.reply('Не удалось сформировать отчёт.');
  }
});

bot.command('history', async (ctx) => {
  try {
    const transactions = await getTransactions();
    const recent = transactions.slice(-5);

    if (recent.length === 0) {
      await ctx.reply('Операций пока нет.');
      return;
    }

    await ctx.reply(`🧾 Последние операции\n💰 Баланс: ${money(calculateBalance(transactions))}`);

    for (const transaction of recent) {
      const isIncome = transaction.type === 'Доход';
      const sign = isIncome ? '+' : '−';
      const indicator = isIncome ? '🟢' : '🔴';
      await ctx.reply(
        `${transaction.date}\n` +
          `${indicator} ${transaction.type}: ${sign}${money(transaction.amount)}\n` +
          `📁 ${transaction.category}\n💬 ${transaction.comment}\n`,
        Markup.inlineKeyboard([
          [
            Markup.button.callback('✏️ Изменить', `edit:${transaction.id}`),
            Markup.button.callback('🗑 Удалить', `delete:${transaction.id}`),
          ],
        ]),
      );
    }
  } catch (error) {
    console.error(error);
    await ctx.reply('Не удалось получить историю.');
  }
});

// Совместимость с командами, если захочешь использовать их без кнопок.
bot.command('edit', async (ctx) => {
  const input = ctx.message.text.trim().replace(/^\/edit\s*/i, '');
  const match = input.match(/^(\S+)\s+([+-])\s*(\d+(?:[.,]\d{1,2})?)\s*(?:€|евро?|eur)?\s*(.*)$/i);

  if (!match) {
    await ctx.reply('Формат: /edit ID -25 продукты');
    return;
  }

  const [, id, sign, rawAmount, rawComment] = match;
  const amount = Number(rawAmount.replace(',', '.'));
  if (!Number.isFinite(amount) || amount <= 0) {
    await ctx.reply('Сумма должна быть больше нуля.');
    return;
  }

  try {
    const transaction = (await getTransactions()).find((item) => item.id === id);
    if (!transaction) {
      await ctx.reply('Операция не найдена. Проверь /history.');
      return;
    }

    const comment = rawComment.trim() || 'Без комментария';
    await updateTransaction(transaction, {
      type: sign === '+' ? 'Доход' : 'Расход',
      amount,
      comment,
      category: getCategory(comment),
    });

    const updated = await getTransactions();
    await ctx.reply(`✅ Изменено. Новый баланс: ${money(calculateBalance(updated))}`);
  } catch (error) {
    console.error(error);
    await ctx.reply('Не удалось изменить операцию.');
  }
});

bot.command('delete', async (ctx) => {
  const id = ctx.message.text.trim().replace(/^\/delete\s*/i, '');
  if (!id || /\s/.test(id)) {
    await ctx.reply('Формат: /delete ID');
    return;
  }

  try {
    const transaction = (await getTransactions()).find((item) => item.id === id);
    if (!transaction) {
      await ctx.reply('Операция не найдена. Проверь /history.');
      return;
    }
    await deleteTransaction(transaction);
    const updated = await getTransactions();
    await ctx.reply(`🗑 Операция удалена. Новый баланс: ${money(calculateBalance(updated))}`);
  } catch (error) {
    console.error(error);
    await ctx.reply('Не удалось удалить операцию.');
  }
});

// Нажатие кнопки «Изменить»: выбираем поле.
bot.action(/^edit:(.+)$/, async (ctx) => {
  await ctx.answerCbQuery();
  try {
    const id = ctx.match[1];
    const transaction = (await getTransactions()).find((item) => item.id === id);
    if (!transaction) {
      await ctx.reply('Операция не найдена. Открой /history заново.');
      return;
    }

    await ctx.reply(
      `Что изменить?\n${transaction.comment}: ${money(transaction.amount)}`,
      Markup.inlineKeyboard([
        [
          Markup.button.callback('💶 Сумму / тип', `field:amount:${id}`),
          Markup.button.callback('💬 Комментарий', `field:comment:${id}`),
        ],
        [Markup.button.callback('📁 Категорию', `field:category:${id}`)],
        [Markup.button.callback('❌ Отмена', 'cancel-edit')],
      ]),
    );
  } catch (error) {
    console.error(error);
    await ctx.reply('Не удалось открыть редактирование.');
  }
});

bot.action(/^field:(amount|comment|category):(.+)$/, async (ctx) => {
  await ctx.answerCbQuery();
  const field = ctx.match[1] as EditField;
  const id = ctx.match[2];

  editStates.set(ctx.from.id, { id, field });

  const prompts: Record<EditField, string> = {
    amount: 'Введи новую сумму с типом:\n+100 — доход\n-25 — расход\n\nДля отмены отправь /cancel.',
    comment: 'Введи новый комментарий.\n\nДля отмены отправь /cancel.',
    category: 'Введи новую категорию, например «Покупки».\n\nДля отмены отправь /cancel.',
  };

  await ctx.reply(
    prompts[field],
    Markup.inlineKeyboard([[Markup.button.callback('❌ Отмена', 'cancel-edit')]]),
  );
});

bot.action(/^delete:(.+)$/, async (ctx) => {
  await ctx.answerCbQuery();
  try {
    const id = ctx.match[1];
    const transaction = (await getTransactions()).find((item) => item.id === id);
    if (!transaction) {
      await ctx.reply('Операция уже не найдена.');
      return;
    }

    await ctx.reply(
      `⚠️ Удалить операцию?\n\n${transaction.comment}\n${money(transaction.amount)}\n\nЭто действие нельзя отменить.`,
      Markup.inlineKeyboard([
        [
          Markup.button.callback('✅ Да, удалить', `confirmdelete:${id}`),
          Markup.button.callback('❌ Отмена', 'cancel-delete'),
        ],
      ]),
    );
  } catch (error) {
    console.error(error);
    await ctx.reply('Не удалось открыть подтверждение удаления.');
  }
});

bot.action(/^confirmdelete:(.+)$/, async (ctx) => {
  await ctx.answerCbQuery();
  try {
    const id = ctx.match[1];
    const transaction = (await getTransactions()).find((item) => item.id === id);
    if (!transaction) {
      await ctx.editMessageText('Операция уже не найдена.');
      return;
    }

    await deleteTransaction(transaction);
    const updated = await getTransactions();
    await ctx.editMessageText(
      `🗑 Операция удалена.\n${transaction.comment}: ${money(transaction.amount)}\n` +
        `💰 Новый баланс: ${money(calculateBalance(updated))}`,
    );
  } catch (error) {
    console.error(error);
    await ctx.reply('Не удалось удалить операцию.');
  }
});

bot.action('cancel-edit', async (ctx) => {
  editStates.delete(ctx.from.id);
  await ctx.answerCbQuery('Редактирование отменено');
  await ctx.editMessageText('Редактирование отменено.');
});

bot.action('cancel-delete', async (ctx) => {
  await ctx.answerCbQuery('Удаление отменено');
  await ctx.editMessageText('Удаление отменено.');
});

// Кнопки быстрого добавления дохода и расхода.
bot.hears('➕ Доход', async (ctx) => {
  transactionTypeStates.set(ctx.from.id, 'Доход');
  editStates.delete(ctx.from.id);
  await ctx.reply(
    '➕ Введи сумму и комментарий одним сообщением. Например: 1500 зарплата\nДля отмены отправь /cancel.',
  );
});
bot.hears('➖ Расход', async (ctx) => {
  transactionTypeStates.set(ctx.from.id, 'Расход');
  editStates.delete(ctx.from.id);
  await ctx.reply(
    '➖ Введи сумму и комментарий одним сообщением. Например: 20 продукты\nДля отмены отправь /cancel.',
  );
});

// Кнопки нижнего меню.
bot.hears('💰 Баланс', async (ctx) => {
  const transactions = await getTransactions();
  await ctx.reply(`💰 Текущий баланс: ${money(calculateBalance(transactions))}`);
});
bot.hears('📊 Статистика', async (ctx) => {
  await ctx.reply(await buildReport(getMadridMonth()));
});
bot.hears('🧾 История', async (ctx) => {
  try {
    const transactions = await getTransactions();
    const recent = transactions.slice(-5);

    if (recent.length === 0) {
      await ctx.reply('Операций пока нет.');
      return;
    }

    await ctx.reply(`🧾 Последние операции\n💰 Баланс: ${money(calculateBalance(transactions))}`);

    for (const transaction of recent) {
      const isIncome = transaction.type === 'Доход';
      const sign = isIncome ? '+' : '−';
      const indicator = isIncome ? '🟢' : '🔴';
      await ctx.reply(
        `${transaction.date}\n` +
          `${indicator} ${transaction.type}: ${sign}${money(transaction.amount)}\n` +
          `📁 ${transaction.category}\n💬 ${transaction.comment}`,
        Markup.inlineKeyboard([
          [
            Markup.button.callback('✏️ Изменить', `edit:${transaction.id}`),
            Markup.button.callback('🗑 Удалить', `delete:${transaction.id}`),
          ],
        ]),
      );
    }
  } catch (error) {
    console.error(error);
    await ctx.reply('Не удалось получить историю.');
  }
});
bot.hears('📈 Отчёт', async (ctx) => {
  await ctx.reply(await buildReport(getMadridMonth()));
});

bot.on('text', async (ctx) => {
  const text = ctx.message.text.trim();

  if (text === '/cancel') {
    editStates.delete(ctx.from.id);
    transactionTypeStates.delete(ctx.from.id);
    await ctx.reply('Действие отменено.');
    return;
  }

  // Если пользователь выбрал доход или расход, сумма вводится без знака.
  const pendingType = transactionTypeStates.get(ctx.from.id);
  if (pendingType) {
    if (text.startsWith('/')) return;
    const match = text.match(/^(\d+(?:[.,]\d{1,2})?)\s*(?:€|евро?|eur(?:o?s?)?)?\s*(.*)$/i);
    if (!match) {
      await ctx.reply('Введи сумму и комментарий, например: 20 продукты. Для отмены: /cancel.');
      return;
    }
    const amount = Number(match[1].replace(',', '.'));
    if (!Number.isFinite(amount) || amount <= 0) {
      await ctx.reply('Укажи сумму больше нуля.');
      return;
    }
    const comment = match[2].trim() || 'Без комментария';
    try {
      await addTransaction({
        id: createId(),
        date: getMadridDate(),
        type: pendingType,
        amount,
        category: getCategory(comment),
        comment,
      });
      transactionTypeStates.delete(ctx.from.id);
      const transactions = await getTransactions();
      const sign = pendingType === 'Доход' ? '+' : '−';
      await ctx.reply(
        `✅ ${pendingType} сохранён\n\n${sign}${money(amount)}\n📁 Категория: ${getCategory(comment)}\n💬 Комментарий: ${comment}\n💰 Баланс: ${money(calculateBalance(transactions))}`,
      );
    } catch (error) {
      console.error(error);
      await ctx.reply('Не удалось сохранить операцию. Проверь Google Sheets.');
    }
    return;
  }

  // Если нажата кнопка редактирования, следующее сообщение меняет выбранное поле.
  const state = editStates.get(ctx.from.id);
  if (state) {
    if (text.startsWith('/')) {
      await ctx.reply('Сначала заверши редактирование или отправь /cancel.');
      return;
    }

    try {
      const transaction = (await getTransactions()).find((item) => item.id === state.id);
      if (!transaction) {
        editStates.delete(ctx.from.id);
        await ctx.reply('Операция не найдена. Открой /history заново.');
        return;
      }

      let changes: {
        type?: Transaction['type'];
        amount?: number;
        category?: string;
        comment?: string;
      };

      if (state.field === 'amount') {
        const match = text.match(/^([+-])\s*(\d+(?:[.,]\d{1,2})?)\s*(?:€|евро?|eur)?$/i);
        if (!match) {
          await ctx.reply('Введи сумму в формате +100 или -25. Для отмены: /cancel.');
          return;
        }
        const amount = Number(match[2].replace(',', '.'));
        if (!Number.isFinite(amount) || amount <= 0) {
          await ctx.reply('Сумма должна быть больше нуля. Попробуй ещё раз или отправь /cancel.');
          return;
        }
        changes = {
          type: match[1] === '+' ? 'Доход' : 'Расход',
          amount,
        };
      } else if (state.field === 'comment') {
        changes = { comment: text, category: getCategory(text) };
      } else {
        changes = { category: text };
      }

      await updateTransaction(transaction, changes);
      editStates.delete(ctx.from.id);

      const updated = await getTransactions();
      const refreshed = updated.find((item) => item.id === state.id);
      if (!refreshed) {
        await ctx.reply('Операция обновлена, но не удалось перечитать её из таблицы.');
        return;
      }

      const sign = refreshed.type === 'Доход' ? '+' : '−';
      await ctx.reply(
        `✅ Операция обновлена!\n\n` +
          `${refreshed.type}: ${sign}${money(refreshed.amount)}\n` +
          `📁 Категория: ${refreshed.category}\n` +
          `💬 Комментарий: ${refreshed.comment}\n` +
          `💰 Новый баланс: ${money(calculateBalance(updated))}`,
        Markup.inlineKeyboard([
          [
            Markup.button.callback('✏️ Изменить ещё', `edit:${refreshed.id}`),
            Markup.button.callback('🗑 Удалить', `delete:${refreshed.id}`),
          ],
        ]),
      );
    } catch (error) {
      console.error(error);
      await ctx.reply('Не удалось сохранить изменения. Попробуй ещё раз.');
    }
    return;
  }

  // Команды не должны восприниматься как финансовые операции.
  if (text.startsWith('/')) return;

  const match = text.match(/^([+-])\s*(\d+(?:[.,]\d{1,2})?)\s*(?:€|евро?|eur(?:o?s?)?)?\s*(.*)$/i);

  if (!match) {
    await ctx.reply('Не понял операцию.\nПримеры:\n-20 продукты\n+100 зарплата');
    return;
  }

  const [, sign, rawAmount, rawComment] = match;
  const amount = Number(rawAmount.replace(',', '.'));
  if (!Number.isFinite(amount) || amount <= 0) {
    await ctx.reply('Укажи сумму больше нуля.');
    return;
  }

  const comment = rawComment.trim() || 'Без комментария';

  try {
    await addTransaction({
      id: createId(),
      date: getMadridDate(),
      type: sign === '+' ? 'Доход' : 'Расход',
      amount,
      category: getCategory(comment),
      comment,
    });

    const transactions = await getTransactions();
    await ctx.reply(
      `✅ Операция сохранена\n\n` +
        `${sign}${money(amount)}\n` +
        `📁 Категория: ${getCategory(comment)}\n` +
        `💬 Комментарий: ${comment}\n` +
        `💰 Баланс: ${money(calculateBalance(transactions))}`,
    );
  } catch (error) {
    console.error(error);
    await ctx.reply('Не удалось сохранить операцию. Проверь Google Sheets.');
  }
});

bot.catch((error) => {
  console.error('Ошибка обработки Telegram update:', error);
});

bot
  .launch()
  .then(() => console.log('Бот запущен!'))
  .catch((error) => console.error('Не удалось запустить бота:', error));

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
