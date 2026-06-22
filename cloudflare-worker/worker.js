/**
 * Cloudflare Worker — приём отзывов с сайта и пересылка в Telegram-группу.
 *
 * Сайт (index.html) отправляет на адрес этого Worker POST с JSON:
 *   - отзыв:           { name, text }
 *   - заявка на файлы: { type: 'lead', email, magnet: 'guides' | 'book' }
 * Worker добавляет токен бота (он хранится скрытно в переменных Worker, в коде его нет)
 * и пересылает сообщение в группу через Telegram Bot API.
 *
 * Нужные переменные окружения (Settings → Variables and Secrets):
 *   BOT_TOKEN    — токен бота от @BotFather (как секрет)
 *   CHAT_ID      — id вашей Telegram-группы (например, -1001234567890)
 *   ALLOW_ORIGIN — (необязательно) разрешённый источник, по умолчанию '*'.
 *                  Рекомендуется указать адрес сайта, напр. https://urokisestram.com
 *
 * Полная инструкция по настройке — в README.md рядом с этим файлом.
 */

export default {
  async fetch(request, env) {
    const allowOrigin = env.ALLOW_ORIGIN || '*';

    // Префлайт CORS
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: corsHeaders(allowOrigin) });
    }

    if (request.method !== 'POST') {
      return json({ error: 'Method not allowed' }, 405, allowOrigin);
    }

    let data;
    try {
      data = await request.json();
    } catch (e) {
      return json({ error: 'Invalid JSON' }, 400, allowOrigin);
    }

    if (!env.BOT_TOKEN || !env.CHAT_ID) {
      return json({ error: 'Worker is not configured' }, 500, allowOrigin);
    }

    let message;

    if (data.type === 'lead') {
      // Заявка на бесплатные материалы (сбор контактов с форм «получить гайды/книгу»).
      const name = String(data.name || '').trim().slice(0, 100);
      const email = String(data.email || '').trim().slice(0, 150);
      const phone = String(data.phone || '').trim().slice(0, 40);
      const tg = String(data.tg || '').trim().slice(0, 60);
      const magnet = String(data.magnet || '').trim().slice(0, 60);
      if (!email && !phone) {
        return json({ error: 'No contact provided' }, 400, allowOrigin);
      }
      const what = magnet === 'guides' ? 'Гайды по планированию'
        : magnet === 'book' ? 'Справочник постящейся'
        : (magnet || 'материалы');
      message =
        '📩 Новая заявка на материалы\n\n' +
        'Что: ' + what + '\n' +
        'Имя: ' + (name || '—') + '\n' +
        'Почта: ' + (email || '—') + '\n' +
        'Телефон: ' + (phone || '—') + '\n' +
        'Telegram: ' + (tg || '—');
    } else {
      // Отзыв с сайта.
      const name = String(data.name || '').trim().slice(0, 100);
      const text = String(data.text || '').trim().slice(0, 2000);
      if (!text) {
        return json({ error: 'Empty review' }, 400, allowOrigin);
      }
      message =
        '🌸 Новый отзыв с сайта\n\n' +
        'Имя: ' + (name || 'аноним') + '\n\n' +
        text;
    }

    const tgUrl = 'https://api.telegram.org/bot' + env.BOT_TOKEN + '/sendMessage';
    let tgRes;
    try {
      tgRes = await fetch(tgUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          chat_id: env.CHAT_ID,
          text: message,
          disable_web_page_preview: true,
        }),
      });
    } catch (e) {
      return json({ error: 'Telegram request failed' }, 502, allowOrigin);
    }

    if (!tgRes.ok) {
      return json({ error: 'Telegram rejected the message' }, 502, allowOrigin);
    }

    return json({ ok: true }, 200, allowOrigin);
  },
};

function corsHeaders(origin) {
  return {
    'Access-Control-Allow-Origin': origin,
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Max-Age': '86400',
  };
}

function json(body, status, origin) {
  return new Response(JSON.stringify(body), {
    status: status,
    headers: {
      'Content-Type': 'application/json',
      ...corsHeaders(origin),
    },
  });
}
