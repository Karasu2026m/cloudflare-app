/**
 * Cloudflare Workers - GAS API プロキシ
 * 
 * フロントエンドからのリクエストを受け取り、
 * GAS URLとAPIキーを秘匿したままGASに転送する。
 */

const CORS_HEADERS = {
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Max-Age': '86400',
};

function corsHeaders(origin, env) {
  const allowed = env.ALLOWED_ORIGIN || '*';
  return {
    ...CORS_HEADERS,
    'Access-Control-Allow-Origin': allowed === '*' ? '*' : (origin === allowed ? allowed : ''),
    'Content-Type': 'application/json',
  };
}

export default {
  async fetch(request, env, ctx) {
    const origin = request.headers.get('Origin') || '';

    // CORS preflight
    if (request.method === 'OPTIONS') {
      return new Response(null, {
        status: 204,
        headers: corsHeaders(origin, env),
      });
    }

    // POSTのみ受付
    if (request.method !== 'POST') {
      return new Response(JSON.stringify({ success: false, message: 'Method not allowed' }), {
        status: 405,
        headers: corsHeaders(origin, env),
      });
    }

    // Origin検証（設定されている場合）
    if (env.ALLOWED_ORIGIN && env.ALLOWED_ORIGIN !== '*') {
      if (origin !== env.ALLOWED_ORIGIN) {
        return new Response(JSON.stringify({ success: false, message: 'Origin not allowed' }), {
          status: 403,
          headers: corsHeaders(origin, env),
        });
      }
    }

    try {
      // フロントエンドからのリクエストボディを取得
      const body = await request.json();

      // APIキーを注入（フロントエンドには露出しない）
      body.apiKey = env.GAS_API_KEY;

      // GASにリクエスト転送
      const gasResponse = await fetch(env.GAS_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
        redirect: 'follow', // GASはリダイレクトするため必要
      });

      const result = await gasResponse.text();

      return new Response(result, {
        status: 200,
        headers: corsHeaders(origin, env),
      });
    } catch (err) {
      return new Response(JSON.stringify({ success: false, message: 'Proxy error: ' + err.message }), {
        status: 500,
        headers: corsHeaders(origin, env),
      });
    }
  },
};
