/**
 * Cloudflare Pages Functions - /api エンドポイント
 * 
 * フロントエンドの fetch('/api', ...) リクエストを受け取り、
 * GASにAPIキーを付与して転送するプロキシ。
 * 
 * 環境変数（Cloudflare Pages Settings → Environment Variables で設定）:
 *   GAS_URL: GASデプロイURL
 *   GAS_API_KEY: APIキー
 */
export async function onRequestPost(context) {
  const { request, env } = context;

  const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };

  try {
    const body = await request.json();

    // APIキーを注入
    body.apiKey = env.GAS_API_KEY;

    // GASに転送
    const gasRes = await fetch(env.GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
      redirect: 'follow',
    });

    const result = await gasRes.text();

    return new Response(result, {
      status: 200,
      headers: corsHeaders,
    });
  } catch (err) {
    return new Response(JSON.stringify({ success: false, message: 'Proxy error: ' + err.message }), {
      status: 500,
      headers: corsHeaders,
    });
  }
}

export async function onRequestOptions() {
  return new Response(null, {
    status: 204,
    headers: {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Access-Control-Max-Age': '86400',
    },
  });
}
