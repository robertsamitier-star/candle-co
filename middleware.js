const COOKIE_NAME = "__candle_auth";

export default async function middleware(request) {
  const url = new URL(request.url);

  // Skip auth for API routes and robots.txt
  if (url.pathname.startsWith("/api/") || url.pathname === "/robots.txt") {
    return;
  }

  // Skip auth on preview deployments
  if (process.env.VERCEL_ENV === "preview") {
    return;
  }

  // Handle auth form submission
  if (url.pathname === "/__auth" && request.method === "POST") {
    const formData = await request.formData();
    const password = formData.get("password");
    const sitePassword = process.env.SITE_PASSWORD;

    if (!sitePassword || password !== sitePassword) {
      return new Response(loginPage("Wrong password. Try again."), {
        status: 401,
        headers: { "Content-Type": "text/html; charset=utf-8" },
      });
    }

    // Create a simple token from the password hash
    const token = await hashToken(sitePassword);

    return new Response(null, {
      status: 302,
      headers: {
        Location: "/",
        "Set-Cookie": `${COOKIE_NAME}=${token}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=${30 * 24 * 60 * 60}`,
      },
    });
  }

  // Check auth cookie
  const cookie = request.headers.get("cookie") || "";
  const match = cookie.match(new RegExp(`${COOKIE_NAME}=([^;]+)`));

  if (match) {
    const sitePassword = process.env.SITE_PASSWORD;
    if (sitePassword) {
      const expected = await hashToken(sitePassword);
      if (match[1] === expected) {
        return;
      }
    }
  }

  // Not authenticated — show login page
  return new Response(loginPage(), {
    status: 401,
    headers: { "Content-Type": "text/html; charset=utf-8" },
  });
}

async function hashToken(value) {
  const encoder = new TextEncoder();
  const data = encoder.encode(value + "__candle_salt_2026");
  const hash = await crypto.subtle.digest("SHA-256", data);
  return Array.from(new Uint8Array(hash))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}

function loginPage(error = "") {
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="robots" content="noindex, nofollow">
<title>Amber &amp; Essence — Access</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: 'Inter', sans-serif;
    background: #0a0a0f;
    color: #e0e0e0;
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .gate {
    background: #12121a;
    border: 1px solid rgba(255,255,255,0.06);
    border-radius: 16px;
    padding: 48px 40px;
    max-width: 400px;
    width: 90%;
    text-align: center;
  }
  .gate h1 {
    font-size: 1.4rem;
    font-weight: 600;
    margin-bottom: 8px;
    color: #C08B5C;
  }
  .gate p {
    font-size: 0.85rem;
    color: rgba(255,255,255,0.5);
    margin-bottom: 32px;
  }
  .gate input[type="password"] {
    width: 100%;
    padding: 14px 16px;
    background: #0a0a0f;
    border: 1px solid rgba(255,255,255,0.1);
    border-radius: 10px;
    color: #e0e0e0;
    font-size: 0.95rem;
    font-family: 'Inter', sans-serif;
    outline: none;
    transition: border-color 0.2s;
  }
  .gate input[type="password"]:focus {
    border-color: #C08B5C;
  }
  .gate button {
    width: 100%;
    margin-top: 16px;
    padding: 14px;
    background: #C08B5C;
    color: #0a0a0f;
    border: none;
    border-radius: 10px;
    font-size: 0.9rem;
    font-weight: 600;
    font-family: 'Inter', sans-serif;
    cursor: pointer;
    transition: opacity 0.2s;
  }
  .gate button:hover { opacity: 0.85; }
  .error {
    color: #ef4444;
    font-size: 0.82rem;
    margin-top: 12px;
  }
</style>
</head>
<body>
  <form class="gate" method="POST" action="/__auth">
    <h1>Amber &amp; Essence</h1>
    <p>Enter password to continue</p>
    <input type="password" name="password" placeholder="Password" autofocus required>
    <button type="submit">Enter</button>
    ${error ? `<div class="error">${error}</div>` : ""}
  </form>
</body>
</html>`;
}

export const config = {
  matcher: ["/((?!_vercel|_next/static|favicon.ico).*)"],
};
