export function add(a: number, b: number): number {
  return a + b;
}

function contentTypeForFile(path: string): string {
  if (path.endsWith(".html")) return "text/html; charset=utf-8";
  if (path.endsWith(".js")) return "application/javascript; charset=utf-8";
  if (path.endsWith(".css")) return "text/css; charset=utf-8";
  if (path.endsWith(".json")) return "application/json; charset=utf-8";
  if (path.endsWith(".svg")) return "image/svg+xml";
  return "application/octet-stream";
}

async function handleRequest(request: Request): Promise<Response> {
  const url = new URL(request.url);
  const pathname = url.pathname === "/" ? "/index.html" : url.pathname;
  const normalizedPath = pathname.replace(/\.{2,}/g, "").replace(/^\//, "");

  try {
    const file = await Deno.readFile(normalizedPath);
    return new Response(file, {
      status: 200,
      headers: { "content-type": contentTypeForFile(normalizedPath) },
    });
  } catch {
    return new Response("Not found", { status: 404 });
  }
}

if (import.meta.main) {
  const port = Number(Deno.env.get("PORT") || 8000);
  console.log(`Serving SPA at http://localhost:${port}`);
  await Deno.serve({ port }, handleRequest);
}
