import { DurableObject } from "cloudflare:workers";

// ─────────────────────────────────────────────────────────────────────────────
// Proxy de lectura del Boletín Oficial para el Radar de Desregulación (LyP).
//
// El BORA rechaza casi todas las conexiones desde servidores del exterior
// (GitHub Actions, Google, Cloudflare en la costa oeste de EE.UU.). Las lecturas
// hechas desde un Durable Object sí suelen pasar. Cada "lector" (?n=...) es un
// Durable Object distinto, ubicado por Cloudflare; si uno no conecta, el Radar
// prueba con el siguiente.
//
// Uso:
//   /                         → diagnóstico (lee la Primera Sección y resume)
//   /?url=<url del BORA>&n=X  → devuelve esa página del BORA, leída por el lector X
// Seguridad: solo se permiten URLs de www.boletinoficial.gob.ar, y si existe el
// secret TOKEN, el pedido debe traer el header X-Token con ese valor.
// ─────────────────────────────────────────────────────────────────────────────

const HOST_PERMITIDO = "www.boletinoficial.gob.ar";
const PRIMERA = "https://www.boletinoficial.gob.ar/seccion/primera";
const HEADERS_NAV = {
  "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
  "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
  "Accept-Language": "es-AR,es;q=0.9",
};

export class LectorBORA extends DurableObject {
  async fetch(request) {
    const destino = new URL(request.url).searchParams.get("url");
    if (destino) {
      // Modo proxy: devolver la página tal cual (mismo código y content-type).
      try {
        const r = await fetch(destino, { headers: HEADERS_NAV, cache: "no-store" });
        return new Response(r.body, {
          status: r.status,
          headers: { "content-type": r.headers.get("content-type") || "text/html; charset=utf-8" },
        });
      } catch (e) {
        return new Response(`FALLÓ: ${e}`, { status: 502 });
      }
    }
    // Modo diagnóstico.
    let ejecucion = "?";
    try {
      const t = await (await fetch("https://www.cloudflare.com/cdn-cgi/trace")).text();
      ejecucion = (t.match(/colo=(\w+)/) || [])[1] || "?";
    } catch (e) {}
    try {
      const r = await fetch(PRIMERA, { headers: HEADERS_NAV, cache: "no-store" });
      const html = await r.text();
      const avisos = new Set(html.match(/detalleAviso[^"']*/g) || []);
      return new Response(
        `Código HTTP: ${r.status}\nTamaño: ${html.length}\nNormas encontradas: ${avisos.size}\nEjecución (Durable Object): ${ejecucion}`,
        { headers: { "content-type": "text/plain; charset=utf-8" } });
    } catch (e) {
      return new Response(`FALLÓ: ${e}\nEjecución (Durable Object): ${ejecucion}`, { status: 500 });
    }
  }
}

export default {
  async fetch(request, env) {
    const u = new URL(request.url);
    if (env.TOKEN && request.headers.get("X-Token") !== env.TOKEN && u.searchParams.has("url")) {
      return new Response("No autorizado", { status: 403 });
    }
    const destino = u.searchParams.get("url");
    if (destino) {
      let d;
      try { d = new URL(destino); } catch (e) { return new Response("URL inválida", { status: 400 }); }
      if (d.protocol !== "https:" || d.hostname !== HOST_PERMITIDO) {
        return new Response("Solo se permiten URLs del Boletín Oficial", { status: 400 });
      }
    }
    const nombre = u.searchParams.get("n") || "sam-1";
    const stub = env.LECTOR.get(env.LECTOR.idFromName(nombre), { locationHint: "sam" });
    const r = await stub.fetch(request);
    if (destino) {
      const h = new Headers(r.headers);
      h.set("X-Lector", nombre);
      return new Response(r.body, { status: r.status, headers: h });
    }
    const texto = await r.text();
    return new Response(`${texto}\nEntrada: ${request.cf?.colo} | Objeto: ${nombre}`,
      { status: r.status, headers: { "content-type": "text/plain; charset=utf-8" } });
  },
};
