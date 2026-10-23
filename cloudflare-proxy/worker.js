import { DurableObject } from "cloudflare:workers";

// El Durable Object corre en la región pedida (Sudamérica) sin importar
// desde dónde llegue el pedido (GitHub en EE.UU., tu navegador, etc.).
export class LectorBORA extends DurableObject {
  async fetch(request) {
    let ejecucion = "?";
    try {
      const t = await (await fetch("https://www.cloudflare.com/cdn-cgi/trace")).text();
      ejecucion = (t.match(/colo=(\w+)/) || [])[1] || "?";
    } catch (e) {}
    try {
      const r = await fetch("https://www.boletinoficial.gob.ar/seccion/primera", {
        headers: {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
          "Accept-Language": "es-AR,es;q=0.9",
        },
      });
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
    // ?n=otro-nombre crea un objeto nuevo (útil si el primero quedó mal ubicado).
    const nombre = new URL(request.url).searchParams.get("n") || "sam-1";
    const stub = env.LECTOR.get(env.LECTOR.idFromName(nombre), { locationHint: "sam" });
    const r = await stub.fetch(request);
    const texto = await r.text();
    return new Response(`${texto}\nEntrada: ${request.cf?.colo} | Objeto: ${nombre}`,
      { status: r.status, headers: { "content-type": "text/plain; charset=utf-8" } });
  },
};
