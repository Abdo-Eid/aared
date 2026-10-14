/** Shared review chrome only: each format keeps its own cards and evidence. */
export function escapeHtml(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

export function reviewPage(title: string, content: string): string {
  return `<!doctype html>
<meta charset="utf-8"><title>${escapeHtml(title)}</title>
<style>
body{margin:24px;background:#f4f4f4;color:#111;font:16px/1.4 system-ui,sans-serif}
h1{margin:0 0 8px}p.summary{margin:0 0 24px;color:#444}
.pair{display:flex;gap:16px;margin-bottom:24px;align-items:flex-start}
.pair>figure{flex:1;min-width:0}figure{margin:0;background:#fff;padding:12px}
figcaption{font-weight:700;margin-bottom:8px}img{display:block;width:100%;height:auto}
@media(max-width:900px){.pair{flex-direction:column}.pair>figure{box-sizing:border-box;width:100%}}
</style>
<h1>${escapeHtml(title)}</h1>
${content}`;
}
