// Proxies BNR's daily reference exchange rate feed (server-side, avoids browser CORS).
exports.handler = async function () {
  try {
    const res = await fetch("https://www.bnr.ro/nbrfxrates.xml")
    if (!res.ok) throw new Error(`BNR feed responded ${res.status}`)
    const xml = await res.text()
    const rateMatch = xml.match(/<Rate currency="EUR"[^>]*>([\d.]+)<\/Rate>/)
    const dateMatch = xml.match(/<Cube date="([\d-]+)"/)
    const rate = rateMatch ? parseFloat(rateMatch[1]) : null
    if (!rate) throw new Error("EUR rate not found in BNR feed")
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json", "Cache-Control": "public, max-age=3600" },
      body: JSON.stringify({ rate, date: dateMatch ? dateMatch[1] : null, source: "bnr.ro" }),
    }
  } catch (err) {
    return {
      statusCode: 502,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ error: String(err && err.message || err) }),
    }
  }
}
