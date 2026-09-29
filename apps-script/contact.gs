/**
 * Contact form backend for yanxia.art.
 *
 * Deployed as a standalone Apps Script project ("yanxia.art contact"):
 * Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 * Each message becomes a row in the sheet and an email to the sheet owner.
 * The owner's address never reaches the browser.
 */
const SHEET_ID = "1Cs5NaS6kbnVQoMnL6ZrchGvIXpoPDhAMrzH2OX4h0tg";  // "yanxia.art messages"
const MAX_PER_HOUR = 20;   // crude flood guard across all senders

function doPost(e) {
  const p = (e && e.parameter) || {};
  // honeypot: real visitors never see or fill this field
  if (p.website) return reply_(true);

  const name = clip_(p.name, 100);
  const contact = clip_(p.contact, 200);
  const message = clip_(p.message, 5000);
  if (!message) return reply_(false, "empty");

  const cache = CacheService.getScriptCache();
  const count = Number(cache.get("count") || 0);
  if (count >= MAX_PER_HOUR) return reply_(false, "busy");
  cache.put("count", String(count + 1), 3600);

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  if (sheet.getLastRow() === 0) sheet.appendRow(["时间", "称呼", "联系方式", "留言", "语言"]);
  sheet.appendRow([new Date(), name, contact, message, clip_(p.lang, 5)]);

  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: "yanxia.art 留言" + (name ? "：" + name : ""),
    body: "From: " + (name || "(no name)") + "\nContact: " + (contact || "(none)") + "\n\n" + message,
    replyTo: /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(contact) ? contact : undefined,
  });
  return reply_(true);
}

function clip_(v, n) { return String(v || "").trim().slice(0, n); }

function reply_(ok, error) {
  return ContentService.createTextOutput(JSON.stringify({ ok: ok, error: error || null }))
    .setMimeType(ContentService.MimeType.JSON);
}
