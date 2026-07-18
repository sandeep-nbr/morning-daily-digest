# Morning Daily Digest — NBR Group Real Estate

Automated daily research + compilation + email delivery of a "Daily Morning Real
Estate Digest" for NBR Group senior management.

## What it does

Every morning at **9:00 AM IST** (`30 3 * * *` UTC), a scheduled Routine wakes this
Claude Code session and has it:

1. Research current real estate news via web search, drawing on ET Realty, Moneycontrol
   Real Estate/Property, Housing.com, Magicbricks, RealtyNXT and Realty Plus (and
   corroborating business media where a publisher blocks automated access).
2. Compile a `.docx` digest structured in three layers, each covering Residential
   (primary emphasis), Commercial/Industrial/Leasing, and RERA & CREDAI updates:
   - **Layer 1** — Bengaluru & Karnataka
   - **Layer 2** — South India
   - **Layer 3** — National
3. Email the `.docx` to `murli@nbrgroup.in`, `sandeep@nbrgroup.in`, `supraja@nbrgroup.in`
   via the connected Gmail or Microsoft Outlook account (Zapier MCP connector).

## Repo contents

- `scripts/build_digest.js` — reference `docx`-generation script from the first run
  (2026-07-18). Each daily run writes its own script from scratch following the same
  structure/gotchas documented in the `docx` skill; this is kept as a working example,
  not something the daily run re-executes verbatim.
- `digests/` — archive of generated `.docx` digests.

## Operational notes / known limitations

- **Source access**: `moneycontrol.com` blocks the automated fetch/search crawler
  entirely, and direct fetches to `realty.economictimes.indiatimes.com`,
  `realtynxt.com`, and `rprealtyplus.com` frequently return HTTP 403. Content is
  therefore compiled via broad web search (which does index these outlets, just not
  on every run) plus corroborating business-media coverage, with sources cited per
  bullet. This is a best-effort compilation, not a literal scrape of the six named
  portals every day — figures should be independently verified before use in decisions.
- **Email account**: the Gmail connection required re-authentication as of 2026-07-18;
  the digest is currently delivered via the Microsoft Outlook connector instead. If you
  want delivery to switch back to Gmail, re-authenticate it in the Zapier connector
  settings.
- **Attachment delivery**: the Zapier email-send tools' `file` param cannot read a local
  sandbox file path — passing one silently produces a corrupted few-hundred-byte `.txt`
  attachment with no error (discovered 2026-07-18). The working method: commit the day's
  `.docx` to `digests/` on this branch (this repo is public), then pass its
  `raw.githubusercontent.com` URL as the attachment source. Always verify the sent
  message's attachment `contentType`/size afterward (via `microsoft_outlook_find_emails`
  or the Gmail equivalent) before treating the send as successful.
- **Environment**: the sandbox this runs in only ships `libreoffice-core` and no
  `docx` npm package or `pandoc`/`poppler-utils` by default. Each run installs what it
  needs (`npm install docx`, `apt-get install -y libreoffice-writer pandoc poppler-utils`)
  before building/verifying the document.
- **Schedule reliability**: the Routine is bound to this same Claude Code session
  (so it keeps the Zapier connector). If a scheduled run doesn't produce an email,
  check the Routine's history via the Claude Code Remote `list_triggers`/execution
  history, since a fresh session spawned without this session's connectors would not
  be able to send mail.
