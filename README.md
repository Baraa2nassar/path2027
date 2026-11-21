# My 2027 Path

A phone page for the 2027 plan. Your plan, monthly goals, dreams and weekly ticks all live in your
**Gym Log** Google Sheet (new tabs: **Plan**, **Months**, **Dreams**, **2027 Path**). The page itself
has no personal content — anyone opening it without the Sheet URL + key sees an empty shell.

## 1. Update the Gym Log script (once, on a computer)
1. Open the Gym Log sheet → **Extensions → Apps Script**.
2. Delete everything and paste the new `google-sheet/Code.gs` (Gym Log keeps working exactly the same).
3. Save, then **Deploy → Manage deployments → ✏ Edit → Version: New version → Deploy**. The URL stays the same.

## 2. Add your plan to the Sheet (once)
Add three tabs — **Plan**, **Months**, **Dreams** — and paste in your content. See
`path2027-seed.md` (not in the repo; the assistant will give it to you) for a ready-made plan to paste.

- **Plan** — columns: `week`, `task`. `week` is the Monday date as `YYYY-MM-DD`, kept as **plain text**.
- **Months** — columns: `month`, `quran`, `course`, `film`, `body`, `family`, `money`. `month` is `YYYY-MM`, plain text.
- **Dreams** — columns: `group`, `icon`, `title`, `text`, `main`. Valid icons:
  `quran, course, gym, film, family, money, pen, heart, people, moon, hands, scroll`. Put `yes` in `main` for the hero card.

## 3. Put the page online (GitHub Pages)
Already done at `https://baraa2nassar.github.io/path2027/`. To update the page, commit and push.

## 4. Connect on your phone
1. Open the page, scroll to **Google Sheet**, paste the Gym Log URL and key, tap **Connect**.
2. Safari → Share → **Add to Home Screen** so it opens like an app.

Offline ticks wait on the phone and upload when you're back online. The sheet is the master copy.
