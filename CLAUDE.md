# lele_auto

<!-- FILL IN: 2-3 sentences on what this project is and who it's for. -->
TODO: describe this project in 2-3 sentences.

## Stack
TODO: languages, frameworks, key dependencies.

## Commands
- Local preview: `python3 -m http.server 8147 --bind 0.0.0.0 --directory site`
  -> http://192.168.1.100:8147
- Build / test / run: TODO

## Critical conventions
- Web content lives in `site/`. Ship a favicon from day one.
- Building or upgrading the site → invoke the **site-craft** skill FIRST, then pull from the design/GSAP skill map in `docs/WEB_DESIGN.md`.
- Code, comments and commit messages in English; chat with Pascal in French.
- TODO: project-specific rules worth loading every message.

## How to talk to Pascal
- No flattering preamble ("Bonne idée", "Excellente question", "Tu as raison",
  "Je comprends"). Go straight to the answer.
- No emotional validation, no compliments on requests.
- Concise and factual. If asked X, answer X — no padding around it.
- Do not recap what Pascal just said before answering.
- If you disagree, or an approach is bad, say so directly, without softening it.

## Context docs — read on demand (keep THIS file minimal: it loads every message)
- `docs/ARCHITECTURE.md` — structure, tech choices + rationale. Read before structural changes.
- `docs/DECISIONS.md` — decision log (date, decision, reason). Read to learn why things are as they are.
- `docs/ROADMAP.md` — current state, in-progress, next, done. Read to know what to work on.
- `docs/CONVENTIONS.md` — detailed conventions, patterns, pitfalls specific to this project. Read before writing code.
- `docs/WEB_DESIGN.md` — building/upgrading the site: start with the **site-craft** skill, then the full design & GSAP skill map. Read before any site work.
- `SHARED_FILES.md` — documents Pascal drops arrive in `./share/`.
- Knowledge graph: on a large project, `/graphify` can build one on demand (expensive — never build unprompted). If `graphify-out/graph.json` exists, query it before broad structural greps.

## Permanent instruction
After each significant work session, update `docs/ROADMAP.md` (and `docs/DECISIONS.md` if relevant) so no context is lost on `/clear`.
