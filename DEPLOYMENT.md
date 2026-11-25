# Deploying AutoForm publicly (free)

This guide puts the backend on the internet and publishes the extension on Firefox Add-ons (AMO). Everything below fits in free tiers.

```
Firefox extension ──HTTPS──► FastAPI on Render (free web service)
                                 │
                                 ├──► Postgres on Neon (free, persistent)
                                 └──► Groq API (free tier) for written-answer drafts
```

## Why these pieces

| Problem with the local version | What the hosted version does |
| --- | --- |
| One shared profile (`id = 1`) for everyone | Each extension install gets an anonymous account and API key. Every query is scoped to that user. |
| SQLite file on disk | Render's free disk is wiped on every deploy and restart, so use Postgres through `DATABASE_URL`. SQLite is still the local default. |
| Ollama needs several GB of RAM or a GPU | Render's free plan has 512 MB and no GPU. Answer drafts go to any **OpenAI-compatible API**. Groq is free and fast; OpenRouter, Google Gemini, Together and OpenAI also work. Ollama still works locally and on a self-hosted GPU box. |
| No abuse protection | Per-IP limit on account creation, per-user limits on LLM drafts and CV parsing, and a cap on fields per scan. |
| CV files written to local disk | Only the extracted text is stored, and only in the user's own rows. |

## 1. Create a free Postgres database (Neon)

1. Sign up at <https://neon.tech> and create a project.
2. Copy the connection string. It looks like
   `postgresql://user:password@ep-xxx.region.aws.neon.tech/neondb?sslmode=require`.

Supabase or any other Postgres works too. Avoid Render's free Postgres for real users: it expires after 30 days.

## 2. Get a free LLM API key (Groq)

1. Sign up at <https://console.groq.com> and create an API key.
2. Check the model list (<https://console.groq.com/docs/models>). This deployment uses `openai/gpt-oss-20b`.

To use another provider, change two variables:

| Provider | `LLM_BASE_URL` | Example `LLM_MODEL` |
| --- | --- | --- |
| Groq | `https://api.groq.com/openai/v1` | `openai/gpt-oss-20b` |
| OpenRouter | `https://openrouter.ai/api/v1` | a `:free` model from their list |
| Google Gemini | `https://generativelanguage.googleapis.com/v1beta/openai` | `gemini-2.0-flash` |
| Self-hosted Ollama (GPU server) | set `LLM_PROVIDER=ollama`, `OLLAMA_BASE_URL=https://your-server` | `qwen3:4b` |

## 3. Deploy the backend on Render

1. Push this repository to GitHub.
2. In Render, go to **New → Blueprint** and choose the repository. Render reads [`render.yaml`](render.yaml).
3. When asked, paste:
   - `DATABASE_URL`: the Neon connection string
   - `LLM_API_KEY`: the Groq key
4. The Blueprint sets `LLM_MODEL=openai/gpt-oss-20b`, `LLM_REASONING_EFFORT=low`, and `LLM_MAX_TOKENS_EXTRA=200`.
5. Wait for the deploy, then open `https://autoform-shreya.onrender.com/health`. You should see `"status":"ok"` and `"llm_provider":"openai"`.

Notes:

- The free plan sleeps after about 15 minutes idle. The first request afterwards takes 30–60 s, and the popup shows a "waking up" message. A paid instance or an external uptime pinger avoids this.
- Rate limits are kept in memory, which is right for a single instance. Tune them with `GENERATE_LIMIT_PER_HOUR`, `CV_UPLOAD_LIMIT_PER_HOUR` and `REGISTER_LIMIT_PER_HOUR`.
- To deploy somewhere else (Railway, Fly.io, Koyeb, Cloud Run, a VPS), use [`backend/Dockerfile`](backend/Dockerfile) with the same environment variables. The container listens on `$PORT`.

## 4. Point the extension at your backend

The Render service is named `autoform-shreya` and uses `https://autoform-shreya.onrender.com`. If your Render URL differs, change it in **two places**:

1. `extension/background.js`: `const DEFAULT_API = 'https://<your-service>.onrender.com';`
2. `extension/manifest.json`: the matching entry in `host_permissions`.

For local development, keep the code unchanged and point an installed copy at your machine. From the extension's background console (`about:debugging` → Inspect):

```js
chrome.storage.local.set({ apiBase: 'http://127.0.0.1:8000' })
```

## 5. Test in Firefox

1. Open `about:debugging#/runtime/this-firefox` → **Load Temporary Add-on…** → pick `extension/manifest.json`.
2. Open a form page, click the AutoForm icon, then **Edit profile**. Your dashboard opens on the hosted backend.
3. Lint the package the way AMO will:

```bash
npx web-ext lint --source-dir extension
```

## 6. Publish on addons.mozilla.org

1. Package the extension: `npx web-ext build --source-dir extension`. This creates a zip in `web-ext-artifacts/`.
2. Go to <https://addons.mozilla.org/developers/> → **Submit a New Add-on** → **On this site** (listed and free).
3. Upload the zip. The source is not minified, so no separate source upload is needed.
4. Add a **privacy policy**. You can start from [`PRIVACY.md`](PRIVACY.md); fill in your contact details and host it somewhere public, for example as a GitHub page. AMO requires one because profile data and form labels are sent to your server and the LLM provider.
5. The manifest already declares `data_collection_permissions` (`personallyIdentifyingInfo`, `websiteContent`), which AMO requires for new add-ons.
6. The add-on ID in `browser_specific_settings.gecko.id` is permanent once published, so keep it.

Optional before submitting: add icons (`"icons": {"48": "icon-48.png", "96": "icon-96.png"}` in the manifest) so the listing does not show a generic puzzle piece.

## Running locally (unchanged workflow)

```bash
cd backend
python -m venv .venv && . .venv/bin/activate     # Windows: .venv\Scripts\activate
pip install -r requirements-dev.txt
python training/train.py
uvicorn app.main:app --host 127.0.0.1 --port 8000
pytest                                         # TEST_DATABASE_URL=postgresql://... to test against Postgres
```

With no `LLM_API_KEY` set, generation uses local Ollama exactly as before.
