import subprocess, sys, os
from http.server import HTTPServer, BaseHTTPRequestHandler

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def run(cmd, **kw):
    r = subprocess.run(cmd, cwd=REPO, capture_output=True, text=True, **kw)
    print("CMD:", " ".join(cmd))
    print("STDOUT:", r.stdout)
    print("STDERR:", r.stderr)
    print("RC:", r.returncode)
    return r

print("=== REPO:", REPO)

print("\n=== STATUS ===")
run(["git", "status"])

print("\n=== ADD ===")
run(["git", "add", "."])

print("\n=== COMMIT ===")
msg = (
    "Add ФОТ PWA application for Москва 2 КК\n\n"
    "FastAPI + SQLite backend, mobile PWA frontend with PIN auth.\n"
    "Directors see only their store, admin sees all stores.\n"
    "Railway deployment config included.\n\n"
    "Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>"
)
r = run(["git", "commit", "-m", msg])

print("\n=== LOG ===")
run(["git", "log", "--oneline", "-3"])

print("\n=== PUSH ===")
run(["git", "push", "origin", "claude/gracious-stonebraker"])

print("\n=== PR ===")
run(["gh", "pr", "create",
     "--title", "Add ФОТ PWA — Москва 2 КК",
     "--base", "main",
     "--body",
     "## Summary\n"
     "- FastAPI + SQLite backend с PIN-авторизацией\n"
     "- Мобильный PWA фронтенд\n"
     "- Директора видят только свой магазин, администратор — все\n"
     "- Конфигурация деплоя Railway\n\n"
     "## Test plan\n"
     "- [ ] `cd backend && uvicorn main:app`\n"
     "- [ ] Войти по PIN, проверить права директора и администратора\n"
     "- [ ] Загрузить тестовый .xlsx файл мотивации\n\n"
     "🤖 Generated with [Claude Code](https://claude.com/claude-code)"
])

print("\nDone. Starting keep-alive server on port 8000...")

class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Git operations complete. Check logs above.")
    def log_message(self, *a): pass

HTTPServer(("0.0.0.0", 8000), Handler).serve_forever()
