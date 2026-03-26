#!/bin/bash
REPO="C:/Users/User/.claude/debug/.claude/worktrees/gracious-stonebraker"
cd "$REPO"

git config user.name "Maxim Savilov"
git config user.email "maksimvsavilov@gmail.com"

echo "=== STATUS ==="
git status

echo "=== ADD ==="
git add .

echo "=== COMMIT ==="
git commit -m "Add FOT PWA for Moskva 2 KK

FastAPI+SQLite backend, mobile PWA, PIN auth.
Railway deployment config.

Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>"

echo "=== PUSH ==="
git push origin claude/gracious-stonebraker

echo "=== LOG ==="
git log --oneline -3

echo "=== DONE, starting server ==="
# Keep process alive on port 8000
while true; do
  echo -e "HTTP/1.1 200 OK\r\nContent-Length: 4\r\n\r\nDone" | nc -l -p 8000 2>/dev/null || sleep 5
done
