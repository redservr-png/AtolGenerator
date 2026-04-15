import urllib.request, urllib.error, json, subprocess, sys

proc = subprocess.run(["git", "credential", "fill"], input=b"protocol=https\nhost=github.com\n\n", capture_output=True)
token = ""
for line in proc.stdout.decode("utf-8", errors="replace").splitlines():
    if line.startswith("password="):
        token = line[9:].strip()

if not token:
    print("No token"); sys.exit(1)

REPO = "redservr-png/AtolGenerator"
TAG  = "v1.5.2"
BODY = """## v1.5.2 — Предпросмотр чеков для всех типов операций

### Что изменилось
- **Предпросмотр чека** теперь появляется после «Сформировать документы» для **всех** типов чеков:
  - Оплаты (`ПРИХОД`)
  - Реализации (`ПРИХОД` с позициями товаров)
  - Коррекции прихода / расхода
  - Возвраты прихода / расхода
- Панель «ПРЕДПРОСМОТР» отображается в правой колонке под списком сгенерированных файлов
- Текст в стиле кассовой ленты (Courier New): шапка ИП, тип операции, позиции, итог, НДС, кассир

> В v1.5 предпросмотр был только для исправительных чеков — теперь он доступен везде.
"""

headers = {
    "Authorization": f"token {token}",
    "Accept": "application/vnd.github+json",
    "Content-Type": "application/json",
    "X-GitHub-Api-Version": "2022-11-28",
}
base = "https://api.github.com"

try:
    req = urllib.request.Request(f"{base}/repos/{REPO}/releases/tags/{TAG}", headers=headers)
    data = json.loads(urllib.request.urlopen(req).read())
    rid = data["id"]
    body = json.dumps({"name": TAG, "body": BODY, "tag_name": TAG}).encode()
    req2 = urllib.request.Request(f"{base}/repos/{REPO}/releases/{rid}", data=body, headers=headers, method="PATCH")
    print("Updated:", json.loads(urllib.request.urlopen(req2).read())["html_url"])
except urllib.error.HTTPError as e:
    if e.code == 404:
        body = json.dumps({"tag_name": TAG, "name": TAG, "body": BODY}).encode()
        req3 = urllib.request.Request(f"{base}/repos/{REPO}/releases", data=body, headers=headers, method="POST")
        print("Created:", json.loads(urllib.request.urlopen(req3).read())["html_url"])
    else:
        print(f"Error {e.code}:", e.read().decode()); sys.exit(1)
