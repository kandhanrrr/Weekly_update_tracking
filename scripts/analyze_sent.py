import json
import re
from pathlib import Path

BODY_FILE = (
    r"c:\Users\kandhanr\AppData\Roaming\Code\User\workspaceStorage"
    r"\60aa117833b3e2b7932e0883ed2edba2\GitHub.copilot-chat\chat-session-resources"
    r"\0c53dead-b1f2-4b19-982f-723cd8466184"
    r"\toolu_bdrk_013GinzhNsqnq3h3AmkwBAuD__vscode-1785224037330\content.json"
)

GREETING = re.compile(
    r"(?:Hi|Hello|Dear)\s+([A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+)?)\s*[,\n]",
    re.MULTILINE,
)

data = json.loads(Path(BODY_FILE).read_text(encoding="utf-8"))

for e in data.get("emails", []):
    subj = e.get("subject", "")
    body_raw = e.get("body") or ""
    body = (body_raw.get("content", "") if isinstance(body_raw, dict) else body_raw) or e.get("preview", "")
    body_clean = re.sub(r"<[^>]+>", " ", body)
    body_clean = re.sub(r"\s+", " ", body_clean).strip()[:400]
    to_list = e.get("to", [])
    to_addrs = [
        x.get("emailAddress", {}).get("address", "") if isinstance(x, dict) else str(x)
        for x in to_list[:3]
    ]
    date = str(e.get("sentDateTime", ""))[:10]
    m = GREETING.search(body_clean[:300])
    greeting = m.group(1) if m else ""
    print(f"DATE: {date}")
    print(f"SUBJ: {subj[:90]}")
    print(f"TO  : {to_addrs}")
    print(f"GREETING: {greeting}")
    print(f"BODY: {body_clean[:350]}")
    print("-" * 80)
