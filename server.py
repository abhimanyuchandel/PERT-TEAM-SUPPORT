#!/usr/bin/env python3
import json
import mimetypes
import os
import secrets
import time
import urllib.error
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

HOST = os.environ.get("HOST", "127.0.0.1")
PORT = int(os.environ.get("PORT", "3000"))
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "").strip()
TOKEN_TTL_SECONDS = max(60, int(os.environ.get("TOKEN_TTL_SECONDS", "1800")))
ALLOWED_MODELS = {"gpt-4.1-mini", "gpt-4.1"}
WORKSPACE_ROOT = Path.cwd().resolve()

TOKENS = {}


def now() -> int:
    return int(time.time())


def cleanup_tokens() -> None:
    current = now()
    expired = [token for token, exp in TOKENS.items() if exp <= current]
    for token in expired:
        TOKENS.pop(token, None)


def issue_token() -> str:
    token = secrets.token_hex(24)
    TOKENS[token] = now() + TOKEN_TTL_SECONDS
    return token


def validate_token(token: str) -> bool:
    cleanup_tokens()
    if not token:
        return False
    exp = TOKENS.get(token)
    if exp is None or exp <= now():
        TOKENS.pop(token, None)
        return False
    TOKENS[token] = now() + TOKEN_TTL_SECONDS
    return True


def parse_bearer(auth_header: str) -> str:
    if not auth_header:
        return ""
    prefix = "bearer "
    if auth_header.lower().startswith(prefix):
        return auth_header[len(prefix):].strip()
    return ""


def extract_response_text(payload: dict) -> str:
    if not isinstance(payload, dict):
        return ""
    output_text = payload.get("output_text")
    if isinstance(output_text, str):
        return output_text
    if isinstance(output_text, list):
        return "\n".join([str(item) for item in output_text])
    output = payload.get("output")
    if isinstance(output, list):
        chunks = []
        for item in output:
            content = item.get("content") if isinstance(item, dict) else None
            if not isinstance(content, list):
                continue
            for part in content:
                if isinstance(part, dict) and isinstance(part.get("text"), str):
                    chunks.append(part["text"])
        if chunks:
            return "\n".join(chunks)
    choices = payload.get("choices")
    if isinstance(choices, list) and choices:
        message = choices[0].get("message") if isinstance(choices[0], dict) else None
        content = message.get("content") if isinstance(message, dict) else None
        if isinstance(content, str):
            return content
    return ""


def parse_bullets(text: str):
    lines = [line.strip() for line in (text or "").splitlines() if line.strip()]
    bullets = []
    for line in lines:
        if line.startswith(("-", "*", "•")):
            line = line[1:].strip()
        elif line[:2].isdigit() and line[2:3] == ".":
            line = line[3:].strip()
        bullets.append(line)
    if not bullets and text.strip():
        bullets = [text.strip()]
    return bullets[:6]


class Handler(BaseHTTPRequestHandler):
    server_version = "PERTServer/1.0"

    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.send_header("Access-Control-Allow-Methods", "GET,POST,OPTIONS")
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def _send_json(self, status_code: int, payload: dict):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_text(self, status_code: int, message: str):
        body = message.encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self):
        length = int(self.headers.get("Content-Length", "0"))
        if length > 2 * 1024 * 1024:
            raise ValueError("Request body too large")
        raw = self.rfile.read(length) if length else b"{}"
        try:
            return json.loads(raw.decode("utf-8"))
        except json.JSONDecodeError as exc:
            raise ValueError("Invalid JSON body") from exc

    def _safe_file_path(self):
        req_path = self.path.split("?", 1)[0]
        if req_path == "/":
            req_path = "/index.html"
        candidate = (WORKSPACE_ROOT / req_path.lstrip("/")).resolve()
        try:
            candidate.relative_to(WORKSPACE_ROOT)
        except ValueError:
            return None
        return candidate

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Allow", "GET,POST,OPTIONS")
        self.end_headers()

    def do_POST(self):
        route = self.path.split("?", 1)[0]
        if route == "/api/token":
            self.handle_token()
            return
        if route == "/api/ai-addendum":
            self.handle_ai_addendum()
            return
        self._send_text(404, "Not found")

    def do_GET(self):
        route = self.path.split("?", 1)[0]
        if route == "/api/health":
            self._send_json(200, {"ok": True, "openaiConfigured": bool(OPENAI_API_KEY)})
            return
        self.serve_static(send_body=True)

    def do_HEAD(self):
        self.serve_static(send_body=False)

    def handle_token(self):
        if not OPENAI_API_KEY:
            self._send_json(503, {"error": "OPENAI_API_KEY is not configured on server."})
            return
        token = issue_token()
        self._send_json(200, {"token": token, "expiresInSeconds": TOKEN_TTL_SECONDS})

    def handle_ai_addendum(self):
        if not OPENAI_API_KEY:
            self._send_json(503, {"error": "OPENAI_API_KEY is not configured on server."})
            return

        auth_token = parse_bearer(self.headers.get("Authorization", ""))
        if not validate_token(auth_token):
            self._send_json(401, {"error": "Invalid or expired session token."})
            return

        try:
            body = self._read_json()
        except ValueError as exc:
            self._send_json(400, {"error": str(exc)})
            return

        model = body.get("model")
        if model not in ALLOWED_MODELS:
            model = "gpt-4.1-mini"
        narrative = (body.get("narrative") or "").strip()
        profile = body.get("profile") if isinstance(body.get("profile"), dict) else {}
        if not narrative:
            self._send_json(400, {"error": "Narrative is required."})
            return

        prompt = (
            "Structured profile:\n"
            f"- Category: {profile.get('category', 'n/a')}\n"
            f"- Profile: {profile.get('descriptor', 'n/a')}\n"
            f"- Diagnosis status: {profile.get('diagnosisStatus', 'n/a')}\n"
            f"- Hemodynamics: persistent hypotension={'yes' if profile.get('hemodynamics', {}).get('persistentHypotension') else 'no'}, transient hypotension={'yes' if profile.get('hemodynamics', {}).get('transientHypotension') else 'no'}, MAP={profile.get('hemodynamics', {}).get('map', 'n/a')}, lactate={profile.get('hemodynamics', {}).get('lactate', 'n/a')}, vasopressors={profile.get('hemodynamics', {}).get('vasopressors', 'n/a')}\n"
            f"- Respiratory support: {profile.get('respiratory', {}).get('oxygenSupport', 'n/a')}, RR={profile.get('respiratory', {}).get('rr', 'n/a')}\n"
            f"- Contraindications: anticoag={'yes' if profile.get('contraindications', {}).get('anticoagulation') else 'no'}, thrombolysis={'yes' if profile.get('contraindications', {}).get('thrombolysis') else 'no'}, high bleeding risk={'yes' if profile.get('contraindications', {}).get('highBleedingRisk') else 'no'}\n"
            f"- Special populations: pregnancy={'yes' if profile.get('specialPopulations', {}).get('pregnancy') else 'no'}, breastfeeding={'yes' if profile.get('specialPopulations', {}).get('breastfeeding') else 'no'}, APS={'yes' if profile.get('specialPopulations', {}).get('aps') else 'no'}, severe CKD={'yes' if profile.get('specialPopulations', {}).get('severeCKD') else 'no'}\n"
            f"- Existing immediate strategy items: {' | '.join(profile.get('immediateStrategy', [])) if isinstance(profile.get('immediateStrategy'), list) else 'n/a'}\n"
            f"- Existing medication strategy items: {' | '.join(profile.get('medicationStrategy', [])) if isinstance(profile.get('medicationStrategy'), list) else 'n/a'}\n\n"
            f"De-identified case narrative:\n{narrative}\n\n"
            "Return only bullet points for immediate care (0-24 hours), avoiding any mention of patient identifiers."
        )

        payload = {
            "model": model,
            "input": [
                {
                    "role": "system",
                    "content": (
                        "You are a pulmonary embolism clinical support assistant. "
                        "Use the structured and narrative inputs to produce immediate, actionable recommendations for frontline clinicians. "
                        "Do not include patient identifiers. Output 3 to 6 concise bullet points only."
                    ),
                },
                {"role": "user", "content": prompt},
            ],
            "temperature": 0.2,
        }

        req = urllib.request.Request(
            "https://api.openai.com/v1/responses",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {OPENAI_API_KEY}",
            },
            method="POST",
        )

        try:
            with urllib.request.urlopen(req, timeout=90) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            self._send_json(exc.code, {"error": "OpenAI request failed.", "details": detail[:1000]})
            return
        except Exception as exc:
            self._send_json(500, {"error": f"AI request failed: {exc}"})
            return

        bullets = parse_bullets(extract_response_text(data))
        if not bullets:
            self._send_json(502, {"error": "OpenAI response did not contain recommendation text."})
            return
        self._send_json(200, {"bullets": bullets})

    def serve_static(self, send_body=True):
        file_path = self._safe_file_path()
        if file_path is None:
            self._send_text(403, "Forbidden")
            return
        if not file_path.exists() or not file_path.is_file():
            self._send_text(404, "Not found")
            return
        mime_type = mimetypes.guess_type(str(file_path))[0] or "application/octet-stream"
        content = file_path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", mime_type)
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        if send_body:
            self.wfile.write(content)


def main():
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    key_status = "configured" if OPENAI_API_KEY else "missing"
    print(f"PE tool server listening on http://{HOST}:{PORT} (OPENAI_API_KEY {key_status})")
    server.serve_forever()


if __name__ == "__main__":
    main()
