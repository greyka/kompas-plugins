# -*- coding: utf-8 -*-
"""Скачивание по HTTPS для обновлятора и отчётов — устойчиво к «кривым» сертификатам Windows.

20.10 у Димы: urllib → CERTIFICATE_VERIFY_FAILED (unable to get local issuer certificate) для
github.com. В хранилище Windows не было нужного корневого сертификата: браузеры Windows докачивают
корни сами, Python — нет. Поэтому:
1) проверяем сертификат по хранилищу Windows + своему набору cacert.pem (certifi, лежит рядом);
2) при ошибке SSL — качаем встроенным в Windows curl.exe (Schannel, проверка средствами Windows).
"""
import os
import ssl
import subprocess
import urllib.error
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
CAFILE = os.path.join(HERE, "cacert.pem")
UA = "kompas-plugins-updater"
_ctx = None


def context():
    global _ctx
    if _ctx is None:
        _ctx = ssl.create_default_context()          # хранилище Windows (ROOT/CA)
        if os.path.isfile(CAFILE):
            try:
                _ctx.load_verify_locations(cafile=CAFILE)   # + certifi
            except Exception:
                pass
    return _ctx


def _is_ssl(e):
    return isinstance(e, ssl.SSLError) or isinstance(getattr(e, "reason", None), ssl.SSLError) \
        or "CERTIFICATE_VERIFY_FAILED" in str(e)


def _curl():
    exe = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "curl.exe")
    return exe if os.path.isfile(exe) else None


def _curl_get(url, timeout, head_only=False):
    """→ (байты, итоговый URL после редиректов) через curl.exe Windows."""
    exe = _curl()
    if not exe:
        raise RuntimeError("нет curl.exe")
    args = [exe, "-fsSL", "--max-time", str(int(timeout)), "-A", UA, "-w", "\n%{url_effective}"]
    if head_only:
        args += ["-o", os.devnull]
    r = subprocess.run(args + [url], capture_output=True, creationflags=0x08000000)
    if r.returncode:
        raise RuntimeError(f"curl.exe: {r.stderr.decode('utf-8', 'replace').strip() or r.returncode}")
    body, _, final = r.stdout.rpartition(b"\n")
    return body, final.decode("utf-8", "replace").strip()


def fetch(url, timeout=30, data=None, headers=None):
    """→ (байты, итоговый URL). data — POST (только через urllib)."""
    req = urllib.request.Request(url, data, {"User-Agent": UA, **(headers or {})})
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=context()) as r:
            return r.read(), r.geturl()
    except Exception as e:
        if data is not None or not _is_ssl(e):
            raise
        try:
            return _curl_get(url, timeout)
        except Exception as e2:
            raise RuntimeError(f"{e}; через curl.exe тоже не вышло: {e2}") from None


def get(url, timeout=30):
    return fetch(url, timeout)[0]


def final_url(url, timeout=30):
    """Куда ведёт ссылка (редирект github.com/…/releases/latest → …/tag/vX.Y.Z)."""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": UA})
        with urllib.request.urlopen(req, timeout=timeout, context=context()) as r:
            return r.geturl()
    except Exception as e:
        if not _is_ssl(e):
            raise
        return _curl_get(url, timeout, head_only=True)[1]
