"""Embedded icon library with ~60 SVG paths and keyword-based auto-detection.

All icons are normalized to a 24x24 viewBox. Colors are applied at render time
so the paths here are pure geometry -- no fill or stroke attributes.
"""

from __future__ import annotations

from typing import Optional


# ── Icon Registry ──

# Each icon: { "path": "<svg d>", "categories": [keyword, ...] }
ICONS: dict[str, dict[str, object]] = {}


def _register(name: str, path: str, categories: list[str]) -> None:
    """Register an icon with its SVG path and keyword categories."""
    ICONS[name] = {"path": path, "categories": categories}


# ── Security Icons ──
_REGISTERED_SECURITY = [
    ("lock", "M12 2C9.24 2 7 4.24 7 7v3H6a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-8a2 2 0 0 0-2-2h-1V7c0-2.76-2.24-5-5-5zm0 2c1.65 0 3 1.35 3 3v3H9V7c0-1.65 1.35-3 3-3zm-2.5 10a2.5 2.5 0 1 1 0 5 2.5 2.5 0 0 1 0-5zm0 1.5a1 1 0 1 0 0 2 1 1 0 0 0 0-2z",
     ["lock", "security", "downtime", "session", "flock", "authentication", "secure", "blocked", "berkunci"]),
    ("unlock", "M12 2C9.24 2 7 4.24 7 7v3H6a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-8a2 2 0 0 0-2-2h-1V7c0-2.76-2.24-5-5-5zm0 2c1.65 0 3 1.35 3 3v3H9V7c0-1.65 1.35-3 3-3z",
     ["unlock", "unlocked", "free", "available", "tersedia"]),
    ("shield", "M12 2L4 5v6.09c0 5.06 3.41 9.76 8 10.91 4.59-1.15 8-5.85 8-10.91V5l-8-3zm0 2.18l6 2.25v5.66c0 4.09-2.68 7.81-6 8.93-3.32-1.12-6-4.84-6-8.93V6.43l6-2.25z",
     ["shield", "protection", "protect", "defense", "perlindungan", "aman"]),
    ("key", "M12.65 10C11.83 7.67 9.61 6 7 6c-3.31 0-6 2.69-6 6s2.69 6 6 6c2.61 0 4.83-1.67 5.65-4H17v4h4v-4h2v-4H12.65zM7 14c-1.65 0-3-1.35-3-3s1.35-3 3-3 3 1.35 3 3-1.35 3-3 3z",
     ["key", "access", "token", "password", "kunci", "akses"]),
    ("fingerprint", "M2 2h20v20H2V2zm2 2v16h16V4H4zm10 2c-1.1 0-2 .9-2 2v2h-2v2h2v2c0 1.1.9 2 2 2s2-.9 2-2v-2h2v-2h-2V8c0-1.1-.9-2-2-2z",
     ["fingerprint", "biometric", "identity", "identitas", "sidik jari"]),
]

# ── Server / Infrastructure Icons ──
_REGISTERED_SERVER = [
    ("server", "M3 5h18v4H3V5zm0 6h18v4H3v-4zm0 6h18v4H3v-4zm2 1v2h2v-2H5zm8 0v2h2v-2h-2z",
     ["server", "infrastruktur", "hosting", "backend", "pelayan"]),
    ("database", "M12 4c-4.42 0-8 2.24-8 5s3.58 5 8 5 8-2.24 8-5-3.58-5-8-5zm0 2c3.31 0 6 1.34 6 3s-2.69 3-6 3-6-1.34-6-3 2.69-3 6-3zm0 10c-4.42 0-8 2.24-8 5s3.58 5 8 5 8-2.24 8-5-3.58-5-8-5zm0 2c3.31 0 6 1.34 6 3s-2.69 3-6 3-6-1.34-6-3 2.69-3 6-3z",
     ["database", "db", "storage", "penyimpanan", "data", "mysql", "postgres", "redis"]),
    ("cloud", "M12 4c-3.31 0-6 2.69-6 6 0 1.46.52 2.79 1.37 3.83C6.17 14.86 5.5 15.81 5.5 16.9 5.5 18.61 6.89 20 8.6 20h7.8c2.21 0 4-1.79 4-4s-1.79-4-4-4c-.56 0-1.09.12-1.58.33C13.77 10.94 12.52 10 11.07 10c-.72 0-1.4.13-2.03.36C9.31 9.14 10.52 8 12 8",
     ["cloud", "aws", "azure", "gcp", "hosting", "awan", "online", "remote"]),
    ("docker", "M12 2L4 6v8c0 4.42 3.58 8 8 8s8-3.58 8-8V6l-8-4zm-2 12v-4h4v4h-4z",
     ["docker", "container", "kubernetes", "k8s", "deploy"]),
    ("network", "M4 6h16v2H4V6zm4 4h8v2H8v-2zm-2 4h12v2H6v-2z",
     ["network", "jaringan", "connect", "koneksi", "link", "router"]),
    ("cpu", "M8 3h8v2h2v2h2v8h-2v2h-2v2H8v-2H6v-2H4V7h2V5h2V3zm2 4v6h4V7h-4z",
     ["cpu", "processor", "processing", "komputer", "processor"]),
    ("memory", "M3 3h18v18H3V3zm2 2v14h14V5H5zm3 3h2v8H8V8zm4 0h2v8h-2V8z",
     ["memory", "ram", "storage", "penyimpanan", "cache"]),
    ("disk", "M12 4a8 8 0 1 0 0 16 8 8 0 0 0 0-16zm0 14a6 6 0 1 1 0-12 6 6 0 0 1 0 12zm0-10a4 4 0 1 0 0 8 4 4 0 0 0 0-8z",
     ["disk", "hdd", "ssd", "storage", "penyimpanan"]),
]

# ── Business Icons ──
_REGISTERED_BUSINESS = [
    ("dollar", "M12 2v20M11 4h2M11 20h2M7 8c0-1.5 1.5-2.5 5-2.5s5 1 5 2.5-2 2.5-5 3-5 1.5-5 3 1.5 2.5 5 2.5 5-1 5-2.5",
     ["dollar", "money", "uang", "revenue", "pendapatan", "income", "finance"]),
    ("trend-up", "M3 17l4-4 3 3 6-6M3 21h18M3 3v18",
     ["trend", "up", "growth", "naik", "pertumbuhan", "increase", "上升"]),
    ("trend-down", "M3 7l4 4 3-3 6 6M3 21h18M3 3v18",
     ["trend", "down", "decline", "turun", "penurunan", "decrease"]),
    ("target", "M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm0 16a6 6 0 1 1 0-12 6 6 0 0 1 0 12zm0-4a2 2 0 1 1 0-4 2 2 0 0 1 0 4z",
     ["target", "goal", "tujuan", "sasaran", "objective"]),
    ("award", "M12 2l3 7h7l-5.5 4 2 7L12 16l-6.5 4 2-7L2 9h7l3-7z",
     ["award", "achievement", "prestasi", "penghargaan", "star"]),
    ("users", "M7 7a3 3 0 1 0 0 6 3 3 0 0 0 0-6zm-5 9c0-2 1.5-3.5 4-4 1.5-1 3.5-1.5 6-1.5s4.5.5 6 1.5c2.5.5 4 2 4 4v2H2v-2z",
     ["users", "team", "tim", "orang", "employee", "karyawan"]),
    ("calendar", "M4 4h16v16H4V4zm2 2v12h12V6H6zm0-2h2v2H6V4zm8 0h2v2h-2V4z",
     ["calendar", "tanggal", "date", "jadwal", "schedule"]),
]

# ── Action / Status Icons ──
_REGISTERED_ACTION = [
    ("alert", "M12 2L2 20h20L12 2zm0 4v6m0 4v2",
     ["alert", "warning", "peringatan", "notice", "attention", "hati-hati"]),
    ("warning", "M12 2L1 21h22L12 2zm0 6v5m0 5v1",
     ["warning", "hati-hati", "caution", "bahaya", "berbahaya"]),
    ("check", "M20 6L9 17l-5-5",
     ["check", "done", "berhasil", "complete", "selesai", "success"]),
    ("xmark", "M18 6L6 18M6 6l12 12",
     ["xmark", "close", "tutup", "cancel", " Batal", "error"]),
    ("info", "M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm0 16v-6m0-4h.01",
     ["info", "information", "informasi", "detail", "details"]),
    ("play", "M6 4l14 8-14 8V4z",
     ["play", "putar", "start", "mulai", "run", "eksekusi"]),
    ("pause", "M6 4h4v16H6V4zm8 0h4v16h-4V4z",
     ["pause", "jeda", "stop", "berhenti", "hold"]),
    ("refresh", "M3 12a9 9 0 0 1 9-9 9 9 0 0 1 6 2.75M21 3v4h-4M21 21a9 9 0 0 1-9-9 9 9 0 0 1-6-2.75M3 21v-4h4",
     ["refresh", "refresh", "reload", "muat ulang", "sync", "sinkronisasi"]),
    ("settings", "M12 8a4 4 0 1 0 0 8 4 4 0 0 0 0-8zm0 2a2 2 0 1 1 0 4 2 2 0 0 1 0-4z",
     ["settings", "konfigurasi", "config", "opsi", "options"]),
]

# ── Data / Analytics Icons ──
_REGISTERED_DATA = [
    ("chart-bar", "M3 20h18M6 20V10M10 20V4M14 20V8M18 20V12",
     ["chart", "bar", "grafik", "statistics", "statistik", "data", "analisis"]),
    ("chart-pie", "M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm0 2a8 8 0 0 1 0 16 8 8 0 0 1 0-16zm0 4a4 4 0 0 0 0 8",
     ["chart", "pie", "bagian", "porsi", "persentase", "percentage"]),
    ("chart-line", "M3 20L7 12l4 4 4-8 6 4M3 21h18",
     ["chart", "line", "grafik garis", "trend", "kenaikan", "penurunan"]),
    ("table", "M3 3h18v18H3V3zm0 6h18M9 3v18",
     ["table", "tabel", "grid", "spreadsheet", "xls"]),
    ("file", "M4 2h8l6 6v12H4V2zm8 0v6h6",
     ["file", "dokumen", "document", "laporan", "report"]),
    ("folder", "M2 6h8l2 2h10v12H2V6z",
     ["folder", "direktori", "directory", " Arsip", "arsip"]),
]

# ── Communication Icons ──
_REGISTERED_COMM = [
    ("mail", "M4 4h16v14H4V4zm0 2l8 5 8-5",
     ["mail", "email", "surat", "pesanan", "message"]),
    ("chat", "M4 4h16v12H8l-4 4V4zm0 0h16M8 8h8M8 12h5",
     ["chat", "obrolan", "conversation", "diskusi", "discuss"]),
    ("phone", "M3 3l4 2-1 3 3 3 3-1 2 4-3 3-5-2-3-3-2-5 3-3",
     ["phone", "telepon", "call", "hubungi", "contact"]),
    ("globe", "M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm-2 16V6h4v12M4 12h4M16 12h4",
     ["globe", "internet", "web", "global", "website", "online"]),
]

# ── DevOps Icons ──
_REGISTERED_DEVOPS = [
    ("git-branch", "M6 4a2 2 0 1 0 0 4 2 2 0 0 0 0-4zm12 2a2 2 0 1 0 0 4 2 2 0 0 0 0-4zM6 8v4a2 2 0 0 0 2 2h4a2 2 0 0 0 2-2V8H6z",
     ["git", "branch", "cabang", "version", "versi", "repository"]),
    ("terminal", "M3 3h18v16H3V3zm2 2l4 4-4 4M11 13h6",
     ["terminal", "konsol", "console", "command", "cmd", "bash"]),
    ("code", "M4 8l-2 2 2 2M20 8l-2 2 2 2M10 16l-4 4-4-4",
     ["code", "kode", "programming", "development", "dev"]),
    ("pipeline", "M3 6h4v4H3V6zm8 0h4v4h-4V6zm8 0h2v4h-2V6zM3 14h4v4H3v-4zm8 0h4v4h-4v-4zm8 0h2v4h-2v-4z",
     ["pipeline", "CI/CD", "build", "deploy", "otomatisasi", "automation"]),
    ("bug", "M12 2a5 5 0 0 0-5 5v1H5v2h1v4H4v2h1v2H4v2h16v-2h-1v-2h2v-2h-1v-4h1V8h-2V7a5 5 0 0 0-5-5zm-2 8h4v2h-4V10zm0 4h4v2h-4v-2z",
     ["bug", "serangan", "error", "keluhan", "masalah", "issue"]),
]

# ── Additional Icons ──
_REGISTERED_EXTRA = [
    # ── Flow / Diagram ──
    ("arrow-right", "M5 12h14m-4-4l4 4-4 4", ["arrow", "right", "kanan", "lanjut", "next", "forward"]),
    ("arrow-left", "M19 12H5m4-4l-4 4 4 4", ["arrow", "left", "kiri", "back", "mundur", "previous"]),
    ("arrow-up", "M12 5v14m-4-4l4 4 4-4", ["arrow", "up", "atas", "naik", "increase", "kembali"]),
    ("arrow-down", "M12 19V5m-4 4l4-4 4 4", ["arrow", "down", "bawah", "turun", "decrease", "download"]),
    ("flow", "M4 6h4M16 6h4M6 6v12h12V6M8 18h8", ["flow", "aliran", "process", "proses", "workflow"]),
    ("connection", "M3 12h4l3-6 3 12 3-6h5", ["connection", "konektif", "connect", "koneksi", "link", "hubung"]),

    # ── Time ──
    ("clock", "M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm0 4v6l4 2", ["clock", "jam", "time", "waktu", "durasi", "duration"]),
    ("hourglass", "M6 4h12v4l-4 4 4 4v4H6v-4l4-4-4-4V4z", ["hourglass", "waktu", "time", "countdown", "berakhir", "expiring"]),

    # ── Process ──
    ("process", "M4 4h6v6H4V4zm10 0h6v6h-6V4zM4 14h6v6H4v-6zm10 0h6v6h-6v-6z", ["process", "proses", "step", "langkah", "flow"]),
    ("sync", "M4 12a8 8 0 0 1 14-5M20 12a8 8 0 0 1-14 5", ["sync", "sinkronisasi", "refresh", "reload", "update"]),
    ("repeat", "M3 8h8l3 4-3 4H3M21 16h-8l-3-4 3-4h8", ["repeat", "ulang", "loop", "iterasi", "cycling"]),

    # ── Quality ──
    ("quality", "M12 2l2 7h7l-5.5 4 2 7L12 16l-6.5 4 2-7L2 9h7l3-7z", ["quality", "kualitas", "star", "prestasi", "excellent"]),
    ("speed", "M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm0 4l2 6-6 2", ["speed", "kecepatan", "fast", "lambat", "slow", "performa"]),

    # ── Communication Extended ──
    ("bell", "M12 2a7 7 0 0 0-7 7c0 5.25-2 7-2 7h18s-2-1.75-2-7a7 7 0 0 0-7-7zm0 14a2 2 0 1 0 0 4 2 2 0 0 0 0-4z", ["bell", "lonceng", "notifikasi", "notification", "alert", "peringatan"]),
    ("megaphone", "M4 8l12 4-3 8 4 2-6-10H4V8z", ["megaphone", "pengumuman", "announcement", "broadcast", "info", "berita"]),

    # ── Infrastructure Extended ──
    ("load-balancer", "M12 2v6M4 8h16M4 14h5M15 14h5M4 20h5M15 20h5M9 14v6m6-6v6", ["load-balancer", "lb", "load balance", "distribusi", "trafik", "traffic"]),
    ("firewall", "M3 3h18v18H3V3zm2 2v14h14V5H5zm2 2h10v2H7V7zm0 4h7v2H7v-2z", ["firewall", "dinding api", "keamanan", "security", "proteksi"]),
    ("backup", "M4 4h16v16H4V4zm2 2v12h12V6H6zm4 3h4v5h-4V9zm-2 6h8v2H8v-2z", ["backup", "cadangan", "restore", "recovery", "arsip"]),
    ("monitor", "M3 3h18v12H3V3zm2 2v8h14V5H5zm4 14h6v2H9v-2z", ["monitor", "layar", "screen", "display", "monitoring"]),
    ("web", "M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm0 3v4m-3 0h6M4 12h4m8 0h4", ["web", "website", "browser", "internet", "http", "url"]),

    # ── Error / Status Extended ──
    ("error", "M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm4-4-8 8m0 8 8-8", ["error", "error", "gagal", "failed", "problem", "masalah", "x"]),
    ("success", "M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm4-4l-5 5-4-4", ["success", "sukses", "berhasil", "selesai", "complete", "check"]),
    ("question", "M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm-1 14v-1m0-4a3 3 0 0 1 3-3c1.5 0 2.5 1 2.5 2.5 0 1.5-1.5 2-2.5 2.5S11 14 11 14", ["question", "pertanyaan", "tanya", "FAQ", "help", "bantuan"]),

    # ── Presentation Elements ──
    ("presentation", "M3 3h18v14H3V3zm2 2v10l4-3 4 3 5-4V5H5zm4 14v2m8-2v2", ["presentation", "presentasi", "slide", "ppt", "deck"]),
    ("image", "M3 3h18v18H3V3zm2 2v14h14V5H5zm4 4l3 4 4-5 2 3M9 11a2 2 0 1 0 0-4 2 2 0 0 0 0 4", ["image", "gambar", "photo", "foto", "picture", "foto"]),

    # ── Cloud / Network Extended ──
    ("api", "M4 8h4v8H4V8zm12 0h4v8h-4V8zM8 6l4 4m0 4l-4 4", ["api", "endpoint", "rest", "graphql", "endpoint", "service"]),
    ("upload", "M12 19V5m-4 4l4-4 4 4M3 15v4h18v-4", ["upload", "unggah", "upload", "send", "kirim", "transfer"]),
    ("download", "M12 5v14m-4-4l4 4 4-4M3 15v4h18v-4", ["download", "unduh", "download", "ambil", "receive", "terima"]),
    ("link", "M8 8l3-3m5 5l3-3M5 17l14-14m-4 6l-3 3m-3-3l-3 3m3-3l-3-3", ["link", "tautan", "hyperlink", "url", "href"]),
    ("search", "M11 2a9 9 0 1 0 0 18 9 9 0 0 0 0-18zm7 15l4 4M8 8h6", ["search", "cari", "find", "telusuri", "lookup"]),

    # ── Incident / Monitoring ──
    ("incident", "M12 2L1 21h22L12 2zm0 6v5m0 5v1", ["incident", "insiden", "kejadian", "incident", "emergency", "darurat"]),
    ("analytics", "M3 3v18h18M7 16l4-4 3 3 4-5", ["analytics", "analisis", "analytics", "data", "insight", "wawasan"]),
    ("audit", "M3 3h18v18H3V3zm2 2v14h14V5H5zm4 4h6v2H9V9zm0 4h4v2H9v-2z", ["audit", "audit", "revisi", "review", "pemeriksaan", "inspection"]),
]

# ── All icon data in one list for registration ──
_ALL_ICONS = (
    _REGISTERED_SECURITY
    + _REGISTERED_SERVER
    + _REGISTERED_BUSINESS
    + _REGISTERED_ACTION
    + _REGISTERED_DATA
    + _REGISTERED_COMM
    + _REGISTERED_DEVOPS
    + _REGISTERED_EXTRA
)


def _init_icons() -> None:
    """Register all icons into the ICONS dict."""
    for name, path, categories in _ALL_ICONS:
        _register(name, path, categories)


# Auto-register on module load
_init_icons()


# ── Public API ──

def get_icon_names() -> list[str]:
    """Return all available icon names."""
    return list(ICONS.keys())


def get_icon(name: str) -> Optional[dict]:
    """Get an icon by name. Returns None if not found."""
    return ICONS.get(name)


def find_icon_by_keyword(text: str) -> Optional[str]:
    """Search icon library by keyword matching against text.

    Uses a two-pass approach:
    1. Whole-word matches get priority (highest score).
    2. Substring matches are fallbacks (lower score).

    This avoids false positives like "line" matching inside "pipeline"
    when a proper "pipeline" icon exists.

    Args:
        text: Title/content text to extract keywords from.

    Returns:
        Icon name if found, None otherwise.
    """
    if not text:
        return None

    text_lower = text.lower()
    words = set(text_lower.split())

    # Collect best match — whole-word > substring
    best_icon = None
    best_score = -1

    for name, icon in ICONS.items():
        for cat in icon["categories"]:
            cat_lower = cat.lower()

            if cat_lower in words:
                # Whole-word match (score 2)
                score = 2
                if score > best_score:
                    best_score = score
                    best_icon = name
                    break  # Can't do better than whole-word match
            elif " " in cat_lower or "/" in cat_lower:
                # Multi-word/special keywords use substring (score 1)
                if cat_lower in text_lower:
                    score = 1
                    if score > best_score:
                        best_score = score
                        best_icon = name
                    break

    return best_icon