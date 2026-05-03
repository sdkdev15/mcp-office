# MCP Office Server

MCP (Model Context Protocol) server untuk menghasilkan dokumen Office — **Excel (.xlsx)**, **Word (.docx)**, dan **PowerPoint (.pptx)** — dengan dukungan format ODF (.ods, .odt, .odp), 5 tema bawaan, dan deployment via Docker.

## Fitur

- **Excel** — Multi-sheet workbook dengan styling, chart, conditional formatting
- **Word** — Dokumen dengan heading, paragraf, tabel, daftar, gambar
- **PowerPoint** — Presentasi dengan slide, chart, tabel, text box
- **ODF Support** — Format LibreOffice (.ods, .odt, .odp)
- **5 Tema Bawaan** — corporate, minimal, creative, academic, dark
- **Custom Template** — Gunakan file .xlsx/.docx/.pptx Anda sebagai template (preserves formatting, formulas, charts, master slides)
- **Multi-format Export** — Generate OOXML + ODF dalam satu panggilan
- **Locale Support** — Indonesia (id_ID) dan Inggris (en_US)
- **Rate Limiting** — Sliding window per user
- **Auto Cleanup** — Hapus file lama otomatis
- **Session Isolation** — Direktori output per sesi
- **Security** — PII redaction, input sanitization, audit trail
- **Docker Ready** — Build dan jalankan via Docker (mode headless/SSE)

---

## Cara Penggunaan

### Mode Desktop (Python langsung via stdio)

Cocok untuk penggunaan di komputer lokal dengan Python terinstall.

**1. Install dependensi:**

```bash
pip install openpyxl python-docx python-pptx odfpy mcp loguru pydantic pydantic-settings aiofiles aiohttp uvicorn starlette
```

**2. Konfigurasi MCP Client:**

```json
{
  "mcpServers": {
    "mcp-office": {
      "type": "stdio",
      "command": "python",
      "args": ["run_server.py"],
      "env": {
        "OUTPUT_DIR": "/path/to/mcp-office/outputs",
        "MCP_TRANSPORT": "stdio"
      }
    }
  }
}
```

> **Catatan:** Pastikan menjalankan dari direktori proyek `mcp-office/` atau set `PYTHONPATH` ke direktori proyek.

---

### Mode Headless / Server (Docker + SSE)

Cocok untuk deployment di server tanpa GUI atau akses jarak jauh.

**1. Build Docker image:**

```bash
docker compose build
```

**2. Jalankan server:**

```bash
docker compose up -d
```

**3. Konfigurasi MCP Client:**

```json
{
  "mcpServers": {
    "mcp-office": {
      "type": "sse",
      "url": "http://YOUR_SERVER_IP:8765/sse"
    }
  }
}
```

> Ganti `YOUR_SERVER_IP` dengan IP atau hostname server Anda. Untuk localhost gunakan `http://localhost:8765/sse`.

---

### Mode Docker stdio (Alternatif)

Jalankan langsung via Docker tanpa SSE:

```json
{
  "mcpServers": {
    "mcp-office": {
      "type": "stdio",
      "command": "docker",
      "args": [
        "run", "-i", "--rm",
        "-v", "mcp-office-outputs:/app/outputs",
        "-e", "MCP_TRANSPORT=stdio",
        "mcp-office-mcp-office:latest"
      ]
    }
  }
}
```

---

## Tool yang Tersedia

| Tool | Deskripsi |
|------|-----------|
| `excel_create` | Buat file Excel dengan sheet, styling, dan tema |
| `excel_export` | Export data ke format xlsx, ods |
| `docx_create` | Buat dokumen Word dengan halaman dan tema |
| `docx_generate_from_prompt` | Generate dokumen Word dari deskripsi bahasa alami |
| `docx_export` | Export dokumen ke format docx, odt |
| `pptx_create` | Buat presentasi PowerPoint dengan slide dan tema |
| `pptx_generate_from_prompt` | Generate presentasi dari deskripsi bahasa alami |
| `pptx_export` | Export presentasi ke format pptx, odp |
| `list_themes_tool` | Lihat semua tema yang tersedia |
| `list_files` | Lihat file yang di-generate untuk sesi tertentu |
| `get_storage_stats` | Statistik penggunaan penyimpanan |

---

## Contoh Penggunaan

### Buat File Excel

```json
{
  "tool": "excel_create",
  "arguments": {
    "filename": "laporan_penjualan.xlsx",
    "sheets": [
      {
        "name": "Data Penjualan",
        "headers": ["Produk", "Q1", "Q2", "Q3", "Q4"],
        "rows": [
          ["Laptop", 150, 200, 180, 250],
          ["HP", 300, 350, 400, 380],
          ["Tablet", 100, 120, 150, 130]
        ]
      }
    ],
    "theme": "corporate"
  }
}
```

### Buat Dokumen Word

```json
{
  "tool": "docx_create",
  "arguments": {
    "filename": "laporan.docx",
    "title": "Laporan Kuartal",
    "theme": "academic",
    "page_size": "A4",
    "orientation": "portrait"
  }
}
```

### Generate Word dari Prompt

```json
{
  "tool": "docx_generate_from_prompt",
  "arguments": {
    "prompt": "Buatkan surat resmi undangan rapat bulanan untuk tim engineering",
    "filename": "undangan_rapat.docx",
    "theme": "corporate"
  }
}
```

### Buat Presentasi PowerPoint

```json
{
  "tool": "pptx_create",
  "arguments": {
    "filename": "presentasi.pptx",
    "title": "Update Proyek",
    "theme": "creative",
    "slide_size": "widescreen"
  }
}
```

### Generate Presentasi dari Prompt

```json
{
  "tool": "pptx_generate_from_prompt",
  "arguments": {
    "prompt": "Buatkan presentasi 5 slide tentang strategi marketing digital untuk UMKM",
    "filename": "marketing_umkm.pptx",
    "theme": "creative"
  }
}
```

### Export Multi-format

```json
{
  "tool": "excel_export",
  "arguments": {
    "sheets": [
      {
        "name": "Data",
        "headers": ["Nama", "Nilai"],
        "rows": [["A", 100], ["B", 200]]
      }
    ],
    "format": "all"
  }
}
```

### Lihat Tema Tersedia

```json
{
  "tool": "list_themes_tool",
  "arguments": {}
}
```

---

## Custom Template (Menggunakan Template Anda Sendiri)

Anda dapat menggunakan file dokumen Anda sendiri sebagai template. Server akan melestarikan semua formatting, styles, formulas, charts, master slides, dan layout dari template Anda, lalu mengisi data baru.

### Excel dengan Template

Gunakan file `.xlsx` Anda sebagai template (misalnya template dengan formula, chart, dan formatting perusahaan):

```json
{
  "tool": "excel_create",
  "arguments": {
    "filename": "laporan_baru.xlsx",
    "template_path": "/path/to/template_perusahaan.xlsx",
    "sheets": [
      {
        "name": "Data",
        "headers": ["Produk", "Q1", "Q2", "Q3", "Q4"],
        "rows": [
          ["Laptop", 150, 200, 180, 250],
          ["HP", 300, 350, 400, 380]
        ]
      }
    ]
  }
}
```

> **Template preserves:** formulas, charts, conditional formatting, pivot tables, data validation, named ranges.

### Word dengan Template

Gunakan file `.docx` Anda sebagai template (misalnya template surat resmi dengan header/logo perusahaan):

```json
{
  "tool": "docx_create",
  "arguments": {
    "filename": "surat_baru.docx",
    "template_path": "/path/to/template_surat.docx",
    "title": "Surat Undangan Rapat",
    "content_paragraphs": [
      "Dengan hormat,",
      "Kami mengundang Bapak/Ibu untuk menghadiri rapat bulanan tim engineering.",
      "Terima kasih."
    ],
    "tables": [
      {
        "headers": ["Agenda", "Waktu", "Pembicara"],
        "rows": [
          ["Review Sprint", "10:00", "PM"],
          ["Planning", "11:00", "Tech Lead"]
        ]
      }
    ]
  }
}
```

> **Template preserves:** headers, footers, page setup, styles, fonts, company logo, watermarks.

### PowerPoint dengan Template

Gunakan file `.pptx` Anda sebagai template (misalnya template presentasi dengan master slide dan branding perusahaan):

```json
{
  "tool": "pptx_create",
  "arguments": {
    "filename": "presentasi_baru.pptx",
    "template_path": "/path/to/template_presentasi.pptx",
    "slides": [
      {
        "layout": "title_and_content",
        "title": "Update Proyek Q1",
        "content": "Presentasi progress kuartal pertama",
        "bullets": ["Target tercapai 95%", "Budget on track", "Tim expanded"]
      },
      {
        "layout": "title_and_content",
        "title": "Rencana Q2",
        "bullets": ["Launch fitur baru", "Scale infrastruktur", "Hire 3 engineer"]
      }
    ]
  }
}
```

> **Template preserves:** master slides, themes, fonts, layouts, company branding, slide transitions.

### Cara Kerja Template

1. **Upload/Tempatkan** file template (.xlsx, .docx, atau .pptx) di server
2. **Berikan path** ke file template via parameter `template_path`
3. **Sertakan data** baru yang ingin diisi (sheets, content_paragraphs, slides, tables)
4. **Server** akan memuat template, melestarikan formatting, dan mengisi data baru
5. **Hasil** disimpan sebagai file baru dengan nama di parameter `filename`

---

## Tema

| Tema | Deskripsi | Warna Utama |
|------|-----------|-------------|
| `corporate` | Profesional dengan nuansa biru | #1E40AF |
| `minimal` | Bersih hitam putih, modern | #000000 |
| `creative` | Cerah ungu dan pink | #7C3AED |
| `academic` | Formal dengan font serif | #1E3A5F |
| `dark` | Dark mode dengan aksen biru | #60A5FA |

---

## Konfigurasi

| Environment Variable | Default | Deskripsi |
|---------------------|---------|-----------|
| `OUTPUT_DIR` | `outputs` | Direktori untuk file hasil generate |
| `FILE_RETENTION_HOURS` | `24` | Jam sebelum file otomatis dihapus |
| `RATE_LIMIT_REQUESTS` | `20` | Maks request per window per user |
| `RATE_LIMIT_WINDOW` | `60` | Window rate limit dalam detik |
| `MCP_TRANSPORT` | `stdio` | Mode transport (stdio atau sse) |
| `LOCALE` | `en_US` | Locale default (en_US, id_ID) |

---

## Struktur Proyek

```
mcp-office/
├── src/
│   ├── server.py              # Entry point MCP server
│   ├── generators/
│   │   ├── excel_generator.py # Excel (.xlsx)
│   │   ├── docx_generator.py  # Word (.docx)
│   │   ├── pptx_generator.py  # PowerPoint (.pptx)
│   │   └── odf_generator.py   # ODF (.ods, .odt, .odp)
│   ├── styles/
│   │   ├── themes.py          # Definisi tema
│   │   └── style_applier.py   # Utility aplikasi style
│   ├── utils/
│   │   ├── file_handler.py    # Operasi file
│   │   ├── cleanup.py         # Auto cleanup file
│   │   ├── rate_limiter.py    # Rate limiting
│   │   ├── validators.py      # Validasi input
│   │   ├── data_transformer.py # Konversi format data
│   │   ├── security.py        # PII redaction, sanitization, audit
│   │   └── logger.py          # Structured logging
│   └── models/
│       └── schemas.py         # Pydantic models
├── run_server.py              # Wrapper script untuk menjalankan server
├── test_all.py                # Test script untuk semua generator
├── tests.py                   # Unit tests
├── .env.example               # Template environment variables
├── Dockerfile                 # Docker image definition
├── docker-compose.yml         # Docker Compose configuration
├── pyproject.toml             # Python project configuration
└── README.md
```

## Dependensi

| Paket | Kegunaan |
|-------|----------|
| **mcp** | Framework MCP server |
| **openpyxl** | Generate file Excel |
| **python-docx** | Generate dokumen Word |
| **python-pptx** | Generate PowerPoint |
| **odfpy** | Support format ODF |
| **pydantic** | Validasi data |
| **pydantic-settings** | Environment configuration |
| **loguru** | Structured logging |
| **aiofiles** | Async file I/O |
| **aiohttp** | HTTP client |
| **uvicorn** | ASGI server (mode SSE) |
| **starlette** | ASGI framework (mode SSE) |

## Testing

```bash
# Test semua generator
python test_all.py

# Run unit tests
pytest tests.py -v
```

## License

MIT