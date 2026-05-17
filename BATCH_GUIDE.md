# MCP Office Batch & Templating Guide

This guide explains how to use the Batch Generation & Templating tools (Track D) to generate multiple personalized documents at once (mail merge) and merge documents together.

## Overview

Track D introduces two powerful tools:
1. `batch_create_documents` - Generates a ZIP file containing multiple personalized documents based on a single template and a list of datasets.
2. `merge_documents` - Combines multiple `.docx` or `.pdf` files into a single master file.

---

## 1. batch_create_documents

Use this tool to generate dozens or hundreds of personalized documents using a Handlebars-like syntax directly in the JSON payload.

### Templating Syntax

The template engine processes the JSON payload as a string before parsing it into a dictionary for generation.

- **Variables**: `{{ name }}` (Substitutes the variable from the dataset)
- **Filters**: `{{ name | uppercase }}` (Transforms the output)
- **JSON Escaping**: `{{ name | json }}` (Escapes quotes and newlines for safe injection into the JSON payload string)
- **Conditionals**: `{{#if show_table}} ... {{/if}}`
- **Loops**: `{{#each items}} ... {{/each}}`

### Example Workflow: Mail Merge

Imagine you want to generate a personalized Word document for multiple clients.

```json
{
  "format": "docx",
  "theme": "corporate",
  "template": "{\"filename\": \"report_{{id}}.docx\", \"sections\": [{\"type\": \"title\", \"text\": \"Welcome {{client_name}}\"}, {{#if include_table}}{\"type\": \"table\", \"headers\": [\"Item\", \"Cost\"], \"rows\": [{{#each purchases}} [\"{{this.name}}\", \"{{this.cost}}\"] {{#if not_last}},{{/if}} {{/each}}]}{{/if}}]}",
  "datasets": [
    {
      "id": "1",
      "client_name": "Acme Corp",
      "include_table": true,
      "purchases": [
        {"name": "Server", "cost": "$5000", "not_last": true},
        {"name": "License", "cost": "$1000", "not_last": false}
      ]
    },
    {
      "id": "2",
      "client_name": "Globex",
      "include_table": false
    }
  ]
}
```

> **Tip:** When constructing loops inside a JSON array, be careful with commas between items. The template engine does not automatically handle JSON trailing commas, so a boolean flag like `not_last` is helpful.

---

## 2. merge_documents

Combines multiple `.docx` or `.pdf` files into a single file. 

**When to use:** When you have generated multiple files and want to provide the user with a single master document for printing or distribution.

### Example Payload
```json
{
  "file_paths": [
    "/absolute/path/to/report_1.docx",
    "/absolute/path/to/report_2.docx"
  ],
  "output_filename": "master_report.docx"
}
```

> **Note on PDFs:** Merging PDFs uses `pypdf`. Ensure the files are valid PDF documents before attempting to merge.
