"""ODF (OpenDocument Format) generator using odfpy for LibreOffice native format support."""

from __future__ import annotations

import io
from typing import Any, Optional

from odf.opendocument import OpenDocumentSpreadsheet, OpenDocumentText, OpenDocumentPresentation
from odf.table import Table, TableRow, TableCell
from odf.text import P, H
from odf.style import Style, TextProperties, ParagraphProperties, TableProperties

from src.styles.themes import get_theme
from src.utils.logger import get_logger

log = get_logger("odf_generator")


class ODFGenerator:
    """Generates ODF format files (.ods, .odt, .odp) for LibreOffice."""

    def __init__(self, theme_name: str = "corporate"):
        self.theme_name = theme_name
        self.theme = get_theme(theme_name)

    def create_spreadsheet(self, sheets: list[dict], metadata: Optional[dict] = None) -> bytes:
        """Create an ODS spreadsheet from sheet data.

        Args:
            sheets: List of sheet data dictionaries with name, headers, rows.
            metadata: Optional document metadata.

        Returns:
            Spreadsheet content as bytes.
        """
        doc = OpenDocumentSpreadsheet()

        if metadata:
            self._apply_metadata(doc, metadata)

        for sheet_data in sheets:
            self._add_sheet(doc, sheet_data)

        return self._save_document(doc)

    def create_text_document(self, title: str = "Document", metadata: Optional[dict] = None) -> OpenDocumentText:
        """Create a new ODT text document.

        Args:
            title: Document title.
            metadata: Optional document metadata.

        Returns:
            OpenDocumentText instance for further modification.
        """
        doc = OpenDocumentText()

        if metadata:
            self._apply_metadata(doc, metadata)

        # Add title heading
        heading = H(text=title, stylename="Heading_1")
        doc.text.append(heading)

        return doc

    def create_presentation(self, title: str = "Presentation", metadata: Optional[dict] = None) -> OpenDocumentPresentation:
        """Create a new ODP presentation.

        Args:
            title: Presentation title.
            metadata: Optional document metadata.

        Returns:
            OpenDocumentPresentation instance.
        """
        doc = OpenDocumentPresentation()

        if metadata:
            self._apply_metadata(doc, metadata)

        return doc

    def add_paragraph_to_odt(self, doc: OpenDocumentText, text: str, bold: bool = False) -> None:
        """Add a paragraph to an ODT document.

        Args:
            doc: OpenDocumentText instance.
            text: Paragraph text.
            bold: Bold text.
        """
        para = P(text=text)
        if bold:
            para.setAttribute("foo", "bar")  # odfpy styling is limited
        doc.text.append(para)

    def add_table_to_odt(self, doc: OpenDocumentText, headers: list[str], rows: list[list[Any]]) -> None:
        """Add a table to an ODT document.

        Args:
            doc: OpenDocumentText instance.
            headers: Table headers.
            rows: Table data rows.
        """
        table = Table(stylename="table1", name="table1")
        table.columns = len(headers)

        # Header row
        header_row = TableRow()
        for header in headers:
            cell = TableCell()
            cell.addElement(P(text=str(header)))
            header_row.addElement(cell)
        table.addElement(header_row)

        # Data rows
        for row_data in rows:
            row = TableRow()
            for value in row_data:
                cell = TableCell()
                cell.addElement(P(text=str(value)))
                row.addElement(cell)
            table.addElement(row)

        doc.text.addElement(table)

    def _add_sheet(self, doc: OpenDocumentSpreadsheet, sheet_data: dict) -> None:
        """Add a sheet to the ODS spreadsheet.

        Args:
            doc: OpenDocumentSpreadsheet instance.
            sheet_data: Sheet data with name, headers, rows.
        """
        name = sheet_data.get("name", "Sheet1")
        headers = sheet_data.get("headers", [])
        rows = sheet_data.get("rows", [])

        table = Table(name=name)
        doc.spreadsheet.addElement(table)

        # Add header row
        if headers:
            header_row = TableRow()
            for header in headers:
                cell = TableCell(stylename="hd")
                cell.addElement(P(text=str(header)))
                header_row.addElement(cell)
            table.addElement(header_row)

        # Add data rows
        for row_data in rows:
            row = TableRow()
            for value in row_data:
                cell = TableCell()
                cell.addElement(P(text=str(value) if value is not None else ""))
                row.addElement(cell)
            table.addElement(row)

    def _apply_metadata(self, doc, metadata: dict) -> None:
        """Apply document metadata.

        Args:
            doc: ODF document instance.
            metadata: Metadata dictionary.
        """
        office_doc = doc.officeDocument
        if metadata.get("author"):
            office_doc.setAttribute("creator", metadata["author"])
        if metadata.get("title"):
            office_doc.setTitle(metadata["title"])
        if metadata.get("subject"):
            office_doc.setSubject(metadata["subject"])

    def _save_document(self, doc) -> bytes:
        """Save ODF document to bytes.

        Args:
            doc: ODF document instance.

        Returns:
            Document content as bytes.
        """
        buffer = io.BytesIO()
        doc.save(buffer, pretty_print=True)
        buffer.seek(0)
        return buffer.getvalue()

    def save_odt(self, doc: OpenDocumentText) -> bytes:
        """Save ODT document to bytes.

        Args:
            doc: OpenDocumentText instance.

        Returns:
            Document content as bytes.
        """
        return self._save_document(doc)

    def save_odp(self, doc: OpenDocumentPresentation) -> bytes:
        """Save ODP presentation to bytes.

        Args:
            doc: OpenDocumentPresentation instance.

        Returns:
            Document content as bytes.
        """
        return self._save_document(doc)


# Global generator instance
odf_generator = ODFGenerator()