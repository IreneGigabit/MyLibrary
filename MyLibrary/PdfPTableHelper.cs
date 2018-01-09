using iTextSharp.text;
using iTextSharp.text.pdf;

namespace MyLibrary {
	public class PdfPTableHelper {
		PdfPTable _table = null;
		Font _font = null;
		PdfPCell _cell = null;

		public PdfPTableHelper(PdfPTable table) {
			this._table = table;
		}

		public PdfPTableHelper(PdfPTable table, Font font) {
			this._table = table;
			this._font = font;
		}

		public PdfPTableHelper Font(Font font) {
			this._font = font;
			return this;
		}

		public PdfPTableHelper Cell(string str) {
			return this.Cell(str, PdfPCell.ALIGN_LEFT, PdfPCell.ALIGN_TOP);
		}

		public PdfPTableHelper Cell(string str, int align) {
			return this.Cell(str, align, PdfPCell.ALIGN_TOP);
		}

		public PdfPTableHelper Cell(string str, int align, int valign) {
			if (this._font != null)
				this._cell = new PdfPCell(new Phrase(str, this._font));
			else
				this._cell = new PdfPCell(new Phrase(str));

			this._cell.HorizontalAlignment = align;
			this._cell.VerticalAlignment = valign;

			return this;
		}

		public PdfPTableHelper Add() {
			this._table.AddCell(this._cell);

			return this;
		}

		public PdfPTableHelper AddCell(string str) {
			this.Cell(str);
			this._table.AddCell(this._cell);

			return this;
		}

		public PdfPTableHelper AddTable(PdfPTable table) {
			this.AddTable(table, 1);

			return this;
		}

		public PdfPTableHelper AddTable(PdfPTable table, int colSpan) {
			this._cell = new PdfPCell(table);
			this._cell.Colspan = colSpan;
			this._table.AddCell(this._cell);

			return this;
		}

		public PdfPTableHelper Align(int align) {
			this._cell.HorizontalAlignment = align;
			return this;
		}

		public PdfPTableHelper Valign(int valign) {
			this._cell.VerticalAlignment = valign;
			return this;
		}

		public PdfPTableHelper Colspan(int colSpan) {
			this._cell.Colspan = colSpan;
			return this;
		}

		public PdfPTableHelper Border(int border) {
			this._cell.Border = border;
			return this;
		}

		public PdfPTableHelper Border(int topBorder, int leftBorder, int bottonBorder, int rightBorder) {
			this._cell.BorderWidthTop = topBorder;
			this._cell.BorderWidthLeft = leftBorder;
			this._cell.BorderWidthBottom = bottonBorder;
			this._cell.BorderWidthRight = rightBorder;
			return this;
		}

		public PdfPTableHelper BTop(int width) {
			this._cell.BorderWidthTop = width;
			return this;
		}

		public PdfPTableHelper BLeft(int width) {
			this._cell.BorderWidthLeft = width;
			return this;
		}

		public PdfPTableHelper BBottom(int width) {
			this._cell.BorderWidthBottom = width;
			return this;
		}

		public PdfPTableHelper BRight(int width) {
			this._cell.BorderWidthRight = width;
			return this;
		}

	}
}
