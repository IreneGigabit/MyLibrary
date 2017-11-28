using iTextSharp.text;
using iTextSharp.text.pdf;

namespace MyLibrary {
	public static class PdfPCellExt {
		public static PdfPCell Init(this PdfPCell cell, string str, Font font) {
			cell = new PdfPCell(new Phrase(str, font));
			return cell;
		}

		public static PdfPCell Init(this PdfPCell cell, PdfPTable addTable) {
			cell = new PdfPCell(addTable);
			return cell;
		}

		public static PdfPCell Init(this PdfPCell cell, string str, Font font, int align) {
			cell = new PdfPCell(new Phrase(str, font)).Align(align);
			return cell;
		}

		public static PdfPCell Init(this PdfPCell cell, string str, Font font, int align, int valign) {
			cell = new PdfPCell(new Phrase(str, font)).Align(align).Valign(valign);
			return cell;
		}

		public static PdfPCell Align(this PdfPCell cell, int align) {
			cell.HorizontalAlignment = align;
			return cell;
		}

		public static PdfPCell Valign(this PdfPCell cell, int valign) {
			cell.VerticalAlignment = valign;
			return cell;
		}

		public static PdfPCell Colspan(this PdfPCell cell, int colSpan) {
			cell.Colspan = colSpan;
			return cell;
		}

		public static PdfPCell Border(this PdfPCell cell, int topBorder, int leftBorder, int bottonBorder, int rightBorder) {
			cell.BorderWidthTop = topBorder;
			cell.BorderWidthLeft = leftBorder;
			cell.BorderWidthBottom = bottonBorder;
			cell.BorderWidthRight = rightBorder;
			return cell;
		}

		public static PdfPCell Border(this PdfPCell cell, int border) {
			cell.BorderWidth = border;
			return cell;
		}

		public static PdfPCell BorderTop(this PdfPCell cell, int width) {
			cell.BorderWidthTop = width;
			return cell;
		}

		public static PdfPCell BorderLeft(this PdfPCell cell, int width) {
			cell.BorderWidthLeft = width;
			return cell;
		}

		public static PdfPCell BorderBottom(this PdfPCell cell, int width) {
			cell.BorderWidthBottom = width;
			return cell;
		}

		public static PdfPCell BorderRight(this PdfPCell cell, int width) {
			cell.BorderWidthRight = width;
			return cell;
		}

		public static PdfPCell MinHeight(this PdfPCell cell, float height) {
			cell.MinimumHeight = height;
			return cell;
		}

		public static PdfPCell bgColor(this PdfPCell cell, BaseColor color) {
			cell.BackgroundColor = color;
			return cell;
		}
	}

}
