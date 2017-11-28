using System;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;


namespace MyLibrary {
	public static class SheetExt {
		public static IRow Row(this ISheet sheet, int row) {
			IRow iRow=sheet.GetRow(row);
			if (iRow == null)
				iRow = sheet.CreateRow(row);
			return iRow;
		}

		public static ICell Cell(this IRow row, int cell) {
			ICell iCell = row.GetCell(cell);
			if (iCell == null)
				iCell = row.CreateCell(cell);
			return iCell;
		}

		public static ICell Pos(this ISheet sheet, int row, int cell) {
			ICell iCell = sheet.GetRow(row).GetCell(cell);
			return iCell;
		}

		public static void SetValue(this ICell cell, bool value) {
			cell.SetCellValue(value);
		}

		public static void SetValue(this ICell cell, DateTime value) {
			cell.SetCellValue(value);
		}

		public static void SetValue(this ICell cell, double value) {
			cell.SetCellValue(value);
		}

		public static void SetValue(this ICell cell, IRichTextString value) {
			cell.SetCellValue(value);
		}

		public static void SetValue(this ICell cell, string value) {
			cell.SetCellValue(value);
		}

		public static ISheet SetValue(this ISheet sheet, int row, int cell, bool value) {
			sheet.GetRow(row).GetCell(cell).SetCellValue(value);
			return sheet;
		}

		public static ISheet SetValue(this ISheet sheet, int row, int cell, DateTime value) {
			sheet.GetRow(row).GetCell(cell).SetCellValue(value);
			return sheet;
		}

		public static ISheet SetValue(this ISheet sheet, int row, int cell, double value) {
			sheet.GetRow(row).GetCell(cell).SetCellValue(value);
			return sheet;
		}

		public static ISheet SetValue(this ISheet sheet, int row, int cell, IRichTextString value) {
			sheet.GetRow(row).GetCell(cell).SetCellValue(value);
			return sheet;
		}

		public static ISheet SetValue(this ISheet sheet, int row, int cell, string value) {
			sheet.GetRow(row).GetCell(cell).SetCellValue(value);
			return sheet;
		}

		public static IRow CopyRow(this IRow dstRow, ISheet srcSheet, int srcRowNum) {
			IRow sourceRow = srcSheet.GetRow(srcRowNum) as IRow;
			ICell oldCell, newCell;
			int i;

			// Loop through source columns to add to new row
			for (i = 0; i < sourceRow.LastCellNum; i++) {
				// Grab a copy of the old/new cell
				oldCell = sourceRow.GetCell(i) as ICell;
				newCell = dstRow.GetCell(i) as ICell;

				if (newCell == null)
					newCell = dstRow.CreateCell(i) as ICell;

				// If the old cell is null jump to next cell
				if (oldCell == null) {
					newCell = null;
					continue;
				}

				// Copy style from old cell and apply to new cell
				newCell.CellStyle = oldCell.CellStyle;

				// If there is a cell comment, copy
				if (newCell.CellComment != null) newCell.CellComment = oldCell.CellComment;

				// If there is a cell hyperlink, copy
				if (oldCell.Hyperlink != null) newCell.Hyperlink = oldCell.Hyperlink;

				// Set the cell data value
				switch (oldCell.CellType) {
					case CellType.Blank:
						newCell.SetCellValue(oldCell.StringCellValue);
						break;
					case CellType.Boolean:
						newCell.SetCellValue(oldCell.BooleanCellValue);
						break;
					case CellType.Error:
						newCell.SetCellErrorValue(oldCell.ErrorCellValue);
						break;
					case CellType.Formula:
						newCell.CellFormula = oldCell.CellFormula;
						break;
					case CellType.Numeric:
						newCell.SetCellValue(oldCell.NumericCellValue);
						break;
					case CellType.String:
						newCell.SetCellValue(oldCell.RichStringCellValue);
						break;
					case CellType.Unknown:
						newCell.SetCellValue(oldCell.StringCellValue);
						break;
				}
			}

			//If there are are any merged regions in the source row, copy to new row
			NPOI.SS.Util.CellRangeAddress cellRangeAddress = null, newCellRangeAddress = null;
			for (i = 0; i < srcSheet.NumMergedRegions; i++) {
				cellRangeAddress = srcSheet.GetMergedRegion(i);
				if (cellRangeAddress.FirstRow == sourceRow.RowNum) {
					newCellRangeAddress = new NPOI.SS.Util.CellRangeAddress(dstRow.RowNum,
																			(dstRow.RowNum + (cellRangeAddress.LastRow - cellRangeAddress.FirstRow)),
																			cellRangeAddress.FirstColumn,
																			cellRangeAddress.LastColumn);
					dstRow.Sheet.AddMergedRegion(newCellRangeAddress);
				}
			}

			//複製行高到新列
			//if (copyRowHeight)
			dstRow.Height = sourceRow.Height;
			////重製原始列行高
			//if (resetOriginalRowHeight)
			//    sourceRow.Height = worksheet.DefaultRowHeight;
			////清掉原列
			//if (IsRemoveSrcRow == true)
			//    worksheet.RemoveRow(sourceRow);

			return dstRow;
		}


		public static ISheet CopyRow(this ISheet dstSheet, int dstRow, ISheet srcSheet, int srcRow) {
			IRow sourceRow = srcSheet.GetRow(srcRow) as IRow;
			IRow newRow = dstSheet.GetRow(dstRow) as IRow;
			ICell oldCell, newCell;
			int i;

			if (newRow == null)
				newRow = dstSheet.CreateRow(dstRow) as IRow;

			// Loop through source columns to add to new row
			for (i = 0; i < sourceRow.LastCellNum; i++) {
				// Grab a copy of the old/new cell
				oldCell = sourceRow.GetCell(i) as ICell;
				newCell = newRow.GetCell(i) as ICell;

				if (newCell == null)
					newCell = newRow.CreateCell(i) as ICell;

				// If the old cell is null jump to next cell
				if (oldCell == null) {
					newCell = null;
					continue;
				}

				// Copy style from old cell and apply to new cell
				newCell.CellStyle = oldCell.CellStyle;

				// If there is a cell comment, copy
				if (newCell.CellComment != null) newCell.CellComment = oldCell.CellComment;

				// If there is a cell hyperlink, copy
				if (oldCell.Hyperlink != null) newCell.Hyperlink = oldCell.Hyperlink;

				// Set the cell data value
				switch (oldCell.CellType) {
					case CellType.Blank:
						newCell.SetCellValue(oldCell.StringCellValue);
						break;
					case CellType.Boolean:
						newCell.SetCellValue(oldCell.BooleanCellValue);
						break;
					case CellType.Error:
						newCell.SetCellErrorValue(oldCell.ErrorCellValue);
						break;
					case CellType.Formula:
						newCell.CellFormula = oldCell.CellFormula;
						break;
					case CellType.Numeric:
						newCell.SetCellValue(oldCell.NumericCellValue);
						break;
					case CellType.String:
						newCell.SetCellValue(oldCell.RichStringCellValue);
						break;
					case CellType.Unknown:
						newCell.SetCellValue(oldCell.StringCellValue);
						break;
				}
			}

			//If there are are any merged regions in the source row, copy to new row
			NPOI.SS.Util.CellRangeAddress cellRangeAddress = null, newCellRangeAddress = null;
			for (i = 0; i < srcSheet.NumMergedRegions; i++) {
				cellRangeAddress = srcSheet.GetMergedRegion(i);
				if (cellRangeAddress.FirstRow == sourceRow.RowNum) {
					newCellRangeAddress = new NPOI.SS.Util.CellRangeAddress(newRow.RowNum,
																			(newRow.RowNum + (cellRangeAddress.LastRow - cellRangeAddress.FirstRow)),
																			cellRangeAddress.FirstColumn,
																			cellRangeAddress.LastColumn);
					dstSheet.AddMergedRegion(newCellRangeAddress);
				}
			}

			//複製行高到新列
			//if (copyRowHeight)
			newRow.Height = sourceRow.Height;
			////重製原始列行高
			//if (resetOriginalRowHeight)
			//    sourceRow.Height = worksheet.DefaultRowHeight;
			////清掉原列
			//if (IsRemoveSrcRow == true)
			//    worksheet.RemoveRow(sourceRow);

			return dstSheet;
		}
	}
}
