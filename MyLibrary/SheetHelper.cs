using System;
using System.Web;
using System.Web.UI;
using System.Collections.Generic;
using System.Reflection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace MyLibrary {
	public class SheetHelper {
		ISheet _sheet = null;
		int _row = 0;
		int _cell = 0;

		public SheetHelper(ISheet sheet) {
			this._sheet = sheet;
		}

		public SheetHelper Row(int row) {
			this._row = row;
			return this;
		}

		public SheetHelper Cell(int cell) {
			this._cell = cell;
			return this;
		}

		public SheetHelper Pos(int row, int cell) {
			this._row = row;
			this._cell = cell;
			return this;
		}

		public SheetHelper SetValue(bool value) {
			this._sheet.GetRow(this._row).GetCell(this._cell).SetCellValue(value);
			return this;
		}
		public SheetHelper SetValue(DateTime value) {
			this._sheet.GetRow(this._row).GetCell(this._cell).SetCellValue(value);
			return this;
		}
		public SheetHelper SetValue(double value) {
			this._sheet.GetRow(this._row).GetCell(this._cell).SetCellValue(value);
			return this;
		}
		public SheetHelper SetValue(IRichTextString value) {
			this._sheet.GetRow(this._row).GetCell(this._cell).SetCellValue(value);
			return this;
		}
		public SheetHelper SetValue(string value) {
			this._sheet.GetRow(this._row).GetCell(this._cell).SetCellValue(value);
			return this;
		}

		public void CopyRow(ISheet sourceSheet, int row) {
			IRow sourceRow = sourceSheet.GetRow(row) as IRow;
			IRow newRow = this._sheet.GetRow(_row) as IRow;
			ICell oldCell, newCell;
			int i;

			if (newRow == null)
				newRow = this._sheet.CreateRow(_row) as IRow;

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
			for (i = 0; i < sourceSheet.NumMergedRegions; i++) {
				cellRangeAddress = sourceSheet.GetMergedRegion(i);
				if (cellRangeAddress.FirstRow == sourceRow.RowNum) {
					newCellRangeAddress = new NPOI.SS.Util.CellRangeAddress(newRow.RowNum,
																			(newRow.RowNum + (cellRangeAddress.LastRow - cellRangeAddress.FirstRow)),
																			cellRangeAddress.FirstColumn,
																			cellRangeAddress.LastColumn);
					this._sheet.AddMergedRegion(newCellRangeAddress);
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

		}
	}
}
