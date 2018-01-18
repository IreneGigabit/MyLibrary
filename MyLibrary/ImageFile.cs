using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Drawing;

namespace MyLibrary {
	public class ImageFile {
		public string FileName = string.Empty;

		public byte[] BinaryData;

		public Stream getDataStream() {
			//Stream DataStream = new MemoryStream(BinaryData);
			return new MemoryStream(BinaryData);
		}

		public ImagePartType ImageType {
			get {
				var ext = Path.GetExtension(FileName).TrimStart('.').ToLower();
				switch (ext) {
					case "jpg":
						return ImagePartType.Jpeg;
					case "png":
						return ImagePartType.Png;
					case "bmp":
						return ImagePartType.Bmp;
				}
				throw new ApplicationException(string.Format("不支援的格式:{0}", ext));
			}
		}

		public int SourceWidth;
		public int SourceHeight;
		public decimal Width;
		public decimal Height;

		//public long WidthInEMU => Convert.ToInt64(Width * CM_TO_EMU);
		private long WidthInEMU = 0;
		public long GetWidthInEMU() {
			WidthInEMU = Convert.ToInt64(Width * CM_TO_EMU);
			return WidthInEMU;
		}

		//public long HeightInEMU => Convert.ToInt64(Height * CM_TO_EMU);
		private long HeightInEMU = 0;
		public long GetHeightInEMU() {
			HeightInEMU = Convert.ToInt64(Height * CM_TO_EMU);
			return HeightInEMU;
		}

		private const decimal INCH_TO_CM = 2.54M;
		private const decimal CM_TO_EMU = 360000M;
		public string ImageName;

		public ImageFile(string fileName, byte[] data, decimal scale) {
			if (fileName == "") {
				FileName = string.Format("IMG_{0}", Guid.NewGuid().ToString().Substring(0, 8));
				ImageName = FileName;
			} else {
				FileName = fileName;
				ImageName = string.Format("IMG_{0}", Guid.NewGuid().ToString().Substring(0, 8));
			}

			BinaryData = data;
			int dpi = 300;
			Bitmap img = new Bitmap(new MemoryStream(data));
			SourceWidth = img.Width;
			SourceHeight = img.Height;
			Width = ((decimal)SourceWidth) / dpi * scale * INCH_TO_CM;
			Height = ((decimal)SourceHeight) / dpi * scale * INCH_TO_CM;
		}

		public ImageFile(byte[] data) :
			this("", data, 1) {
		}

		public ImageFile(byte[] data, decimal scale) :
			this("", data, scale) {
		}

		public ImageFile(string fileName) :
			this(fileName, File.ReadAllBytes(fileName), 1) {
		}

		public ImageFile(string fileName, decimal scale) :
			this(fileName, File.ReadAllBytes(fileName), scale) {
		}
	}
}
