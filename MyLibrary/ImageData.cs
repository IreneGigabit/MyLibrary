using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Drawing;

public class ImageData
{
	public string FileName = string.Empty;

	public byte[] BinaryData;

	//public Stream DataStream => new MemoryStream(BinaryData);
	private Stream DataStream = null;

	public Stream getDataStream() {
		DataStream = new MemoryStream(BinaryData);
		return DataStream;
	}

	public ImagePartType ImageType
	{
		get
		{
			var ext = Path.GetExtension(FileName).TrimStart('.').ToLower();
			switch (ext) {
				case "jpg":
					return ImagePartType.Jpeg;
				case "png":
					return ImagePartType.Png;
				case "bmp":
					return ImagePartType.Bmp;
			}
			throw new ApplicationException(string.Format("不支援的格式:{0}",ext));
		}
	}

	public int SourceWidth;
	public int SourceHeight;
	public decimal Width;
	public decimal Height;

	//public long WidthInEMU => Convert.ToInt64(Width * CM_TO_EMU);
	private long WidthInEMU=0;
	public long GetWidthInEMU() {
		WidthInEMU = Convert.ToInt64(Width * CM_TO_EMU);
		return WidthInEMU;
	}

	//public long HeightInEMU => Convert.ToInt64(Height * CM_TO_EMU);
	private long HeightInEMU = 0;
	public long GetHeightInEMU() {
		HeightInEMU= Convert.ToInt64(Height * CM_TO_EMU);
		return HeightInEMU;
	}

	private const decimal INCH_TO_CM = 2.54M;
	private const decimal CM_TO_EMU = 360000M;
	public string ImageName;

	public ImageData(string fileName, byte[] data, int dpi) {
		FileName = fileName;
		BinaryData = data;

		Bitmap img = new Bitmap(new MemoryStream(data));
		SourceWidth = img.Width;
		SourceHeight = img.Height;
		Width = ((decimal)SourceWidth) / dpi * INCH_TO_CM;
		Height = ((decimal)SourceHeight) / dpi * INCH_TO_CM;
		//ImageName = $"IMG_{Guid.NewGuid().ToString().Substring(0, 8)}";
		ImageName = string.Format("IMG_{0}", Guid.NewGuid().ToString().Substring(0, 8));
	}

	public ImageData(string fileName) :
		this(fileName, File.ReadAllBytes(fileName), 300) {
	}
}
