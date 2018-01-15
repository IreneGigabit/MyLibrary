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
		//DataStream = new MemoryStream(BinaryData);
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
		DataStream = new MemoryStream(Convert.FromBase64String(ImageToBase64(img)));
	}

	//public ImageData(string fileName) :
	//	this(fileName, File.ReadAllBytes(fileName), 300) {
	//}
	public ImageData(string fileName) :
		this(fileName, File.ReadAllBytes(fileName), 300) {
	}

	/// <summary>
	/// 自動判斷圖片格式
	/// </summary>
	/// <param name="img"></param>
	/// <returns></returns>
	public static System.Drawing.Imaging.ImageFormat GetImageFormat(System.Drawing.Image img) {
		if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Jpeg))
			return System.Drawing.Imaging.ImageFormat.Jpeg;
		if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Bmp))
			return System.Drawing.Imaging.ImageFormat.Bmp;
		if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Png))
			return System.Drawing.Imaging.ImageFormat.Png;
		if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Emf))
			return System.Drawing.Imaging.ImageFormat.Emf;
		if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Exif))
			return System.Drawing.Imaging.ImageFormat.Exif;
		if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Gif))
			return System.Drawing.Imaging.ImageFormat.Gif;
		if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Icon))
			return System.Drawing.Imaging.ImageFormat.Icon;
		if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.MemoryBmp))
			return System.Drawing.Imaging.ImageFormat.MemoryBmp;
		if (img.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.Tiff))
			return System.Drawing.Imaging.ImageFormat.Tiff;
		else
			return System.Drawing.Imaging.ImageFormat.Wmf;
	}

	/// <summary>
	/// 將 Image 物件轉 Base64
	/// </summary>
	/// <param name="image"></param>
	/// <returns></returns>
	public string ImageToBase64(System.Drawing.Image image) {
		MemoryStream ms = new MemoryStream();

		System.Drawing.Imaging.ImageFormat format = GetImageFormat(image);

		// 將圖片轉成 byte[]
		image.Save(ms, format);
		byte[] imageBytes = ms.ToArray();

		// 將 byte[] 轉 base64
		string base64String = Convert.ToBase64String(imageBytes);
		return base64String;
	}
}
