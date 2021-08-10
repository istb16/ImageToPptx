using System;
using System.Drawing;
using System.IO;

using Spire.Presentation;
using Spire.Presentation.Drawing;

using Svg;

namespace ImageToPptx
{
    class Program
    {
        static void Main(string[] args)
        {
            // パラメータが指定されているかチェック
            if (args == null || args.Length < 1)
            {
                Console.WriteLine("Need file path.");
                return;
            }

            foreach (var filepath in args)
            {
                ConvertPptx(filepath);
            }
        }

        static void ConvertPptx(string filepath)
        {
            // 有効なファイルが指定されているかチェック
            if (string.IsNullOrEmpty(filepath))
            {
                Console.WriteLine("Need file path.");
                return;
            }
            else if (!File.Exists(filepath))
            {
                Console.WriteLine("Not found.");
                return;
            }

            // 画像を取得
            Bitmap? img = null;
            try
            {
                // 画像読み込み
                img = ReadImageFile(filepath);
                if (img == null)
                {
                    Console.WriteLine("Can't convert.");
                    return;
                }

                // PowerPointを作成
                var pptx = new Presentation();
                // スライドのサイズを設定
                pptx.SlideSize.Size = img.Size;
                // スライドの背景画像を挿入
                var imageData = pptx.Images.Append(img);
                pptx.Slides[0].Shapes.AppendEmbedImage(
                    ShapeType.Rectangle,
                    imageData,
                    new RectangleF(0, 0, img.Width, img.Height));
                pptx.Slides[0].SlideBackground.Fill.FillType = FillFormatType.Picture;
                pptx.Slides[0].SlideBackground.Fill.PictureFill.Picture.EmbedImage = imageData;

                // ファイルを保存
                pptx.SaveToFile(Path.ChangeExtension(filepath, "pptx"), FileFormat.Pptx2013);
            }
            finally
            {
                img?.Dispose();
            }
        }

        static Bitmap? ReadImageFile(string filepath)
        {
            var ext = Path.GetExtension(filepath)?.ToLower();
            if (ext == ".svg") return ReadSvgFile(filepath);
            else return new Bitmap(filepath);
        }

        static Bitmap? ReadSvgFile(string filepath)
        {
            var doc = SvgDocument.Open(filepath);
            if (doc == null) return null;
            doc.Width = doc.Bounds.Width - doc.Bounds.X;
            doc.Height = doc.Bounds.Height - doc.Bounds.Y;

            // サイズ調整
            var ratio = doc.Height / doc.Width;
            if (doc.Width < 1280)
            {
                doc.Width = 1280;
                doc.Height = doc.Width * ratio;
            }
            return doc.Draw();
        }
    }
}
