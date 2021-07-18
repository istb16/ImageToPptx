using System;
using System.Drawing;
using System.IO;

using Spire.Presentation;
using Spire.Presentation.Drawing;

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
            using var img = new Bitmap(filepath);

            // PowerPointを作成
            var pptx = new Presentation();
            // スライドのサイズを設定
            pptx.SlideSize.Size = img.Size;
            // スライドの背景画像を挿入
            pptx.Slides[0].SlideBackground.Fill.FillType = FillFormatType.Picture;
            pptx.Slides[0].SlideBackground.Fill.PictureFill.Picture.EmbedImage =
                pptx.Slides[0].Shapes.AppendEmbedImage(
                    ShapeType.Rectangle,
                    filepath,
                    new RectangleF(0, 0, img.Width, img.Height)) as IImageData;

            // ファイルを保存
            pptx.SaveToFile(Path.ChangeExtension(filepath, "pptx"), FileFormat.Pptx2013);
        }
    }
}
