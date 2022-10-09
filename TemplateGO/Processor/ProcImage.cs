using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateGO.Parser;
using TemplateGO.Utils;
using A = DocumentFormat.OpenXml.Drawing;
using Path = System.IO.Path;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace TemplateGO.Processor
{
    internal class ProcImage : BaseProcess, IProcessor
    {
        void IProcessor.Process(ProcessParams p)
        {
            // 数据中指定的值
            object? value = GetValueAndTransform(p.Data, p.Parser, p.Options);

            // 设置单元格内容（空值）
            SetCellValue(p.Cell, p.OriginValue, p.Parser, null, p.SharedStringTable);

            var options = new ImageOptions(p.Parser.Options);

            // 处理图片内容
            if (value != null && value.GetType() != typeof(string))
            {
                throw new ArgumentException($"图片属性只能为字符串或null {p.Parser.Property}");
            }

            // 加载到本地的图片
            var image = GetImageLocalFile(value as string, options, p);
            if (string.IsNullOrEmpty(image)) return;

            // 如果要求删除标记用的图片则先删除
            if (options.DeleteMarked)
            {
                DeleteImage(p.Cell, p.WorksheetPart);
            }

            // 添加图片
            AddImage(p.Cell, p.WorksheetPart, image, options);
        }

        protected virtual string? GetImageLocalFile(string? image, ImageOptions options, ProcessParams p)
        {
            // 预处理图片
            string? finalImage = image;
            if (p.Options.PreLoadImage != null)
            {
                finalImage = p.Options.PreLoadImage(finalImage, p.Parser.Property, options);
            }
            if (string.IsNullOrEmpty(finalImage)) return finalImage;

            return ImageUtils.ToLocalFile(finalImage);
        }

        private void DeleteImage(Cell cell, WorksheetPart worksheetPart)
        {
            var row = CellUtils.RowValue(cell.CellReference!) - 1;
            var column = CellUtils.ColumnValue(cell.CellReference!) - 1;

            // 删除 位于 cell 之上的所有图片
            var pics = worksheetPart.DrawingsPart?.WorksheetDrawing.Descendants<Xdr.Picture>().Where((r) =>
            {
                var from = r.Parent?.GetFirstChild<Xdr.FromMarker>();
                if (from == null) return false;
                return from.RowId!.Text == row.ToString() && from.ColumnId!.Text == column.ToString();
            }).ToList();
            if (pics == null) return;

            foreach (var pic in pics)
            {
                var blip = pic.Descendants<A.Blip>().FirstOrDefault();
                if (blip != null && blip.Embed != null)
                {
                    var partId = blip.Embed.Value;
                    if (!string.IsNullOrEmpty(partId))
                    {
                        // 图片可能公用，只有引用为1时删除（其它Sheet不受影响）
                        var count = worksheetPart.DrawingsPart?.WorksheetDrawing.Descendants<A.Blip>().Count(r => r.Embed?.Value == partId);
                        if (count <= 1)
                        {
                            worksheetPart.DrawingsPart?.DeletePart(partId);
                        }
                    }
                }

                // 删除图片
                pic.Parent?.Remove();
            }
        }

        private void AddImage(Cell cell, WorksheetPart worksheetPart, string imageFile, ImageOptions options)
        {
            var drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            if (!worksheetPart.Worksheet.ChildElements.OfType<Drawing>().Any())
            {
                worksheetPart.Worksheet.Append(new Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            drawingsPart.WorksheetDrawing ??= new WorksheetDrawing();
            var worksheetDrawing = drawingsPart.WorksheetDrawing!;

            // 插入图片
            var imagePart = drawingsPart.AddImagePart(ImageUtils.GetImagePartType(Path.GetExtension(imageFile)));
            using (var fs = new FileStream(imageFile, FileMode.Open))
            {
                imagePart.FeedData(fs);
            }

            // 位置与大小
            var rowNumber = CellUtils.RowValue(cell.CellReference!);
            var colNumber = CellUtils.ColumnValue(cell.CellReference!);
            var shape = ImageUtils.GetImageShape(options, imageFile);
            // 对齐方式暂时只能居中
            if (options.FrameWidth != null && options.FrameHeight != null)
            {
                shape = ImageUtils.MoveToCenter(shape, options.FrameWidth.Value, options.FrameHeight.Value);
            }

            var nvps = worksheetDrawing.Descendants<NonVisualDrawingProperties>();
            var nvpId = nvps.Count() > 0 ?
                (UInt32Value)worksheetDrawing.Descendants<NonVisualDrawingProperties>().Max(p => p.Id!.Value) + 1 :
                1U;

            var oneCellAnchor = new OneCellAnchor(
                new Xdr.FromMarker
                {
                    ColumnId = new ColumnId((colNumber - 1).ToString()),
                    RowId = new RowId((rowNumber - 1).ToString()),
                    ColumnOffset = new ColumnOffset(shape.X.ToString()),
                    RowOffset = new RowOffset(shape.Y.ToString())
                },
                new Extent { Cx = shape.W, Cy = shape.H },
                new Xdr.Picture(
                    new NonVisualPictureProperties(
                        new NonVisualDrawingProperties { Id = nvpId, Name = "Picture " + nvpId },
                        new NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })
                    ),
                    new BlipFill(
                        new A.Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = A.BlipCompressionValues.Print },
                        new A.Stretch(new A.FillRectangle())
                    ),
                    new ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = shape.X, Y = shape.Y },
                            new A.Extents { Cx = shape.W, Cy = shape.H }
                        ),
                        new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }
                    )
                ),
                new ClientData()
            );

            worksheetDrawing.Append(oneCellAnchor);
        }
    }
}
