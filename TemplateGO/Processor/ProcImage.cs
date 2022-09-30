using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateGO.Parser;
using TemplateGO.Utils;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace TemplateGO.Processor
{
    internal class ProcImage : BaseProcess, IProcessor
    {
        void IProcessor.Process(ProcessParams p)
        {
            // 数据中指定的值
            object? value = GetValueByProperty(p.Data, p.Parser.Property);

            // 设置单元格内容（空值）
            SetCellValue(p.Cell, p.OriginValue, p.Parser, null, p.SharedStringTable);

            var options = new ImageOptions(p.Parser.Options);

            // 无数据时清空
            if (value != null && value.GetType() == typeof(string) && !string.IsNullOrEmpty(value as string))
            {
                AddImage(p.Cell, p.WorksheetPart, $"{value}", options);
            }
            else
            {
                RemoveImage(p.Cell, p.WorksheetPart);
            }
        }

        private void RemoveImage(Cell cell, WorksheetPart worksheetPart) { }

        private void AddImage(Cell cell, WorksheetPart worksheetPart, string image, ImageOptions options)
        {
            var drawingsPart = worksheetPart.DrawingsPart;
            if (drawingsPart == null) drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            if (!worksheetPart.Worksheet.ChildElements.OfType<Drawing>().Any())
            {
                worksheetPart.Worksheet.Append(new Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            if (drawingsPart.WorksheetDrawing == null) drawingsPart.WorksheetDrawing = new WorksheetDrawing();
            var worksheetDrawing = drawingsPart.WorksheetDrawing!;

            // 插入图片
            var imageFile = ImageUtils.ToLocalFile(image);
            var imagePart = drawingsPart.AddImagePart(ImageUtils.GetImagePartType(Path.GetExtension(imageFile)));
            using (var imageFs = new FileStream(imageFile, FileMode.Open))
            {
                imagePart.FeedData(imageFs!);
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
