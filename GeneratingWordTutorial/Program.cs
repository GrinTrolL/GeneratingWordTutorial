using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace GeneratingWordTutorial
{
    public class Program
    {
        static void Main(string[] args)
        {
            var path = Path.Combine(
               "C:\\Users\\ksowul\\Desktop\\docxTesting\\", DateTime.Now.ToShortDateString() + " " + DateTime.Now.Hour.ToString() + " " + DateTime.Now.Minute.ToString() + " "+ DateTime.Now.Second.ToString() + " lol.docx");
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document, true))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                string headerPath = "C:\\Users\\ksowul\\Desktop\\docxTesting\\image1.jpeg";

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                mainPart.Document.Body = new Body();
                Paragraph paragraph = new Paragraph();
                mainPart.Document.Body.AppendChild(paragraph);
                Run run = paragraph.AppendChild(new Run());
                var Header = GetHeader(headerPath, mainPart);
                run.AppendChild(Header);
                wordDocument.Close();
            }
            
        }
        public static Drawing GetHeader(string imagePath, MainDocumentPart mainPart)
        {
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }
            var drawing = new Drawing(
                                new DW.Anchor(
                                    new DW.SimplePosition() { X=0,Y=0},
                                    new DW.HorizontalPosition( 
                                        new DW.PositionOffset("-1052195"))
                                    { RelativeFrom= HorizontalRelativePositionValues.Column},
                                    new DW.VerticalPosition(new DW.PositionOffset("-899895"))
                                    { RelativeFrom= VerticalRelativePositionValues.Paragraph},
                                    new DW.Extent() { Cx = 7574400L, Cy = 3412800L },
                                    new DW.EffectExtent()
                                    {
                                        LeftEdge = 0L,
                                        TopEdge = 0L,
                                        RightEdge = 0L,
                                        BottomEdge = 0L
                                    },
                                    new DW.WrapSquare { WrapText=WrapTextValues.BothSides},
                                    new DW.DocProperties()
                                    {
                                        Id = (UInt32Value)1U,
                                        Name = "Picture 1"
                                    },

                                    new DW.NonVisualGraphicFrameDrawingProperties(
                                        new A.GraphicFrameLocks() { NoChangeAspect = true }),
                                        new A.Graphic(
                                            new A.GraphicData(
                                                new PIC.Picture(
                                                    new PIC.NonVisualPictureProperties(
                                                        new PIC.NonVisualDrawingProperties()
                                                        {
                                                            Id = (UInt32Value)0U,
                                                            Name = "New Bitmap Image.jpg"
                                                        },
                                                        new PIC.NonVisualPictureDrawingProperties()),
                                                        new PIC.BlipFill(
                                                            new A.Blip(
                                                                new A.BlipExtensionList(
                                                                    new A.BlipExtension()
                                                                    {
                                                                        Uri =
                                                                           "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                                    })
                                                                )
                                                            {
                                                                Embed = mainPart.GetIdOfPart(imagePart),
                                                                CompressionState =
                                                                    A.BlipCompressionValues.Print
                                                            },
                                                                new A.Stretch(
                                                                    new A.FillRectangle())
                                                            ),
                                                            new PIC.ShapeProperties(
                                                                new A.Transform2D(
                                                                    new A.Offset() { X = 0L, Y = 0L },
                                                                    new A.Extents() { Cx = 7574400L, Cy = 3412800L }),
                                                                    new A.PresetGeometry(
                                                                        new A.AdjustValueList()
                                                                    )
                                                                    { Preset = A.ShapeTypeValues.Rectangle })
                                                )
                                            )
                                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                                    )
                                {
                                    DistanceFromTop = (UInt32Value)0U,
                                    DistanceFromBottom = (UInt32Value)0U,
                                    DistanceFromLeft = (UInt32Value)114300U,
                                    DistanceFromRight = (UInt32Value)114300U,
                                    SimplePos = false,
                                    RelativeHeight = 251658240,
                                    BehindDoc = false,
                                    Locked = false,
                                    LayoutInCell = false,
                                    AllowOverlap=true
                                }
                    );
            return drawing;
        }
    }
}
