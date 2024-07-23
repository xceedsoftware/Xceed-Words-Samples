/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Image Sample Application
Copyright (c) 2009-2024 - Xceed Software Inc.
 
This application demonstrates how to create, copy or modify a picture when using the API 
from the Xceed Words for .NET.
 
This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using Xceed.Document.NET;

namespace Xceed.Words.NET.Examples
{
  public class ImageSample
  {
    #region Private Members

    private const string ImageSampleResourcesDirectory = Program.SampleDirectory + @"Image\Resources\";
    private const string ImageSampleOutputDirectory = Program.SampleDirectory + @"Image\Output\";

    #endregion

    #region Constructors

    static ImageSample()
    {
      if( !Directory.Exists( ImageSample.ImageSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( ImageSample.ImageSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Add a picture loaded from disk or stream to a document.
    /// </summary>
    public static void AddPicture()
    {
      Console.WriteLine( "\tAddPicture()" );

      // Create a document.
      using( var document = DocX.Create( ImageSample.ImageSampleOutputDirectory + @"AddPicture.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Adding Pictures" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add a simple image from disk.
        var image = document.AddImage( ImageSample.ImageSampleResourcesDirectory + @"balloon.jpg" );
        var picture = image.CreatePicture( 112.5f, 112.5f );
        var p = document.InsertParagraph( "- Here is a simple picture added from disk:\n" );
        p.AppendPicture( picture );

        // Insert incremental "Figure 1" under picture by Picture public method.
        picture.InsertCaptionAfterSelf( "Figure" );

        p.SpacingAfter( 40 );

        // Add a rotated image from disk and set some alpha( 0 to 1 ).
        var rotatedPicture = image.CreatePicture( 112f, 112f );
        rotatedPicture.Rotation = 25;
#if !OPEN_SOURCE
        rotatedPicture.Alpha = 0.52f;
#endif

        var p2 = document.InsertParagraph( "- Here is the same picture added from disk, but rotated:\n" );
        p2.AppendPicture( rotatedPicture );

        // Insert incremental "Figure 2" under picture by Paragraph public method.
        p2 = p2.InsertCaptionAfterSelf( "Figure" );

        p2.SpacingAfter( 40 );

        // Add a simple image from a stream
        var streamImage = document.AddImage( new FileStream( ImageSample.ImageSampleResourcesDirectory + @"balloon.jpg", FileMode.Open, FileAccess.Read ) );
        var pictureStream = streamImage.CreatePicture( 112f, 112f );
        var p3 = document.InsertParagraph( "- Here is the same picture added from a stream:\n" );
        p3.AppendPicture( pictureStream );

        // Insert incremental "Figure 3" under picture by Paragraph public method.
        p3.InsertCaptionAfterSelf( "Figure" );

        document.Save();
        Console.WriteLine( "\tCreated: AddPicture.docx\n" );
      }
    }

    public static void AddPictureWithTextWrapping()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tAddPictureWithTextWrapping()" );

      // Create a document.
      using( var document = DocX.Create( ImageSample.ImageSampleOutputDirectory + @"AddPictureWithTextWrapping.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Add Pictures with Text Wrapping" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add a simple image from disk and set its wrapping as Square.
        var image = document.AddImage( ImageSample.ImageSampleResourcesDirectory + @"WordsIcon.png" );
        var picture = image.CreatePicture( 45f, 45f );
        picture.WrapStyle = WrapStyle.WrapSquare;
        picture.WrapTextPosition = WrapText.bothSides;
        // Set horizontal alignment with Alignement centered on the page.
        picture.HorizontalAlignment = WrappingHorizontalAlignment.CenteredRelativeToPage;
        // Set vertical alignement with an offset from top of paragraph.
        picture.VerticalOffsetAlignmentFrom = WrappingVerticalOffsetAlignmentFrom.Paragraph;
        picture.VerticalOffset = 22d;
        // Set a buffer on left and right of picture where no text will be drawn.
        picture.DistanceFromTextLeft = 5d;
        picture.DistanceFromTextRight = 5d;

        var p = document.InsertParagraph( "With its easy to use API, Xceed Words for .NET lets your application create new Microsoft Word .docx or PDF documents, or modify existing .docx documents. It gives you complete control over all content in a Word document, and lets you add or remove all commonly used element types, such as paragraphs, bulleted or numbered lists, images, tables, charts, headers and footers, sections, bookmarks, and more. Create PDF documents using the same API for creating Word documents." );
        p.Alignment = Alignment.both;
        p.InsertPicture( picture );
        p.SpacingAfter( 30 );

        // Add another simple image from disk and set its wrapping as Through.
        var imageW = document.AddImage( ImageSample.ImageSampleResourcesDirectory + @"W.png" );
        var pictureW = imageW.CreatePicture( 131f, 177f );
        pictureW.WrapStyle = WrapStyle.WrapThrough;
        pictureW.WrapTextPosition = WrapText.bothSides;
        // Set horizontal alignment with Alignement centered on the page.
        pictureW.HorizontalAlignment = WrappingHorizontalAlignment.CenteredRelativeToPage;
        // Set vertical alignement with an offset from top of page.
        pictureW.VerticalOffsetAlignmentFrom = WrappingVerticalOffsetAlignmentFrom.Page;
        pictureW.VerticalOffset = 255d;
        // Define the wrap polygon where the text will not be drawn.
        // The wrap polygon  is used when the Picture's WrappingStyle property is set to WrapTight or WrapThrough.
        // The top left of a picture has the coordinates( 0, 0 ), while the bottom right has the coordinates( 21600, 21600 ).
        var pts = new List<Point>();
        pts.Add( new Point( 0, 0 ) );
        pts.Add( new Point( 4027, 21477 ) );
        pts.Add( new Point( 8695, 21353 ) );
        pts.Add( new Point( 10800, 7282 ) );
        pts.Add( new Point( 13912, 21847 ) );
        pts.Add( new Point( 18031, 21847 ) );
        pts.Add( new Point( 21875, 0 ) );
        pts.Add( new Point( 18305, 0 ) );
        pts.Add( new Point( 15651, 16539 ) );
        pts.Add( new Point( 12631, 0 ) );
        pts.Add( new Point( 9336, 0 ) );
        pts.Add( new Point( 6590, 16539 ) );
        pts.Add( new Point( 3661, 0 ) );
        pts.Add( new Point( 0, 0 ) );
        pictureW.WrapPolygon = pts;

        var p2 = document.InsertParagraph( "Xceed Words for .NET lets you create company reports that you first design with the familiar and rich editing capabilities of Microsoft Word instead of with a reporting tool’s custom editor. Use the designed document as a created template that you will programmatically customize before sending each report out.  You can also use Xceed Words for .NET to programmatically create invoices, add data to documents, perform mail merge functionality, and more. Based on our popular CodePlex project, known as DocX, it has benefited from 7 years of widespread use and has been downloaded over 250,000 times there and on NuGet. The large user base has resulted in abundant comments, requests and bug reports which are used to improve the library. You also get complete control over the document’s properties, including margins, page size, line spacing, page numbering, text direction and alignment, indentation, and more. You can also quickly and easily set or modify any text’s formatting, fonts and font sizes, colors, boldness, underline, italics, strikethrough, highlighting, and more. Search and replace text, add or remove password protection, join documents, copy documents, or apply templates – everything your application may need to do. It even supports modifying many Word files in parallel for greater speed." );
        p2.FontSize( 8d );
        p2.Alignment = Alignment.both;
        p2.InsertPicture( pictureW );
        p2.SpacingAfter( 30 );

        document.Save();
        Console.WriteLine( "\tCreated: AddPictureWithTextWrapping.docx\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Copy a picture from a paragraph.
    /// </summary>
    public static void CopyPicture()
    {
      Console.WriteLine( "\tCopyPicture()" );

      // Create a document.
      using( var document = DocX.Create( ImageSample.ImageSampleOutputDirectory + @"CopyPicture.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Copying Pictures" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add a paragraph containing an image.
        var image = document.AddImage( ImageSample.ImageSampleResourcesDirectory + @"balloon.jpg" );
        var picture = image.CreatePicture( 75f, 75f );
        var p = document.InsertParagraph( "This is the first paragraph. " );
        p.AppendPicture( picture );
        p.AppendLine("It contains an image added from disk.");
        p.SpacingAfter( 50 );

        // Add a second paragraph containing no image. 
        var p2 = document.InsertParagraph( "This is the second paragraph. " );
        p2.AppendLine( "It contains a copy of the image located in the first paragraph." ).AppendLine();

        // Extract the first Picture from the first Paragraph.
        var firstPicture = p.Pictures.FirstOrDefault();
        if( firstPicture != null )
        {
          // copy it at the end of the second paragraph.
          p2.AppendPicture( firstPicture );
        }

        document.Save();
        Console.WriteLine( "\tCreated: CopyPicture.docx\n" );
      }
    }

    /// <summary>
    /// Modify an image from a document by writing text into it.
    /// </summary>
    public static void ModifyImage()
    {
      Console.WriteLine( "\tModifyImage()" );

      // Open the document Input.docx.
      using( var document = DocX.Load( ImageSample.ImageSampleResourcesDirectory + @"Input.docx" ) )
      {
        // Add a title
        document.InsertParagraph( 0, "Modifying Image by adding text/circle into the following image", false ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Get the first image in the document.
        var image = document.Images.FirstOrDefault();
        if( image != null )
        {
          // Create a bitmap from the image.
          Bitmap bitmap;
          using( var stream = image.GetStream( FileMode.Open, FileAccess.ReadWrite ) )
          {
            bitmap = new Bitmap( stream );
          }
          // Get the graphic from the bitmap to be able to draw in it.
          var graphic = Graphics.FromImage( bitmap );
          if( graphic != null )
          {
            // Draw a string with a specific font, font size and color at (0,10) from top left of the image.
            graphic.DrawString( "@copyright", new System.Drawing.Font( "Arial Bold", 12 ), Brushes.Red, new PointF( 0f, 10f ) );
            // Draw a blue circle of 10x10 at (30, 5) from the top left of the image.
            graphic.FillEllipse( Brushes.Blue, 30, 5, 10, 10 );

            // Save this Bitmap back into the document using a Create\Write stream.
            bitmap.Save( image.GetStream( FileMode.Create, FileAccess.Write ), ImageFormat.Png );
          }
        }

        document.SaveAs( ImageSample.ImageSampleOutputDirectory + @"ModifyImage.docx" );
        Console.WriteLine( "\tCreated: ModifyImage.docx\n" );
      }
    }

    #endregion
  }
}
