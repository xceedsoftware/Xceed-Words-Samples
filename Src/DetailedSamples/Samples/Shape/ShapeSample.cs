/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Section Sample Application
Copyright (c) 2009-2024 - Xceed Software Inc.
 
This application demonstrates how to insert sections when using the API 
from the Xceed Words for .NET.
 
This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using Xceed.Document.NET;
using System.Linq;

namespace Xceed.Words.NET.Examples
{
  public class ShapeSample
  {
    #region Private Members

    private const string ShapeSampleOutputDirectory = Program.SampleDirectory + @"Shape\Output\";

    #endregion

    #region Constructors

    static ShapeSample()
    {
      if( !Directory.Exists( ShapeSample.ShapeSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( ShapeSample.ShapeSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Create a document and insert shapes and paragraphs into it.
    /// </summary>
    public static void AddShape()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tAddShape()" );

      // Create a document.
      using( var document = DocX.Create( ShapeSample.ShapeSampleOutputDirectory + @"AddShape.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Inserting shapes" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Adding first shape.
        var shape = document.AddShape( 50.5f, 50.5f );

        // Create a paragraph and insert the shape at its 16th character.
        var p = document.InsertParagraph( "Here is a simple default rectangle positioned on the 16th character of this paragraph." );
        p.InsertShape( shape, 16 );
        p.SpacingAfter( 30 );

        // adding second shape.
        var shape2 = document.AddShape( 100f, 0f );
        shape2.FillColor = Color.Orange;
        shape2.Height = 175;
        shape2.OutlineColor = Color.Black;
        shape2.OutlineWidth = 4f;
        shape2.OutlineDash = DashStyle.Dot;

        // Create a paragraph and append the shape to it.
        var p2 = document.InsertParagraph( "Here is another rectangle appended to this paragraph: " );
        p2.AppendShape( shape2 );

        // Modify OutlineColor from shape in second paragraph.
        p2.Shapes.First().OutlineColor = Color.Red;

        document.Save();
        Console.WriteLine( "\tCreated: AddShape.docx\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Create a document and insert wrapping shapes and paragraphs into it.
    /// </summary>
    public static void AddShapeWithTextWrapping()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tAddShapeWithTextWrapping()" );

      // Create a document.
      using( var document = DocX.Create( ShapeSample.ShapeSampleOutputDirectory + @"AddShapeWithTextWrapping.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Add shapes with Text Wrapping" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add a shape and set its wrapping as Square.
        var shape = document.AddShape( 45f, 45f, Color.LightGray );
        shape.WrapStyle = WrapStyle.WrapSquare;
        shape.WrapTextPosition = WrapText.bothSides;
        // Set horizontal alignment with Alignment centered on the page.
        shape.HorizontalAlignment = WrappingHorizontalAlignment.CenteredRelativeToPage;
        // Set vertical alignment with an offset from top of paragraph.
        shape.VerticalOffsetAlignmentFrom = WrappingVerticalOffsetAlignmentFrom.Paragraph;
        shape.VerticalOffset = 20d;
        // Set a buffer on left and right of shape where no text will be drawn.
        shape.DistanceFromTextLeft = 5d;
        shape.DistanceFromTextRight = 5d;

        // Create a paragraph and append the shape to it.
        var p = document.InsertParagraph( "With its easy to use API, Xceed Words for .NET lets your application create new Microsoft Word .docx or PDF documents, or modify existing .docx documents. It gives you complete control over all content in a Word document, and lets you add or remove all commonly used element types, such as paragraphs, bulleted or numbered lists, images, tables, charts, headers and footers, sections, bookmarks, and more. Create PDF documents using the same API for creating Word documents." );
        p.Alignment = Alignment.both;
        p.AppendShape( shape );
        p.SpacingAfter( 50 );

        // Add another shape and set its wrapping as Top Bottom.
        var shape2 = document.AddShape( 250f, 25f, Color.DarkBlue, Color.LightBlue, 3f );
        shape2.WrapStyle = WrapStyle.WrapTopAndBottom;
        // Set horizontal alignment with Alignement centered on the page.
        shape2.HorizontalAlignment = WrappingHorizontalAlignment.CenteredRelativeToPage;
        // Set vertical alignment with an offset from top of page.
        shape2.VerticalOffsetAlignmentFrom = WrappingVerticalOffsetAlignmentFrom.Page;
        shape2.VerticalOffset = 320d;

        // Create a paragraph and append the shape to it.
        var p2 = document.InsertParagraph( "Xceed Words for .NET lets you create company reports that you first design with the familiar and rich editing capabilities of Microsoft Word instead of with a reporting tool’s custom editor. Use the designed document as a created template that you will programmatically customize before sending each report out.  You can also use Xceed Words for .NET to programmatically create invoices, add data to documents, perform mail merge functionality, and more. Based on our popular CodePlex project, known as DocX, it has benefited from 7 years of widespread use and has been downloaded over 250,000 times there and on NuGet. The large user base has resulted in abundant comments, requests and bug reports which are used to improve the library. You also get complete control over the document’s properties, including margins, page size, line spacing, page numbering, text direction and alignment, indentation, and more. You can also quickly and easily set or modify any text’s formatting, fonts and font sizes, colors, boldness, underline, italics, strikethrough, highlighting, and more. Search and replace text, add or remove password protection, join documents, copy documents, or apply templates – everything your application may need to do. It even supports modifying many Word files in parallel for greater speed." );
        p2.FontSize( 8d );
        p2.Alignment = Alignment.both;
        p2.AppendShape( shape2 );
        p2.SpacingAfter( 30 );

        document.Save();
        Console.WriteLine( "\tCreated: AddShapeWithTextWrapping.docx\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Create a document and insert a TextBox and paragraphs into it.
    /// </summary>
    public static void AddTextBox()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tAddTextBox()" );

      // Create a document.
      using( var document = DocX.Create( ShapeSample.ShapeSampleOutputDirectory + @"AddTextBox.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Add TextBox" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a TextBox with text formatting.
        var textBox = document.AddTextBox( 100f, 100f, "My TextBox", new Formatting() { FontColor = Color.Green } );
        textBox.TextVerticalAlignment = VerticalAlignment.Bottom;
        textBox.TextMarginBottom = 5d;
        textBox.TextMarginTop = 5d;
        textBox.TextMarginLeft = 5d;
        textBox.TextMarginRight = 5d;

        // Create a paragraph and insert the textBox at its 16th character.
        var p = document.InsertParagraph( "Here is a simple TextBox positioned on the 16th character of this paragraph." );
        p.InsertShape( textBox, 16 );
        p.SpacingAfter( 30 );

        // Add a bold paragraph to the TextBox.
        document.TextBoxes[ 0 ].InsertParagraph( "My New Paragraph" ).Bold();

        document.Save();
        Console.WriteLine( "\tCreated: AddTextBox.docx\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Create a document and insert wrapping textboxes and paragraphs into it.
    /// </summary>
    public static void AddTextBoxWithTextWrapping()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tAddTextBoxWithTextWrapping()" );

      // Create a document.
      using( var document = DocX.Create( ShapeSample.ShapeSampleOutputDirectory + @"AddTextBoxWithTextWrapping.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Add TextBoxes with Text Wrapping" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add a textBox and set its wrapping as Square.
        var textBox = document.AddTextBox( 45f, 45f, "My TextBox", new Formatting() { FontColor = Color.Yellow }, Color.Blue, Color.LightBlue );
        textBox.WrapStyle = WrapStyle.WrapSquare;
        textBox.WrapTextPosition = WrapText.bothSides;
        // Set horizontal alignment with alignment relative to the right of the margin.
        textBox.HorizontalAlignment = WrappingHorizontalAlignment.RightRelativeToMargin;
        // Set vertical alignment with an offset from top of paragraph.
        textBox.VerticalOffsetAlignmentFrom = WrappingVerticalOffsetAlignmentFrom.Paragraph;
        textBox.VerticalOffset = 20d;
        // Set a buffer on left and right of textBox where no text will be drawn.
        textBox.DistanceFromTextLeft = 5d;
        textBox.DistanceFromTextRight = 5d;
        // Set the text properties
        textBox.TextVerticalAlignment = VerticalAlignment.Bottom;
        textBox.TextMarginBottom = 5d;
        textBox.TextMarginTop = 5d;
        textBox.TextMarginLeft = 5d;
        textBox.TextMarginRight = 5d;
        textBox.IsTextWrap = false;

        // Create a paragraph and append the textbox to it.
        var p = document.InsertParagraph( "With its easy to use API, Xceed Words for .NET lets your application create new Microsoft Word .docx or PDF documents, or modify existing .docx documents. It gives you complete control over all content in a Word document, and lets you add or remove all commonly used element types, such as paragraphs, bulleted or numbered lists, images, tables, charts, headers and footers, sections, bookmarks, and more. Create PDF documents using the same API for creating Word documents." );
        p.Alignment = Alignment.both;
        p.AppendShape( textBox );
        p.SpacingAfter( 50 );

        // Add a shape with text and set its wrapping as Top Bottom.
        var shape = document.AddShape( 200f, 30f, Color.Orange, Color.Brown, 1.5f );
        shape.WrapStyle = WrapStyle.WrapTopAndBottom;
        // Set horizontal alignment with Alignment centered on the page.
        shape.HorizontalAlignment = WrappingHorizontalAlignment.CenteredRelativeToPage;
        // Set vertical alignment with an offset from top of page.
        shape.VerticalOffsetAlignmentFrom = WrappingVerticalOffsetAlignmentFrom.Page;
        shape.VerticalOffset = 320d;
        // Add text in shape
        shape.InsertParagraph( "Text in shape" ).Color( Color.Red ).Alignment = Alignment.center;

        // Create a paragraph and append the shape to it.
        var p2 = document.InsertParagraph( "Xceed Words for .NET lets you create company reports that you first design with the familiar and rich editing capabilities of Microsoft Word instead of with a reporting tool’s custom editor. Use the designed document as a created template that you will programmatically customize before sending each report out.  You can also use Xceed Words for .NET to programmatically create invoices, add data to documents, perform mail merge functionality, and more. Based on our popular CodePlex project, known as DocX, it has benefited from 7 years of widespread use and has been downloaded over 250,000 times there and on NuGet. The large user base has resulted in abundant comments, requests and bug reports which are used to improve the library. You also get complete control over the document’s properties, including margins, page size, line spacing, page numbering, text direction and alignment, indentation, and more. You can also quickly and easily set or modify any text’s formatting, fonts and font sizes, colors, boldness, underline, italics, strikethrough, highlighting, and more. Search and replace text, add or remove password protection, join documents, copy documents, or apply templates – everything your application may need to do. It even supports modifying many Word files in parallel for greater speed." );
        p2.FontSize( 8d );
        p2.Alignment = Alignment.both;
        p2.AppendShape( shape );
        p2.SpacingAfter( 30 );

        document.Save();
        Console.WriteLine( "\tCreated: AddTextBoxWithTextWrapping.docx\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    #endregion
  }
}
