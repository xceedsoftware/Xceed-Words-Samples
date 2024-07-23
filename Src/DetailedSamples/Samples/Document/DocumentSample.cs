/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Document Sample Application
Copyright (c) 2009-2024 - Xceed Software Inc.
 
This application demonstrates how to modify the content of a document when using the API 
from the Xceed Words for .NET.
 
This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Xceed.Document.NET;

namespace Xceed.Words.NET.Examples
{
  public class DocumentSample
  {
    #region Private Members

    private static Dictionary<string, string> _replacePatterns = new Dictionary<string, string>()
    {
        { "OPPONENT", "Pittsburgh Penguins" },
        { "GAME_TIME", "7:30pm" },
        { "GAME_NUMBER", "161" },
        { "DATE", "October 18 2016" },
    };

    private const string DocumentSampleResourcesDirectory = Program.SampleDirectory + @"Document\Resources\";
    private const string DocumentSampleOutputDirectory = Program.SampleDirectory + @"Document\Output\";

    #endregion

    #region Constructors

    static DocumentSample()
    {
      if( !Directory.Exists( DocumentSample.DocumentSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( DocumentSample.DocumentSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Load a document and replace texts following a replace pattern.
    /// </summary>
    public static void ReplaceTextWithText()
    {
      Console.WriteLine( "\tReplaceTextWithText()" );

      // Load a document.
      using( var document = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"ReplaceText.docx" ) )
      {
        // Check if some of the replace patterns are used in the loaded document.
        if( document.FindUniqueByPattern( @"<[\w \=]{4,}>", RegexOptions.IgnoreCase ).Count > 0 )
        {
          // Do the replacement of all the found tags and with green bold strings.
          var replaceTextOptions = new FunctionReplaceTextOptions()
          {
            FindPattern = "<(.*?)>",
            RegexMatchHandler = DocumentSample.ReplaceFunc,
            RegExOptions = RegexOptions.IgnoreCase,
            NewFormatting = new Formatting() { Bold = true, FontColor = System.Drawing.Color.Green }
          };
          document.ReplaceText( replaceTextOptions );

          // Save this document to disk.
          document.SaveAs( DocumentSample.DocumentSampleOutputDirectory + @"ReplacedText.docx" );
          Console.WriteLine( "\tCreated: ReplacedTextWithText.docx\n" );
        }
      }
    }

    /// <summary>
    /// Load a document and replace texts with images.
    /// </summary>
    public static void ReplaceTextWithObjects()
    {
      Console.WriteLine( "\tReplaceTextWithObjects()" );

      // Load a document.
      using( var document = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"ReplaceTextWithObjects.docx" ) )
      {
        // Create the image from disk and set its size.
        var image = document.AddImage( DocumentSample.DocumentSampleResourcesDirectory + @"2018.jpg" );
        var picture = image.CreatePicture( 175f, 325f );

        // Do the replacement of all the found tags with the specified image and ignore the case when searching for the tags.
        document.ReplaceTextWithObject( new ObjectReplaceTextOptions() { SearchValue = "<yEaR_IMAGE>", NewObject = picture, RegExOptions = RegexOptions.IgnoreCase } );

        // Create the hyperlink.
        var hyperlink = document.AddHyperlink( "(ref)", new Uri( "https://en.wikipedia.org/wiki/New_Year" ) );
        // Do the replacement of all the found tags with the specified hyperlink.
        document.ReplaceTextWithObject( new ObjectReplaceTextOptions() { SearchValue = "<year_link>", NewObject = hyperlink } );

        // Add a Table into the document and sets its values.
        var t = document.AddTable( 1, 2 );
        t.Design = TableDesign.DarkListAccent4;
        t.AutoFit = AutoFit.Window;
        t.Rows[ 0 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "xceed.com" );
        t.Rows[ 0 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "@copyright 2024" );
        document.ReplaceTextWithObject( new ObjectReplaceTextOptions() { SearchValue = "<year_table>", NewObject = t } );

        // Save this document to disk.
        document.SaveAs( DocumentSample.DocumentSampleOutputDirectory + @"ReplacedTextWithObjects.docx" );
        Console.WriteLine( "\tCreated: ReplacedTextWithObjects.docx\n" );
      }
    }

    /// <summary>
    /// Add custom properties to a document.
    /// </summary>
    public static void AddCustomProperties()
    {
      Console.WriteLine( "\tAddCustomProperties()" );

      // Create a new document.
      using( var document = DocX.Create( DocumentSample.DocumentSampleOutputDirectory + @"AddCustomProperties.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Adding Custom Properties to a document" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        //Add custom properties to document.
        document.AddCustomProperty( new CustomProperty( "CompanyName", "Xceed Software inc." ) );
        document.AddCustomProperty( new CustomProperty( "Product", "Xceed Words for .NET" ) );
        document.AddCustomProperty( new CustomProperty( "Address", "3141 Taschereau, Greenfield Park" ) );
        document.AddCustomProperty( new CustomProperty( "Date", DateTime.Now ) );

        // Add a paragraph displaying the number of custom properties.
        var p = document.InsertParagraph( "This document contains " ).Append( document.CustomProperties.Count.ToString() ).Append( " Custom Properties :" );
        p.SpacingAfter( 30 );

        // Display each propertie's name and value.
        foreach( var prop in document.CustomProperties )
        {
          document.InsertParagraph( prop.Value.Name ).Append( " = " ).Append( prop.Value.Value.ToString() ).AppendLine();
        }

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: AddCustomProperties.docx\n" );
      }
    }

    /// <summary>
    /// Add a template to a document.
    /// </summary>
    public static void ApplyTemplate()
    {
      Console.WriteLine( "\tApplyTemplate()" );

      // Create a new document.
      using( var document = DocX.Create( DocumentSample.DocumentSampleOutputDirectory + @"ApplyTemplate.docx" ) )
      {
        // The path to a template document,
        var templatePath = DocumentSample.DocumentSampleResourcesDirectory + @"Template.docx";

        document.DifferentOddAndEvenPages = true;

        // Apply a template to the document based on a path.
        document.ApplyTemplate( templatePath );

        // Add a paragraph at the end of the template.
        document.InsertParagraph( "This paragraph is not part of the template." ).FontSize( 15d ).UnderlineStyle( UnderlineStyle.singleLine ).SpacingBefore( 50d );

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: ApplyTemplate.docx\n" );
      }
    }

    /// <summary>
    /// Insert a document at the end of another document.
    /// </summary>
    public static void AppendDocument()
    {
      Console.WriteLine( "\tAppendDocument()" );

      // Load the first document.
      using( var document1 = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"First.docx" ) )
      {
        // Load the second document.
        using( var document2 = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"Second.docx" ) )
        {
          // Add a title
          document1.InsertParagraph( 0, "Append Document", false ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

          // Insert a document at the end of another document.
          // When true, document is added at the end. When false, document is added at beginning.
          document1.InsertDocument( document2, true, true );

#if !OPEN_SOURCE
          // Obtain the new page count.
          var pageCount = document1.GetPageCount();
#endif

          // Save this document to disk.
          document1.SaveAs( DocumentSample.DocumentSampleOutputDirectory + @"AppendDocument.docx" );
          Console.WriteLine( "\tCreated: AppendDocument.docx\n" );
        }
      }
    }

    /// <summary>
    /// Insert a document inside another document.
    /// </summary>
    public static void InsertDocument()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tInsertDocument()" );

      // Load the first document.
      using( var document1 = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"First.docx" ) )
      {
        // Load the second document.
        using( var document2 = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"Second.docx" ) )
        {
          // Add a title
          document1.InsertParagraph( 0, "Insert Document", false ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

          // Insert "Second.docx" in "First.docx".
          // When true, document is added at the end of the specified paragraph. When false, document is added at beginning of the paragraph.
          var pararagraph = document1.Paragraphs.FirstOrDefault( p => p.Text.Contains( "Simple paragraph" ) );

          if( pararagraph != null )
          {
            document1.InsertDocument( document2, pararagraph, true, true, MergingMode.Both );
          }

          // Save this document to disk.
          document1.SaveAs( DocumentSample.DocumentSampleOutputDirectory + @"InsertDocument.docx" );
          Console.WriteLine( "\tCreated: InsertDocument.docx\n" );
        }
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    public static void LoadDocumentWithFilename()
    {
      using( var doc = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"First.docx" ) )
      {
        // Add a title
        doc.InsertParagraph( 0, "Load Document with File name", false ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert a Paragraph into this document.
        var p = doc.InsertParagraph();

        // Append some text and add formatting.
        p.Append( "A small paragraph was added." );

        doc.SaveAs( DocumentSample.DocumentSampleOutputDirectory + @"LoadDocumentWithFilename.docx" );
      }
    }

    public static void LoadDocumentWithStream()
    {
      using( var fs = new FileStream( DocumentSample.DocumentSampleResourcesDirectory + @"First.docx", FileMode.Open, FileAccess.Read, FileShare.Read ) )
      {
        using( var doc = DocX.Load( fs ) )
        {
          // Add a title
          doc.InsertParagraph( 0, "Load Document with Stream", false ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

          // Insert a Paragraph into this document.
          var p = doc.InsertParagraph();

          // Append some text and add formatting.
          p.Append( "A small paragraph was added." );

          doc.SaveAs( DocumentSample.DocumentSampleOutputDirectory + @"LoadDocumentWithStream.docx" );
        }
      }
    }

    public static void LoadDocumentWithStringUrl()
    {
      using( var doc = DocX.Load( "https://calibre-ebook.com/downloads/demos/demo.docx" ) )
      {
        // Add a title
        doc.InsertParagraph( 0, "Load Document with string Url", false ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert a Paragraph into this document.
        var p = doc.InsertParagraph();

        // Append some text and add formatting.
        p.Append( "A small paragraph was added." );

        doc.SaveAs( DocumentSample.DocumentSampleOutputDirectory + @"LoadDocumentWithUrl.docx" );
      }
    }

    /// <summary>
    /// Create a document and add html text to it.
    /// </summary>
    public static void AddHtml()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tAddHtml()" );

      // Create a new document.
      using( var document = DocX.Create( DocumentSample.DocumentSampleOutputDirectory + @"AddHtml.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Add Html" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert a Paragraph into this document.
        var p = document.InsertParagraph();

        // Append some text and add formatting.
        p.Append( "This is a simple paragraph added to the docx document. The following content is formatted html text." );

        // Insert an html paragraph in the document, after paragraph p.
        document.InsertContent( "<p>This is an html paragraph <b>(part of this paragraph is bold)</b></p>", ContentType.Html, p );

        // Insert another Paragraph into this document.
        var p2 = document.InsertParagraph( "This is a simple paragraph added to the docx document. The following content is a formatted html table." ).SpacingBefore( 40d );

        // Insert an html table (with its header) in the document.
        document.InsertContent( "<head>" +
                                   "<style>" +
                                      "table {" +
                                      "font-family: arial;" +
                                      "border-collapse: collapse;" +
                                      "width: 100%;" +
                                      "}" +

                                      "td, th {" +
                                      "border: 8px solid #dddddd;" +
                                      "text-align: left;" +
                                      "padding: 8px; " +
                                      "}" +
                                  "</style> " +
                                  "</head>" +

                                  "<body>" +
                                      "<h2> HTML Table </h2>" +

                                      "<table>" +
                                        "<tr>" +
                                          "<th> Company </th>" +
                                          "<th> Contact </th>" +
                                          "<th> Country </th>" +
                                        "</tr>" +
                                        "<tr>" +
                                          "<td> Adidas </td>" +
                                          "<td> Maria Ruiz </td>" +
                                          "<td> Germany </td>" +
                                        "</tr>" +
                                        "<tr>" +
                                          "<td> Nike </td>" +
                                          "<td> Mike Chang </td>" +
                                          "<td> USA </td>" +
                                        "</tr>" +
                                        "<tr>" +
                                          "<td> Reebok </td>" +
                                          "<td> Christian Dupre </td>" +
                                          "<td> Canada </td>" +
                                        "</tr>" +
                                        "<tr>" +
                                          "<td> Sportscheck </td>" +
                                          "<td> Giusepe DiPaolo </td>" +
                                          "<td> Italia </td>" +
                                        "</tr>" +
                                        "<tr>" +
                                          "<td> Puma </td>" +
                                          "<td> John Carlsson </td>" +
                                          "<td> UK </td>" +
                                        "</tr>" +
                                      "</table>" +
                                    "</body>", ContentType.Html );

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: AddHtml.docx\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Create a document and add rtf text to it.
    /// </summary>
    public static void AddRtf()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tAddRtf()" );

      // Create a new document.
      using( var document = DocX.Create( DocumentSample.DocumentSampleOutputDirectory + @"AddRtf.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Add Rtf" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert a Paragraph into this document.
        var p = document.InsertParagraph();

        // Append some text and add formatting.
        p.Append( "This is a simple paragraph added to the docx document. The following content is formatted rtf text." ).SpacingAfter( 40d );

        // Insert an rtf paragraph in the document, after paragraph p.
        document.InsertContent( @"{\rtf1\ansi\ansicpg1252\uc1\htmautsp\deff2{\fonttbl{\f0\fcharset0 Times New Roman;}" +
                                @"{\f2\fcharset0 Segoe UI;}}{\colortbl\red0\green0\blue0;\red255\green255\blue255;" +
                                @"\red255\green0\blue0;}\loch\hich\dbch\pard\plain\ltrpar\itap0{\lang1033\fs18\f2\cf0" +
                                @"\cf0\ql{\f2 {\ltrch This is }{\lang4105\ltrch an }{\lang4105\fs24\cf2\ltrch rtf }" +
                                @"{\lang4105\ltrch text (}{\lang4105\b\i\ltrch part of it is italic-bold}{\lang4105\ltrch ).}" +
                                @"\li0\ri0\sa0\sb0\fi0\ql\par}}}", ContentType.Rtf, p );

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: AddRtf.docx\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Create a document and add html text from an html file.
    /// </summary>
    public static void AddHtmlFromFile()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tAddHtmlFromFile()" );

      // Create a new document.
      using( var document = DocX.Create( DocumentSample.DocumentSampleOutputDirectory + @"AddHtmlFromFile.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Add Html from file" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert html text from an html file.
        document.InsertDocument( DocumentSample.DocumentSampleResourcesDirectory + @"HtmlSample.html", ContentType.Html );

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: AddHtmlFromFile.docx\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Create a document and add rtf text from an rtf file.
    /// </summary>
    public static void AddRtfFromFile()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tAddRtfFromFile()" );

      // Create a new document.
      using( var document = DocX.Create( DocumentSample.DocumentSampleOutputDirectory + @"AddRtfFromFile.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Add Rtf from file" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert rtf text from an rtf file.
        document.InsertDocument( DocumentSample.DocumentSampleResourcesDirectory + @"RtfSample.rtf", ContentType.Rtf );

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: AddRtfFromFile.docx\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Load a document and replace texts with html content.
    /// </summary>
    public static void ReplaceTextWithHtml()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tReplaceTextWithHtml()" );

      // Load a document.
      using( var document = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"Template-2.docx" ) )
      {
        // HTML content to insert path.
        string htmlPath = DocumentSample.DocumentSampleResourcesDirectory + @"HtmlSample.html";

        // Read HTML content.
        string html = File.ReadAllText( htmlPath );

        // Create HtmlReplaceTextOptions Object in order to replace "<html_content>" with the read html content from the path.
        var htmlReplaceTextOptions = new HtmlReplaceTextOptions()
        {
          SearchValue = "<html_content>",
          NewValue = html
        };

        // Do the replacement in the document for all occurance of "<html_content>".
        document.ReplaceTextWithHTML( htmlReplaceTextOptions );

        // Save this document to disk.
        document.SaveAs( DocumentSample.DocumentSampleOutputDirectory + @"ReplaceTextWithHtml.docx" );
        Console.WriteLine( "\tCreated: ReplaceTextWithHtml.docx\n" );
      }
#else
 	      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    #endregion

    #region Private Methods

    private static string ReplaceFunc( string findStr )
    {
      if( _replacePatterns.ContainsKey( findStr ) )
      {
        return _replacePatterns[ findStr ];
      }
      return findStr;
    }

    #endregion
  }
}
