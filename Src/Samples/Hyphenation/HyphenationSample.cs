/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Hyphenation Sample Application
Copyright (c) 2009-2024 - Xceed Software Inc.
 
This application demonstrates how to add and update text hyphenation when using the API 
from the Xceed Words for .NET.
 
This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.IO;
using Xceed.Document.NET;

namespace Xceed.Words.NET.Examples
{
  public class HyphenationSample
  {
    #region Private Members

    private const string HyphenationSampleResourceDirectory = Program.SampleDirectory + @"Hyphenation\Resources\";
    private const string HyphenationSampleOutputDirectory = Program.SampleDirectory + @"Hyphenation\Output\";

    #endregion

    #region Constructors

    static HyphenationSample()
    {
      if( !Directory.Exists( HyphenationSample.HyphenationSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( HyphenationSample.HyphenationSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Create a paragraph and set text hyphenation.
    /// </summary>
    public static void CreateHyphenation()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tCreateHyphenation()" );

      // Create a document.
      using( var document = DocX.Create( HyphenationSample.HyphenationSampleOutputDirectory + @"Hyphenation.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Creating a hyphenated paragraph into a document" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Some content to populate into the document.
        const string text = "Microsoft Word includes many document creation tools and features various text formatting options, such as HYPHENATION. " +
                            "The Word standard for text wrap is no HYPHENATION. Each word that is too long to fit at the end of a line is moved to " +
                            "the next line. Perhaps you are creating a business document and want to use justified text or text that is flush on both edges. " +
                            "Justified text can be accomplished in Word by adjusting the space between the words; " +
                            "however, longer terms may cause the line to have larger spaces than desirable between the words. " +
                            "Non-justified text without HYPHENATION can result in undesirable spaces at the end of lines. " +
                            "Use Word’s automatic HYPHENATION option to PRESENT your clients with a visually appealing document that displays evenly spaced words.";

        // Insert a Paragraph into this document.
        var p = document.InsertParagraph();

        // Append some text.
        p.Append( text );

        // Enable text hyphenation with hyphenation options.
        document.Hyphenation = new Xceed.Document.NET.Hyphenation()
        {
          IsAutomatic = true,
          HyphenateCaps = true,
          HyphenationZone = 16.0f,
          ConsecutiveHyphensLimit = 0
        };

        document.Save();
        Console.WriteLine( "\tCreated: Hyphenation.docx\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Update document hyphenation options.
    /// </summary>
    public static void UpdateHyphenation()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tUpdateHyphenation()" );

      // Load a document.
      using( var document = DocX.Load( HyphenationSample.HyphenationSampleResourceDirectory + @"HyphenatedDocument.docx" ) )
      {
        // Get and update the document hyphenation.
        var currentHyphenation = document.Hyphenation;
        currentHyphenation.HyphenateCaps = false;
        currentHyphenation.HyphenationZone = null;
        currentHyphenation.ConsecutiveHyphensLimit = 1;

        document.SaveAs( HyphenationSample.HyphenationSampleOutputDirectory + @"UpdatedHyphenatedDocument.docx" );
        Console.WriteLine( "\tCreated: UpdatedHyphenatedDocument.docx\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    #endregion
  }
}
