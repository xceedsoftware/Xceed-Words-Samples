/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – PDF Sample Application
Copyright (c) 2009-2024 - Xceed Software Inc.
 
This application demonstrates how to convert a docx file into a pdf file 
when using the API from the Xceed Words for .NET.
 
This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Xceed.Document.NET;

namespace Xceed.Words.NET.Examples
{
  public class PdfSample
  {
    #region Private Members

#if !OPEN_SOURCE
    private const string PdfSampleResourcesDirectory = Program.SampleDirectory + @"Pdf\Resources\";
    private const string PdfSampleOutputDirectory = Program.SampleDirectory + @"Pdf\Output\";
#endif

    #endregion

    #region Constructors

    static PdfSample()
    {
#if !OPEN_SOURCE
      if( !Directory.Exists( PdfSample.PdfSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( PdfSample.PdfSampleOutputDirectory );
      }
#endif
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Load a document and convert it to PDF.
    /// </summary>
    public static void ConvertToPDF()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tConvertToPDF()" );

      // Load a document
      using( var document = DocX.Load( PdfSample.PdfSampleResourcesDirectory + @"DocumentToConvert.docx" ) )
      {
        DocX.ConvertToPdf( document, PdfSample.PdfSampleOutputDirectory + @"ConvertedDocument.pdf" );
        Console.WriteLine( "\tCreated: ConvertedDocument.pdf\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Load a document with uninstalled font and convert it to PDF.
    /// </summary>
    public static void ConvertToPDFWithUninstalledFont()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tConvertToPDFWithUninstalledFont()" );

      // Load a document
      using( var document = DocX.Load( PdfSample.PdfSampleResourcesDirectory + @"DocumentToConvertWithUninstalledFont.docx" ) )
      {
        var extrernalFontList = new List<PdfExternalFont>()
        {
           new PdfExternalFont()
          {
            Name = "The Bugatten",
            Path = PdfSample.PdfSampleResourcesDirectory + @"The Bugatten.ttf"
          }
        };
        DocX.ConvertToPdf( document, PdfSample.PdfSampleOutputDirectory + @"ConvertedDocumentWithUninstalledFont.pdf", extrernalFontList );

        Console.WriteLine( "\tCreated: ConvertToPDFWithUninstalledFont.pdf\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    #endregion
  }
}
