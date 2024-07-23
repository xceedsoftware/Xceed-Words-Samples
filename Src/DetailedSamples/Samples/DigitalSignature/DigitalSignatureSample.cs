/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – DigitalSignature Sample Application
Copyright (c) 2009-2024 - Xceed Software Inc.
 
This application demonstrates how to digitally sign a document when using the API 
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
  public class DigitalSignatureSample
  {
    #region Private Members

    private const string DigitalSignatureSampleOutputDirectory = Program.SampleDirectory + @"DigitalSignature\Output\";
    private const string DigitalSignatureSampleResourcesDirectory = Program.SampleDirectory + @"DigitalSignature\Resources\";

    #endregion

    #region Constructors

    static DigitalSignatureSample()
    {
      if( !Directory.Exists( DigitalSignatureSample.DigitalSignatureSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( DigitalSignatureSample.DigitalSignatureSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Create a document, add 2 SignatureLines and digitally sign it.
    /// </summary>
    public static void SignWithSignatureLine()
    {
#if !OPEN_SOURCE && NETFRAMEWORK
      Console.WriteLine( "\tSignWithSignatureLine()" );

      SignatureLine signatureLine1;
      SignatureLine signatureLine2;
      SignatureLine signatureLine3;

      // Create a document.
      using( var document = DocX.Create( DigitalSignatureSample.DigitalSignatureSampleOutputDirectory + @"SignWithSignatureLine.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Adding Signature Lines and signing document" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create text content for this document.
        document.InsertParagraph( "This is an important top-secret document which is confidential and must not be modified by anyone under any circumstances." )
                .FontSize( 20f )
                .SpacingAfter( 60f );

        // Create a first paragraph.
        var p1 = document.InsertParagraph( "This is the first SignatureLine: \n" );

        // Create a SignatureLineOptions for the first SignatureLine.
        var signatureLineOptions = new SignatureLineOptions()
        {
          AllowComments = true,
          ShowDate = true,
          Instructions = "Please sign this important document.",
          Signer = "Mark Stone",
          SignerTitle = "President",
          SignerEmail = "mark.stone@mycompany.com"
        };

        // Add the first SignatureLine to the document.
        signatureLine1 = document.AddSignatureLine( signatureLineOptions );

        // Insert the first SignatureLine in a document's paragraph.
        p1.AppendSignatureLine( signatureLine1 ).SpacingAfter( 50d );

        // Create a second paragraph.
        var p2 = document.InsertParagraph( "This is the second SignatureLine : \n" );

        // Create a second SignatureLineOptions for the second SignatureLine.
        signatureLineOptions = new SignatureLineOptions()
        {
          AllowComments = false,
          ShowDate = false,
          Instructions = "Please sign this top secret document.",
          Signer = "Jimmy Smith",
          SignerTitle = "Chief Marketing",
          SignerEmail = "jimmy.smith@abc.com"
        };

        // Add the second SignatureLine to the document.
        signatureLine2 = document.AddSignatureLine( signatureLineOptions );

        // Insert the 2nd SignatureLine in a document's paragraph.
        p2.AppendSignatureLine( signatureLine2 ).SpacingAfter( 50d );

        // Create a third paragraph.
        var p3 = document.InsertParagraph( "This is third SignatureLine, which is positioned on the page with Text Wrapping. Please note you may have to click in document in order to view the signature content. \n" );

        // Create a third SignatureLineOptions for the third SignatureLine.
        // Set the position of the third SignatureLine on the page with Text Wrapping.
        signatureLineOptions = new SignatureLineOptions()
        {
          AllowComments = false,
          ShowDate = true,
          Instructions = "This document needs to have a digital signature.",
          Signer = "Tom Sawyer",
          SignerTitle = "CEO",
          SignerEmail = "tom.sawyer@abc.com",
          WrapStyle = WrapStyle.WrapSquare,
          HorizontalAlignment = WrappingHorizontalAlignment.CenteredRelativeToPage,
          VerticalOffsetAlignmentFrom = WrappingVerticalOffsetAlignmentFrom.BottomMargin,
          VerticalOffset = -50d
        };

        // Add the third SignatureLine to the document.
        signatureLine3 = document.AddSignatureLine( signatureLineOptions );

        // Insert the 3rd SignatureLine in a document's paragraph.
        p3.AppendSignatureLine( signatureLine3 );

        document.Save();
        Console.WriteLine( "\tCreated: SignWithSignatureLine.docx\n" );
      }

      // Create a digital certificate in order to sign the document, by using a pfx file and its password.
      var certificate = DigitalCertificate.Create( DigitalSignatureSampleResourcesDirectory + "CustomCertificate.pfx", "" );

      // Create SignOptions for first SignatureLine.
      // Use the SignatureLineId to match the first SignatureLine id and set the image used to sign the first SignatureLine.
      var signOptions1 = new SignOptions()
      {
        SignatureLineId = signatureLine1.Id,
        SignatureLineImage = DigitalSignatureSampleResourcesDirectory + "MarkStoneSignature.png",
        Comments = "This document is now signed by Mark Stone."
      };    

      // Sign the document with the certificate and the first SignOptions, related to the first SignatureLine.
      // First and second paramaters are the same because we overwrite the initial document.
      DocX.Sign( DigitalSignatureSample.DigitalSignatureSampleOutputDirectory + @"SignWithSignatureLine.docx",
                  DigitalSignatureSample.DigitalSignatureSampleOutputDirectory + @"SignWithSignatureLine.docx",
                  certificate, 
                  signOptions1 );

      // Create SignOptions for second SignatureLine.
      // Use the SignatureLineId to match the second SignatureLine id and set the text used to sign the second SignatureLine.
      var signOptions2 = new SignOptions()
      {
        SignatureLineId = signatureLine2.Id,
        SignatureLineText = "J. Smith",
        Comments = "This document is now signed by Jummy Smith."
      };

      // Sign the document with the certificate and the second SignOptions, related to the second SignatureLine.
      // First and second paramaters are the same because we overwrite the initial document.
      DocX.Sign( DigitalSignatureSample.DigitalSignatureSampleOutputDirectory + @"SignWithSignatureLine.docx",
                  DigitalSignatureSample.DigitalSignatureSampleOutputDirectory + @"SignWithSignatureLine.docx",
                  certificate,
                  signOptions2 );

      // Create SignOptions for third SignatureLine.
      // Use the SignatureLineId to match the third SignatureLine id and set the image used to sign the third SignatureLine.
      var signOptions3 = new SignOptions()
      {
        SignatureLineId = signatureLine3.Id,
        SignatureLineImage = DigitalSignatureSampleResourcesDirectory + "TomSawyerSignature.png",
        Comments = "This document is now signed by Tom Sawyer."
      };

      // Sign the document with the certificate and the third SignOptions, related to the third SignatureLine.
      // First and second paramaters are the same because we overwrite the initial document.
      DocX.Sign( DigitalSignatureSample.DigitalSignatureSampleOutputDirectory + @"SignWithSignatureLine.docx",
                  DigitalSignatureSample.DigitalSignatureSampleOutputDirectory + @"SignWithSignatureLine.docx",
                  certificate,
                  signOptions3 );
#else
      // This option is available in .NET Framework and when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Create a document and sign it without SignatureLines.
    /// </summary>
    public static void SignWithoutSignatureLine()
    {
#if !OPEN_SOURCE && NETFRAMEWORK
      Console.WriteLine( "\tSignWithoutSignatureLine()" );

      // Create a document.
      using( var document = DocX.Create( DigitalSignatureSample.DigitalSignatureSampleOutputDirectory + @"SignWithoutSignatureLine.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Signing document without Signature Line" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create text content for this document.
        document.InsertParagraph( "This is an important top-secret document which is confidential and must not be modified by anyone under any circumstances." )
                .FontSize( 20f )
                .SpacingAfter( 150f );       

        document.Save();

        // Create a digital certificate in order to sign the document, by using a pfx file and its password.
        var certificate = DigitalCertificate.Create( DigitalSignatureSampleResourcesDirectory + "CustomCertificate.pfx", "" );

        // Sign the document with the certificate..
        // First and second paramaters are the same because we overwrite the initial document.
        DocX.Sign( DigitalSignatureSample.DigitalSignatureSampleOutputDirectory + @"SignWithoutSignatureLine.docx",
                   DigitalSignatureSample.DigitalSignatureSampleOutputDirectory + @"SignWithoutSignatureLine.docx",
                   certificate );       

        Console.WriteLine( "\tCreated: SignWithoutSignatureLine.docx\n" );
      }
#else
      // This option is available in .NET Framework and when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Load a document and verify the signatures and SignatureLines validity.
    /// </summary>
    public static void VerifySignatures()
    {
#if !OPEN_SOURCE && NETFRAMEWORK
      Console.WriteLine( "\tVerifySignature()" );

      // The document is signed.
      var isDocumentSigned = DocX.IsSigned( DigitalSignatureSample.DigitalSignatureSampleResourcesDirectory + @"SignedDocument.docx" );
      var isDocumentSignatureValid = DocX.AreSignaturesValid( DigitalSignatureSample.DigitalSignatureSampleResourcesDirectory + @"SignedDocument.docx" );

      Console.WriteLine( "\tdocument SignedDocument.docx is " 
                          + (isDocumentSigned ? "signed" : "not signed")
                          + " and signatures are " 
                          + ( isDocumentSignatureValid ? "valid" : "not valid" ) );

      int counter = 0;
      // Verify each document's Signatures.
      foreach( DigitalSignature signature in DocX.GetSignatures( DigitalSignatureSample.DigitalSignatureSampleResourcesDirectory + @"SignedDocument.docx" ) )
      {
        Console.WriteLine( "\t  Signature" + ++counter +  " details:" );
        Console.WriteLine( "\t\tIs valid: " + signature.IsValid );
        Console.WriteLine( "\t\tsigning comments: " + signature.Comments );
        Console.WriteLine( "\t\tTime of signing: " + signature.SignTime );
        Console.WriteLine( "\t\tOwner name: " + signature.CertificateOwner );
        Console.WriteLine( "\t\tIssuer name: " + signature.CertificateIssuer );
      }

      // Load a document.
      using( var document = DocX.Load( DigitalSignatureSample.DigitalSignatureSampleResourcesDirectory + @"SignedDocument.docx" ) )
      {
        counter = 0;
        // Verify each SignatureLine.
        foreach( SignatureLine signatureLine in document.SignatureLines )
        {
          var isSignatureLineSigned = document.SignatureLines[ 0 ].IsSigned;
          var isSignatureLineValid = document.SignatureLines[ 0 ].IsValid;

          Console.WriteLine( "\t  SignatureLine" + ++counter + " is "
                              + ( isSignatureLineSigned ? "signed" : "not signed" ) + " and "
                              + ( isSignatureLineValid ? "valid" : "not valid" ) );
        }
      }

      Console.WriteLine( "\n" );
#else
      // This option is available in .NET Framework and when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Load a document and remove all the signatures, but keep the SignatureLines.
    /// </summary>
    public static void RemoveSignatures()
    {
#if !OPEN_SOURCE && NETFRAMEWORK
      Console.WriteLine( "\tRemoveSignatures()" );

      var inputDocument = DigitalSignatureSample.DigitalSignatureSampleResourcesDirectory + @"SignedDocument.docx";
      var outputDocument = DigitalSignatureSample.DigitalSignatureSampleOutputDirectory + @"SignaturesRemoved.docx";

      // The document is signed.
      if( DocX.IsSigned( inputDocument ) )
      {
        // Remove all the signatures, but keep the SignatureLines.
        DocX.RemoveAllSignatures( inputDocument, outputDocument );

        using( var document = DocX.Load( outputDocument ) )
        {
          //Remove title.
          document.Paragraphs[ 0 ].Remove( false );
          // Add a title.
          document.InsertParagraph( 0, "Remove all the Signatures", false ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;
          // Save update document.
          document.Save();
        }

        Console.WriteLine( "\tCreated: SignaturesRemoved.docx\n" );
      }
#else
      // This option is available in .NET Framework and when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Load a document and remove all the signatureLines.
    /// </summary>
    public static void RemoveSignatureLines()
    {
#if !OPEN_SOURCE && NETFRAMEWORK
      Console.WriteLine( "\tRemoveSignatureLines()" );

      // Load a document.
      using( var document = DocX.Load( DigitalSignatureSample.DigitalSignatureSampleResourcesDirectory + @"SignedDocument.docx" ) )
      {
        //Remove title.
        document.Paragraphs[ 0 ].Remove( false );
        // Add a title.
        document.InsertParagraph( 0, "Remove all the SignatureLines", false ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Remove all SignatureLines from the document, and the related signatures if the signatureLines are signed.
        document.SignatureLines.ForEach( signatureLine => signatureLine.Remove() );

        document.SaveAs( DigitalSignatureSample.DigitalSignatureSampleOutputDirectory + @"SignatureLinesRemoved.docx" );
        Console.WriteLine( "\tCreated: SignatureLinesRemoved.docx\n" );
      }
#else
      // This option is available in .NET Framework and when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    #endregion
  }
}
