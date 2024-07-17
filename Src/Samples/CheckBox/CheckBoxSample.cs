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
using System.IO;
using System.Linq;

namespace Xceed.Words.NET.Examples
{
  public class CheckBoxSample
  {
    #region Private Members

    private const string CheckBoxSampleResourcesDirectory = Program.SampleDirectory + @"CheckBox\Resources\";
    private const string CheckBoxSampleOutputDirectory = Program.SampleDirectory + @"CheckBox\Output\";

    #endregion

    #region Constructors

    static CheckBoxSample()
    {
      if( !Directory.Exists( CheckBoxSample.CheckBoxSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( CheckBoxSample.CheckBoxSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Modify a checkbox in a document.
    /// </summary>
    public static void ModifyCheckBox()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tModifyCheckBox()" );

      // Load a document
      using( var document = DocX.Load( CheckBoxSample.CheckBoxSampleResourcesDirectory+ @"DocumentWithCheckBoxes.docx" ) )
      {
        // Get the bookmark associated to a specific paragraph.
        var canWriteBookmark = document.Bookmarks[ "CanWrite_0_100" ];
        if( canWriteBookmark != null )
        {
          // get the checkBox from that paragraph.
          var canWriteCheckBox = canWriteBookmark.Paragraph.CheckBoxes.FirstOrDefault();
          if( canWriteCheckBox != null )
          {
            // Check the checkBox.
            canWriteCheckBox.IsChecked = true;
          }
        }        

        document.SaveAs( CheckBoxSample.CheckBoxSampleOutputDirectory + @"ModifyCheckBox.docx" );
        Console.WriteLine( "\tCreated: ModifyCheckBox.docx\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Add a checkbox in a document.
    /// </summary>
    public static void AddCheckBox()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tAddCheckBox()" );

      // Load a document
      using( var document = DocX.Load( CheckBoxSample.CheckBoxSampleResourcesDirectory + @"DocumentWithCheckBoxes.docx" ) )
      {
        // Insert a paragraph.
        document.InsertParagraph( "Student completes work neatly\t\t\t\t\t\t\t" );
        // Create a checkBox.
        var checkBox = document.AddCheckBox( true );
        // Add the checkBox to the last paragraph of the document.
        var p = document.Paragraphs.Last();        
        p.AppendCheckBox( checkBox );

        document.SaveAs( CheckBoxSample.CheckBoxSampleOutputDirectory + @"AddCheckBox.docx" );
        Console.WriteLine( "\tCreated: AddCheckBox.docx\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    #endregion
  }
}
