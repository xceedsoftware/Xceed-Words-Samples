/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – List Sample Application
Copyright (c) 2009-2024 - Xceed Software Inc.
 
This application demonstrates how to add lists when using the API 
from the Xceed Words for .NET.
 
This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Linq;
using Xceed.Document.NET;

namespace Xceed.Words.NET.Examples
{
  public class ListSample
  {
    #region Private Members

    private const string ListSampleResourceDirectory = Program.SampleDirectory + @"List\Resources\";
    private const string ListSampleOutputDirectory = Program.SampleDirectory + @"List\Output\";

    #endregion

    #region Constructors

    static ListSample()
    {
      if( !Directory.Exists( ListSample.ListSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( ListSample.ListSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Create a numbered and a bulleted lists with different listItem's levels.
    /// </summary>
    public static void AddList()
    {
      Console.WriteLine( "\tAddList()" );

      // Create a document.
      using( var document = DocX.Create( ListSample.ListSampleOutputDirectory + @"AddList.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Adding lists into a document" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

#if !OPEN_SOURCE
        // Add a numbered list.
        var numberedList = document.AddList( new ListOptions() );
        // Add the first ListItem
        numberedList.AddListItem( "Berries", 0 );
        // Add Sub-items(level 1) to the preceding ListItem.
        numberedList.AddListItem( "Strawberries", 1 );
        numberedList.AddListItem( "Blueberries", 1 );
        numberedList.AddListItem( "Raspberries", 1 );
        // Add an item (level 0)
        numberedList.AddListItem( "Banana", 0 );
        // Add an item (level 0)
        numberedList.AddListItem( "Apple", 0 );
        // Add Sub-items(level 1) to the preceding ListItem.
        numberedList.AddListItem( "Red", 1 );
        numberedList.AddListItem( "Green", 1 );
        numberedList.AddListItem( "Yellow", 1 );
#else
        // Add a numbered list where the first ListItem is starting with number 1.
        var numberedList = document.AddList( "Berries", 0, ListItemType.Numbered, 1 );
        // Add Sub-items(level 1) to the preceding ListItem.
        document.AddListItem( numberedList, "Strawberries", 1 );
        document.AddListItem( numberedList, "Blueberries", 1 );
        document.AddListItem( numberedList, "Raspberries", 1 );
        // Add an item (level 0)
        document.AddListItem( numberedList, "Banana", 0 );
        // Add an item (level 0)
        document.AddListItem( numberedList, "Apple", 0 );
        // Add Sub-items(level 1) to the preceding ListItem.
        document.AddListItem( numberedList, "Red", 1 );
        document.AddListItem( numberedList, "Green", 1 );
        document.AddListItem( numberedList, "Yellow", 1 );
#endif
#if !OPEN_SOURCE
        // Add a bulleted list.
        var bulletedList = document.AddList( new ListOptions() { ListType = ListItemType.Bulleted } );
        // Add the first ListItem
        bulletedList.AddListItem( "Canada", 0 );
        // Add Sub-items(level 1) to the preceding ListItem.
        bulletedList.AddListItem( "Toronto", 1 );
        bulletedList.AddListItem( "Montreal", 1 );
        // Add an item (level 0)
        bulletedList.AddListItem( "Brazil" );
        // Add an item (level 0)
        bulletedList.AddListItem( "USA" );
        // Add Sub-items(level 1) to the preceding ListItem.
        bulletedList.AddListItem( "New York", 1 );
        // Add Sub-items(level 2) to the preceding ListItem.
        bulletedList.AddListItem( "Brooklyn", 2 );
        bulletedList.AddListItem( "Manhattan", 2 );
        bulletedList.AddListItem( "Los Angeles", 1 );
        bulletedList.AddListItem( "Miami", 1 );
        // Add an item (level 0)
        bulletedList.AddListItem( "France" );
        // Add Sub-items(level 1) to the preceding ListItem.
        bulletedList.AddListItem( "Paris", 1 );
#else
        // Add a bulleted list with its first item.
        var bulletedList = document.AddList( "Canada", 0, ListItemType.Bulleted );
        // Add Sub-items(level 1) to the preceding ListItem.
        document.AddListItem( bulletedList, "Toronto", 1 );
        document.AddListItem( bulletedList, "Montreal", 1 );
        // Add an item (level 0)
        document.AddListItem( bulletedList, "Brazil" );
        // Add an item (level 0)
        document.AddListItem( bulletedList, "USA" );
        // Add Sub-items(level 1) to the preceding ListItem.
        document.AddListItem( bulletedList, "New York", 1 );
        // Add Sub-items(level 2) to the preceding ListItem.
        document.AddListItem( bulletedList, "Brooklyn", 2 );
        document.AddListItem( bulletedList, "Manhattan", 2 );
        document.AddListItem( bulletedList, "Los Angeles", 1 );
        document.AddListItem( bulletedList, "Miami", 1 );
        // Add an item (level 0)
        document.AddListItem( bulletedList, "France" );
        // Add Sub-items(level 1) to the preceding ListItem.
        document.AddListItem( bulletedList, "Paris", 1 );
#endif
        // Insert the lists into the document.
        document.InsertParagraph( "This is a Numbered List:\n" );
        document.InsertList( numberedList );
        document.InsertParagraph().SpacingAfter( 40d );
        document.InsertParagraph( "This is a Bulleted List:\n" );
        document.InsertList( bulletedList, new Xceed.Document.NET.Font( "Cooper Black" ), 15 );

        document.Save();
        Console.WriteLine( "\tCreated: AddList.docx\n" );
      }
    }

    /// <summary>
    /// Create a custom numbered list with different listItem's levels, items and numbering formatting.
    /// </summary>
    public static void AddCustomNumberedList()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tAddCustomNumberedList()" );

      // Create a document.
      using( var document = DocX.Create( ListSample.ListSampleOutputDirectory + @"AddCustomNumberedList.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Adding a custom numbered list into a document" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add a custom numbered list.
        // Create the list levels configuration
        var customNumberedListOptions = new ListOptions()
        {
          ListType = ListItemType.Numbered,

          LevelsConfigs = new ObservableCollection<ListLevelConfig>()
            {
               new ListLevelConfig()
                   {
                      NumberingFormat = NumberingFormat.upperRoman,
                      NumberingLevelText = "%1)",
                      Justification = Justification.left,
                      Indentation = new Indentation()
                      {
                        Hanging = 18f,
                        Left = 16f
                      },
                      Formatting = new Formatting()
                      {
                        FontColor = Color.DarkMagenta,
                        Bold = true
                      }
               },
                 new ListLevelConfig()

                   {
                      NumberingFormat = NumberingFormat.lowerRoman,
                      NumberingLevelText = "%2.",
                      Justification = Justification.left,
                      Formatting = new Formatting()
                      {
                        FontColor = Color.Blue,
                        Bold = true
                      }
                   }
              ,

              new ListLevelConfig()
              {
                NumberingFormat = NumberingFormat.decimalNormal,
                NumberingLevelText = "%3.",
                Justification = Justification.left
              }
            ,
              new ListLevelConfig()
              {
                NumberingFormat = NumberingFormat.bullet,
                NumberingLevelText = "*",
                Justification = Justification.left
              }
            }

        };

        // Create a numbered list with the specified list options.
        var customNumberedList = document.AddList( customNumberedListOptions );
        // Add first item in the list.
        customNumberedList.AddListItem( "Cars" );
        // Add Sub-items(level 1) to the preceding ListItem.
        customNumberedList.AddListItem( "Cadillac", 1 );
        customNumberedList.AddListItem( "Mercedes-Benz", 1 );
        // Add Sub-items(level 2) to the preceding ListItem.
        customNumberedList.AddListItem( "SUVs", 2 );
        // Add Sub-items(level 3) to the preceding ListItem.
        customNumberedList.AddListItem( "GLA SUV", 3 );
        // Add Sub-items(level 4) to the preceding ListItem.
        customNumberedList.AddListItem( "GLA 250 4MATIC SUV", 4 );
        customNumberedList.AddListItem( "AMG GLA 35 4MATIC SUV", 4 );
        // Add Sub-items(level 5) to the preceding ListItem.
        customNumberedList.AddListItem( "AMG-enhanced 2.0L inline-4 turbo Engine", 5 );
        customNumberedList.AddListItem( "AMG SPEEDSHIFT DCT 8-speed dual-clutch Automatic", 5 );
        customNumberedList.AddListItem( "AMG GLA 45 4MATIC SUV", 4 );
        customNumberedList.AddListItem( "GLB SUV", 3 );
        customNumberedList.AddListItem( "GLC SUV", 3 );
        customNumberedList.AddListItem( "GLE SUV Coupe", 3 );
        customNumberedList.AddListItem( "Sedans & Wagons", 2 );
        // Add Sub-items(level 3) to the preceding ListItem.
        customNumberedList.AddListItem( "A-Class Hatch", 3 );
        customNumberedList.AddListItem( "A-Class Sedan", 3 );
        customNumberedList.AddListItem( "C-Class Sedan", 3 );
        customNumberedList.AddListItem( "Coupes", 2 );
        customNumberedList.AddListItem( "Convertibles & Roadsters", 2 );
        // Add more items at level 1.
        customNumberedList.AddListItem( "Maybach", 1 );
        customNumberedList.AddListItem( "Toyota", 1 );
        customNumberedList.AddListItem( "BMW", 1 );
        // Add an item (level 0)
        customNumberedList.AddListItem( "Boats" );
        customNumberedList.AddListItem( "Dinghies", 1 );
        customNumberedList.AddListItem( "Fish-and-ski Boat", 1 );
        customNumberedList.AddListItem( "Sportfishing Yachts", 1 );
        customNumberedList.AddListItem( "Trawlers", 1 );
        // Add an item (level 0)
        customNumberedList.AddListItem( "Planes" );
        // Add Sub-items(level 1) to the preceding ListItem.
        customNumberedList.AddListItem( "Aerospatiale", 1 );
        customNumberedList.AddListItem( "Boeing", 1 );
        customNumberedList.AddListItem( "Bombardier", 1 );
        customNumberedList.AddListItem( "Eclipse Aerospace", 1 );
        customNumberedList.AddListItem( "Embraer", 1 );

        // Insert the list into the document.
        document.InsertParagraph( "This is a custom Numbered List:\n" );
        document.InsertList( customNumberedList );

        document.Save();
        Console.WriteLine( "\tCreated: AddCustomNumberedList.docx\n" );
      }
#else
        // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Create a custom bulleted lists with different listItem's levels, items and numbering formatting.
    /// </summary>
    public static void AddCustomBulletedList()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tAddCustomBulletedList()" );

      // Create a document.
      using( var document = DocX.Create( ListSample.ListSampleOutputDirectory + @"AddCustomBulletedList.docx" ) )
      {
        // Set the document default font
        document.SetDefaultFont( new Document.NET.Font( "Times New Roman" ), 10d );

        // Add a title
        document.InsertParagraph( "Adding a custom bulleted list into a document" ).FontSize( 12d ).SpacingAfter( 25d ).Alignment = Alignment.center;

        // Add a custom bulleted list.
        // Define list levels symbols
        var firstLevelSymbol = new ListSymbol( "Wingdings", 216 );   // ⮚
        var secondLevelSymbol = new ListSymbol( "Wingdings", 175 );  // ✵	
        var thirdLevelSymbol = new ListSymbol( "Wingdings", 118 );   // ❖
        var forthLevelSymbol = new ListSymbol( "Wingdings", 240 );   // ⇨
        var fithLevelSymbol = new ListSymbol( "Wingdings", 70 );     // ☞
        var sixthLevelSymbol = new ListSymbol( "Wingdings", 182 );   // ✰

        // Create the list levels configuration
        var customBulletedListOptions = new ListOptions()
        {
          ListType = ListItemType.Bulleted,

          LevelsConfigs = new ObservableCollection<ListLevelConfig>()
            {

                new ListLevelConfig()
                   {
                      NumberingFormat = NumberingFormat.bullet,
                      NumberingLevelText = firstLevelSymbol.UnicodeToString(),
                      Justification = Justification.left,
                      Formatting = new Formatting()
                      {
                        FontFamily = new Document.NET.Font(firstLevelSymbol.FontName),
                        Size = 9d,
                        FontColor = Color.Red,
                        Bold = true
                      }
                   }
              ,

                new ListLevelConfig()
                   {
                      NumberingFormat = NumberingFormat.bullet,
                      NumberingLevelText = secondLevelSymbol.UnicodeToString(),
                      Justification = Justification.left,
                      Formatting = new Formatting()
                      {
                        FontFamily = new Document.NET.Font(secondLevelSymbol.FontName),
                        Bold = true
                      }
                   }
              ,

            new ListLevelConfig()
            {
              NumberingFormat = NumberingFormat.bullet,
              NumberingLevelText = thirdLevelSymbol.UnicodeToString(),
              Justification = Justification.left,
              Formatting = new Formatting()
              {
                FontFamily = new Document.NET.Font( secondLevelSymbol.FontName ),
                FontColor = Color.DarkSlateBlue,
                Bold = true
              }
            }
              ,

            new ListLevelConfig()
            {
              NumberingFormat = NumberingFormat.bullet,
              NumberingLevelText = forthLevelSymbol.UnicodeToString(),
              Justification = Justification.left,
              Formatting = new Formatting()
              {
                FontFamily = new Document.NET.Font( forthLevelSymbol.FontName ),
                FontColor = Color.Olive,
                Bold = true
              }
            }
              ,


              new ListLevelConfig()
              {
                NumberingFormat = NumberingFormat.bullet,
                NumberingLevelText = fithLevelSymbol.UnicodeToString(),
                Justification = Justification.left,
                Formatting = new Formatting()
                {
                  FontFamily = new Document.NET.Font( fithLevelSymbol.FontName ),
                  FontColor = Color.Chocolate,
                  Bold = true
                }
              }
           ,


              new ListLevelConfig()
              {
                NumberingFormat = NumberingFormat.bullet,
                NumberingLevelText = sixthLevelSymbol.UnicodeToString(),
                Justification = Justification.left,
                Formatting = new Formatting()
                {
                  FontFamily = new Document.NET.Font( sixthLevelSymbol.FontName ),
                  Bold = true
                }
              }
            }

        };

        // Create a bulleted list with the specified list options.
        var customBulletedList = document.AddList( customBulletedListOptions );
        // Add a first item with its sub-levels items to the list
        customBulletedList.AddListItem( "Xceed UI Components", 0, new Formatting() { Bold = true, UnderlineStyle = UnderlineStyle.singleLine } );
        // Add Sub-items(level 1) to the preceding ListItem.
        customBulletedList.AddListItem( "JavaScript", 1 );
        // Add Sub-items(level 2) to the preceding ListItem.
        customBulletedList.AddListItem( "Xceed DataGrid for JavaScript", 2, new Formatting() { FontColor = Color.DarkOrange, Bold = true } );
        // Add Sub-items(level 3) to the preceding ListItem.
        customBulletedList.AddListItem( "Already Feature Rich", 3 );
        customBulletedList.AddListItem( "Simple Column Management", 3 );
        // Add Sub-items(level 4) to the preceding ListItem.
        customBulletedList.AddListItem( "Viewers", 4 );
        // Add another Sub-items(level 3).
        customBulletedList.AddListItem( "Support for Frameworks and Libraries", 3 );
        // Add Sub-items(level 4) to the preceding ListItem.
        customBulletedList.AddListItem( "Angular", 4 );
        customBulletedList.AddListItem( "React", 4 );
        customBulletedList.AddListItem( "Vue.js (coming soon)", 4 );
        // Add some other Sub-items(level 3).
        customBulletedList.AddListItem( "Modern Development Approach", 3 );
        customBulletedList.AddListItem( "Invest in Piece of Mind", 3 );

        // Add more items and sub-items to the list
        customBulletedList.AddListItem( "Windows Presentation Foundation", 1 );
        customBulletedList.AddListItem( "Xceed DataGrid for WPF", 2, new Formatting() { FontColor = Color.DarkOrange, Bold = true } );
        customBulletedList.AddListItem( "Features", 3 );
        customBulletedList.AddListItem( "Editing, Printing, Exporting", 4 );
        customBulletedList.AddListItem( "Rich in-place Editing", 5 );
        customBulletedList.AddListItem( "Themed, and \"Theme-able\" Editor Controls", 5 );
        customBulletedList.AddListItem( "Provides all the Mouse and Keyboard Interactivity", 5 );
        customBulletedList.AddListItem( "Rich Printing for Easy Report Creation", 5 );
        customBulletedList.AddListItem( "Flexible Filtering and Powerful Sorting Capabilities", 5 );

        customBulletedList.AddListItem( "Grouping, Master-Detail", 4 );
        customBulletedList.AddListItem( "Supports Multi-level Grouping with all the Related Features.", 5 );
        customBulletedList.AddListItem( "Supports Master/Detail Hierarchy", 5 );

        customBulletedList.AddListItem( "Themes", 4 );
        customBulletedList.AddListItem( "MVVM support", 4 );


        customBulletedList.AddListItem( "Xceed Toolkit for WPF", 2, new Formatting() { FontColor = Color.DarkOrange, Bold = true } );
        customBulletedList.AddListItem( "Xceed Pro Theme for WPF", 2, new Formatting() { FontColor = Color.DarkOrange, Bold = true } );

        customBulletedList.AddListItem( "WinForms", 1 );


        // Add a second item with its sub-levels items to the list
        customBulletedList.AddListItem( "Xceed Data Manipulation Components", 0, new Formatting() { Bold = true, UnderlineStyle = UnderlineStyle.singleLine } );
        // Add Sub-items(level 1) to the preceding ListItem.
        customBulletedList.AddListItem( ".NET/Core", 1 );
        // Add Sub-items(level 2) to the preceding ListItem.
        customBulletedList.AddListItem( "Xceed Words for .NET", 2, new Formatting() { FontColor = Color.DarkOrange, Bold = true } );
        // Add Sub-items(level 3) to the preceding ListItem.
        customBulletedList.AddListItem( "Modify Existing .docx Documents.", 3 );
        customBulletedList.AddListItem( "Create new Microsoft Word .docx or Pdf Documents.", 3 );
        customBulletedList.AddListItem( "Complete Control Over the Document’s Properties", 3 );
        customBulletedList.AddListItem( "Easily Set or Modify Any Text’s Formatting", 3 );
        customBulletedList.AddListItem( "Search and Replace Text", 3 );
        customBulletedList.AddListItem( "Add or Remove Password Protection", 3 );
        customBulletedList.AddListItem( "Join Documents, Copy Documents, or Apply Templates", 3 );
        customBulletedList.AddListItem( "Supports Modifying Many Word Files in Parallel for Greater Speed", 3 );
        customBulletedList.AddListItem( "Add or Remove Password Protection", 3 );
        customBulletedList.AddListItem( "Add or Remove all Commonly Used Element Types", 3 );
        // Add Sub-items(level 4) to the preceding ListItem.
        customBulletedList.AddListItem( "Paragraphs", 4 );
        customBulletedList.AddListItem( "Bulleted or Numbered Lists", 4 );
        customBulletedList.AddListItem( "Images", 4 );
        customBulletedList.AddListItem( "Tables", 4 );
        customBulletedList.AddListItem( "Charts", 4 );
        customBulletedList.AddListItem( "Headers and Footers, Sections, Bookmarks, and More", 4 );

        // Add some other Sub-items(level 1).
        customBulletedList.AddListItem( "Xamarin", 1 );
        customBulletedList.AddListItem( "ActiveX", 1 );

        // Insert the list into the document.
        document.InsertParagraph( "This is a custom Bulleted List:\n" );
        document.InsertList( customBulletedList );

        document.Save();
        Console.WriteLine( "\tCreated: AddCustomBulletedList.docx\n" );
      }
#else
        // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Create a chapter numbered list with different listItem's levels, items and numbering formatting.
    /// </summary>
    public static void AddChapterList()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tAddChapterList()" );

      // Create a document.
      using( var document = DocX.Create( ListSample.ListSampleOutputDirectory + @"AddChapterList.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Adding a chapter-style list into a document" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add a chapter numbered list.
        var chapterListOptions = new ListOptions()
        {
          ListType = ListItemType.Numbered,

          LevelsConfigs = new ObservableCollection<ListLevelConfig>()
            {
              new ListLevelConfig()
                   {
                      NumberingFormat = NumberingFormat.upperRoman,
                      NumberingLevelText = "%1.",
                      Justification = Justification.left,
                      Indentation = new Indentation()
                      {
                        Hanging = 0f,
                        Left = 12f,
                        FirstLine = 0f
                      }
                   }
              ,
              new ListLevelConfig()
                   {
                      NumberingFormat = NumberingFormat.upperLetter,
                      NumberingLevelText = "%1. %2.",
                      Justification = Justification.left,
                      Indentation = new Indentation()
                      {
                        Hanging = 0f,
                        Left = 32f,
                        FirstLine = 0f
                      }
                   }
              ,
              new ListLevelConfig()
                   {
                      NumberingFormat = NumberingFormat.decimalNormal,
                      NumberingLevelText = "%1. %2. %3.",
                      Justification = Justification.left,
                      Indentation = new Indentation()
                      {
                        Hanging = 0f,
                        Left = 64f,
                        FirstLine = 0f
                      }
                   }
             ,
              new ListLevelConfig()
                   {
                      NumberingFormat = NumberingFormat.decimalNormal,
                      NumberingLevelText = "%1. %2. %3. %4.",
                      Justification = Justification.left,
                      Indentation = new Indentation()
                      {
                        Hanging = 0f,
                        Left = 72f,
                        FirstLine = 0f
                      }
                   }
              ,
               new ListLevelConfig()
                   {
                      NumberingFormat = NumberingFormat.decimalNormal,
                      NumberingLevelText = "%1. %2. %3. %4. %5.",
                      Justification = Justification.left,
                      Indentation = new Indentation()
                      {
                        Hanging = 0f,
                        Left = 88f,
                        FirstLine = 0f
                      }
                   }
              , new ListLevelConfig()
                   {
                      NumberingFormat = NumberingFormat.decimalNormal,
                      NumberingLevelText = "%1. %2. %3. %4. %5. %6.",
                      Justification = Justification.left,
                      Indentation = new Indentation()
                      {
                        Hanging = 0f,
                        Left = 104f,
                        FirstLine = 0f
                      }
                   }
              }
        };

        // Add a chapter numbered with the specified list options.
        var chapterList = document.AddList( chapterListOptions );
        // Add the first ListItem.
        chapterList.AddListItem( "Cars", 0 );
        // Add Sub-items(level 1) to the preceding ListItem.
        chapterList.AddListItem( "Cadillac", 1 );
        chapterList.AddListItem( "Mercedes-Benz", 1 );
        // Add Sub-items(level 2) to the preceding ListItem.
        chapterList.AddListItem( "SUVs", 2 );
        // Add Sub-items(level 3) to the preceding ListItem.
        chapterList.AddListItem( "GLA SUV", 3 );
        // Add Sub-items(level 4) to the preceding ListItem.
        chapterList.AddListItem( "GLA 250 4MATIC SUV", 4 );
        chapterList.AddListItem( "AMG GLA 35 4MATIC SUV", 4 );
        // Add Sub-items(level 5) to the preceding ListItem.
        chapterList.AddListItem( "AMG-enhanced 2.0L inline-4 turbo Engine", 5 );
        chapterList.AddListItem( "AMG SPEEDSHIFT DCT 8-speed dual-clutch Automatic", 5 );
        chapterList.AddListItem( "AMG GLA 45 4MATIC SUV", 4 );
        chapterList.AddListItem( "GLB SUV", 3 );
        chapterList.AddListItem( "GLC SUV", 3 );
        chapterList.AddListItem( "GLE SUV Coupe", 3 );
        chapterList.AddListItem( "Sedans & Wagons", 2 );
        // Add Sub-items(level 3) to the preceding ListItem.
        chapterList.AddListItem( "A-Class Hatch", 3 );
        chapterList.AddListItem( "A-Class Sedan", 3 );
        chapterList.AddListItem( "C-Class Sedan", 3 );
        chapterList.AddListItem( "Coupes", 2 );
        chapterList.AddListItem( "Convertibles & Roadsters", 2 );
        // Add more items at level 1.
        chapterList.AddListItem( "Maybach", 1 );
        chapterList.AddListItem( "Toyota", 1 );
        chapterList.AddListItem( "BMW", 1 );
        // Add an item (level 0)
        chapterList.AddListItem( "Boats" );
        chapterList.AddListItem( "Dinghies", 1 );
        chapterList.AddListItem( "Fish-and-ski Boat", 1 );
        chapterList.AddListItem( "Sportfishing Yachts", 1 );
        chapterList.AddListItem( "Trawlers", 1 );
        // Add an item (level 0)
        chapterList.AddListItem( "Planes" );
        // Add Sub-items(level 1) to the preceding ListItem.
        chapterList.AddListItem( "Aerospatiale", 1 );
        chapterList.AddListItem( "Boeing", 1 );
        chapterList.AddListItem( "Bombardier", 1 );
        chapterList.AddListItem( "Eclipse Aerospace", 1 );
        chapterList.AddListItem( "Embraer", 1 );

        // Insert the list into the document.
        document.InsertParagraph( "This is a Chapter Numbered List:\n" );
        document.InsertList( chapterList );

        document.Save();
        Console.WriteLine( "\tCreated: AddChapterList.docx\n" );
      }
#else
        // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }


    /// <summary>
    /// Changing document list numbering type and levels configurations.
    /// </summary>
    public static void ModifyList()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tModifyList()" );

      // Load a document with numbering lists.
      // The first list is a decimal numbering list.
      // The second list is a bulleted numbered list.
      using( var document = DocX.Load( ListSample.ListSampleResourceDirectory + @"Lists.docx" ) )
      {
        // Get the docment first list.
        var firstList = document.Lists[ 0 ];

        // Update the list numbering type and levels configurations.
        if( firstList.ListOptions.ListType == ListItemType.Numbered )
        {
          firstList.ListOptions = new ListOptions()
          {
            ListType = ListItemType.Bulleted,
            LevelsConfigs = new ObservableCollection<ListLevelConfig>()
              {
                   new ListLevelConfig()
                       {
                          NumberingFormat = NumberingFormat.bullet,
                          NumberingLevelText = ">",
                          Justification = Justification.left
                       }
              }
          };
        }

        // Get the document second list.
        var secondList = document.Lists[ 1 ];

        // Update the list numbering type and levels configurations.
        if( secondList.ListOptions.ListType == ListItemType.Bulleted )
        {
          secondList.ListOptions.ListType = ListItemType.Numbered;
          var levelsConfigs = secondList.ListOptions.LevelsConfigs;

          if( levelsConfigs.Count > 0 )
          {
            secondList.ListOptions.LevelsConfigs[ 0 ] = new ListLevelConfig()
            {
              NumberingFormat = NumberingFormat.upperRoman,
              NumberingLevelText = "%1)"
            };
          }
        }

        // Update the document title and lists name.
        var title = document.Paragraphs.FirstOrDefault( p => p.Text.Contains( "numbering" ) );
        var firstListTitle = document.Paragraphs.FirstOrDefault( p => p.Text.Contains( "Numbered" ) );
        var secondListTitle = document.Paragraphs.FirstOrDefault( p => p.Text.Contains( "Bulleted" ) );

        title?.RemoveText( 0 );
        title?.Append( "Updated Lists" );

        firstListTitle?.RemoveText( 0 );
        firstListTitle?.Append( "This is a Bulleted list:" );

        secondListTitle?.RemoveText( 0 );
        secondListTitle?.Append( "This is a Numbered list:" );

        document.SaveAs( ListSample.ListSampleOutputDirectory + @"UpdatedLists.docx" );
        Console.WriteLine( "\tCreated: UpdatedLists.docx\n" );
      }
#else
        // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Clone a list and modify some ListItems.
    /// </summary>
    public static void CloneLists()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tCloneList()" );

      // Load a document.
      using( var document = DocX.Load( ListSample.ListSampleResourceDirectory + @"TemplateLists.docx" ) )
      {
        // Get the first list : the numbered list, and clone it.
        var newNumberedList = document.AddList( document.Lists[ 0 ] );
        // Add a listItem in the new list.
        newNumberedList.AddListItem( "Orange" );
        // Add a formatted listItem in the new list.
        newNumberedList.AddListItem( "Strawberry", 0, new Formatting() { Bold = true } );
        // Insert the new list in the document.
        document.InsertParagraph( "This is the new numbered list: " ).SpacingBefore( 20d ).InsertListAfterSelf( newNumberedList );

        // Get the second list : the bulleted list, and clone it.
        var newBulletedList = document.AddList( document.Lists[ 1 ] );
        // Add a listItem in the new list.
        newBulletedList.AddListItem( "Fifth" );
        // Add a formatted listItem in the new list.
        newBulletedList.AddListItem( "Sixth", 0, new Formatting() { UnderlineStyle = UnderlineStyle.singleLine } );
        // Insert the new list in the document.
        document.InsertParagraph( "This is the new bulleted list: " ).SpacingBefore( 30d ).InsertListAfterSelf( newBulletedList );

        document.SaveAs( ListSample.ListSampleOutputDirectory + @"CloneLists.docx" );
        Console.WriteLine( "\tCreated: CloneLists.docx\n" );
      }
#else
            // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    #endregion
  }


#if !OPEN_SOURCE
  public class ListSymbol
  {
    public string FontName;
    public int Code;

    public ListSymbol( string fontName, int code )
    {
      this.FontName = fontName;
      this.Code = code;
    }

    public string UnicodeToString()
    {
      return char.ConvertFromUtf32( this.Code );
    }
  }
#endif
}
