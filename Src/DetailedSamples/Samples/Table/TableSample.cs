﻿/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Table Sample Application
Copyright (c) 2009-2024 - Xceed Software Inc.
 
This application demonstrates how to create and format a table when using the API 
from the Xceed Words for .NET.
 
This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.Drawing;
using System.IO;
using System.Linq;
using Xceed.Document.NET;

namespace Xceed.Words.NET.Examples
{
  public class TableSample
  {
    #region Private Members

    private static Random rand = new Random();

    private const string TableSampleResourcesDirectory = Program.SampleDirectory + @"Table\Resources\";
    private const string TableSampleOutputDirectory = Program.SampleDirectory + @"Table\Output\";

    #endregion

    #region Constructors

    static TableSample()
    {
      if( !Directory.Exists( TableSample.TableSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( TableSample.TableSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Create a table, insert rows, image and replace text.
    /// </summary>
    public static void InsertRowAndImageTable()
    {
      Console.WriteLine( "\tInsertRowAndImageTable()" );

      // Create a document.
      using( var document = DocX.Create( TableSample.TableSampleOutputDirectory + @"InsertRowAndImageTable.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Inserting table" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add a Table into the document and sets its values.
        var t = document.AddTable( 5, 2 );
        t.Design = TableDesign.ColorfulListAccent1;
        t.Alignment = Alignment.center;
        t.Rows[ 0 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "Mike" );
        t.Rows[ 0 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "65" );
        t.Rows[ 1 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "Kevin" );
        t.Rows[ 1 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "62" );
        t.Rows[ 2 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "Carl" );
        t.Rows[ 2 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "60" );
        t.Rows[ 3 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "Michael" );
        t.Rows[ 3 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "59" );
        t.Rows[ 4 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "Shawn" );
        t.Rows[ 4 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "57" );

        // Set the width of the 2 columns.
        t.SetWidths( new float[] { 115f, 115f } );

        // Add a row at the end of the table and sets its values.
        var r = t.InsertRow();
        r.Cells[ 0 ].Paragraphs[ 0 ].Append( "Mario" );
        r.Cells[ 1 ].Paragraphs[ 0 ].Append( "54" );

        // Add a row at the end of the table which is a copy of another row, and sets its values.
        var newPlayer = t.InsertRow( t.Rows[ 2 ] );
        newPlayer.ReplaceText( new StringReplaceTextOptions() { SearchValue = "Carl", NewValue = "Max" } );
        newPlayer.ReplaceText( new StringReplaceTextOptions() { SearchValue = "60", NewValue = "50" } );

        // Add an image into the document.    
        var image = document.AddImage( TableSample.TableSampleResourcesDirectory + @"logo_xceed.png" );
        // Create a picture from image.
        var picture = image.CreatePicture( 25f, 100f );

        // Calculate totals points from second column in table.
        var totalPts = 0;
        foreach( var row in t.Rows )
        {
          totalPts += int.Parse( row.Cells[ 1 ].Paragraphs[ 0 ].Text );
        }

        // Add a row at the end of the table and sets its values.
        var totalRow = t.InsertRow();
        totalRow.Cells[ 0 ].Paragraphs[ 0 ].Append( "Total for " ).AppendPicture( picture );
        totalRow.Cells[ 1 ].Paragraphs[ 0 ].Append( totalPts.ToString() );
        totalRow.Cells[ 1 ].VerticalAlignment = VerticalAlignment.Center;

        // Insert a new Paragraph into the document.
        var p = document.InsertParagraph( "Xceed Top Players Points:" );
        p.SpacingAfter( 40d );

        // Insert the Table after the Paragraph.
        p.InsertTableAfterSelf( t );

        document.Save();
        Console.WriteLine( "\tCreated: InsertRowAndImageTable.docx\n" );
      }
    }

    /// <summary>
    /// Create a table, insert rows and make the table wraps around text.
    /// </summary>
    public static void AddTableWithTextWrapping()
    {
#if !OPEN_SOURCE

      Console.WriteLine( "\tAddTableWithTextWrapping()" );

      // Create a document.
      using( var document = DocX.Create( TableSample.TableSampleOutputDirectory + @"AddTableWithTextWrapping.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Add Table with Text Wrapping" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add a Table into the document and sets its values.
        var t = document.AddTable( 3, 2 );
        t.Rows[ 0 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "Mike" );
        t.Rows[ 0 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "65" );
        t.Rows[ 1 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "Kevin" );
        t.Rows[ 1 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "62" );
        t.Rows[ 2 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "Carl" );
        t.Rows[ 2 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "60" );

        // Set columns width.
        t.SetWidths( new float[] { 115f, 115f } );

        // Set the table wrapping as WrapAround.
        t.WrapStyle = TableWrapStyle.WrapAround;
        // Set horizontal alignment with a right Alignement from margin.
        t.HorizontalAlignment = WrappingHorizontalAlignment.RightRelativeToMargin;
        // Set vertical alignement with an offset from top of page.
        t.VerticalOffsetAlignmentFrom = WrappingVerticalOffsetAlignmentFrom.Page;
        t.VerticalOffset = 175d;
        // Set a buffer on left and right of table where no text will be drawn.
        t.DistanceFromTextLeft = 5d;
        t.DistanceFromTextRight = 5d;

        var p = document.InsertParagraph( "With its easy to use API, Xceed Words for .NET lets your application create new Microsoft Word .docx or PDF documents, or modify existing .docx documents. It gives you complete control over all content in a Word document, and lets you add or remove all commonly used element types, such as paragraphs, bulleted or numbered lists, images, tables, charts, headers and footers, sections, bookmarks, and more. Create PDF documents using the same API for creating Word documents." );
        p.Alignment = Alignment.both;
        p.InsertTableAfterSelf( t );
        p.SpacingAfter( 30 );

        document.Save();
        Console.WriteLine( "\tCreated: AddTableWithTextWrapping.docx\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Clone a table and modify data in it.
    /// </summary>
    public static void CloneTable()
    {
#if !OPEN_SOURCE

      Console.WriteLine( "\tCloneTable()" );

      // Load a document.
      using( var document = DocX.Load( TableSample.TableSampleResourcesDirectory + @"TemplateTable.docx" ) )
      {
        // Get the first table and clone it.
        var newTable = document.AddTable( document.Tables[ 0 ] );
        // Modify the data of the second cell in the third row with formatted data.
        newTable.Rows[ 2 ].Cells[ 1 ].ReplaceText( new StringReplaceTextOptions() { SearchValue = "$35.99", NewValue = "$45.99", NewFormatting = new Formatting() { Bold = true } } );
        // Add a new row, based on the last row, in the new table.
        var newRow = newTable.InsertRow( newTable.Rows.Last() );
        newRow.ReplaceText( new StringReplaceTextOptions() { SearchValue = "Helmet", NewValue = "Glove" } );
        newRow.ReplaceText( new StringReplaceTextOptions() { SearchValue = "$29.99", NewValue = "$75.99" } );
        // Insert the new table in the document.
        document.InsertParagraph( "This is the new table: " ).InsertTableAfterSelf( newTable );

        document.SaveAs( TableSample.TableSampleOutputDirectory + @"CloneTable.docx" );
        Console.WriteLine( "\tCreated: CloneTable.docx\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    /// <summary>
    /// Create a table and set the text direction of each cell.
    /// </summary>
    public static void TextDirectionTable()
    {
      Console.WriteLine( "\tTextDirectionTable()" );

      // Create a document.
      using( var document = DocX.Create( TableSample.TableSampleOutputDirectory + @"TextDirectionTable.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Text Direction of Table's cells" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a table.
        var table = document.AddTable( 2, 3 );
        table.Design = TableDesign.ColorfulList;

        // Set the table's values.
        table.Rows[ 0 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "First" );
        table.Rows[ 0 ].Cells[ 0 ].TextDirection = TextDirection.btLr;
        table.Rows[ 0 ].Cells[ 0 ].Paragraphs[ 0 ].Spacing( 5d );
        table.Rows[ 0 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "Second" );
        table.Rows[ 0 ].Cells[ 1 ].TextDirection = TextDirection.right;
        table.Rows[ 0 ].Cells[ 2 ].Paragraphs[ 0 ].Append( "Third" );
        table.Rows[ 0 ].Cells[ 2 ].Paragraphs[ 0 ].Spacing( 5d );
        table.Rows[ 0 ].Cells[ 2 ].TextDirection = TextDirection.btLr;
        table.Rows[ 1 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "Fourth" );
        table.Rows[ 1 ].Cells[ 0 ].TextDirection = TextDirection.btLr;
        table.Rows[ 1 ].Cells[ 0 ].Paragraphs[ 0 ].Spacing( 5d );
        table.Rows[ 1 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "Fifth" );
        table.Rows[ 1 ].Cells[ 2 ].Paragraphs[ 0 ].Append( "Sixth" ).Color( Color.White );
        table.Rows[ 1 ].Cells[ 2 ].TextDirection = TextDirection.btLr;
        // Last cell have a green background
        table.Rows[ 1 ].Cells[ 2 ].FillColor = Color.Green;
        table.Rows[ 1 ].Cells[ 2 ].Paragraphs[ 0 ].Spacing( 5d );

        // Set the table's column width.
        table.SetWidths( new float[] { 200, 300, 100 } );

        // Add the table into the document.
        document.InsertTable( table );

        document.Save();
        Console.WriteLine( "\tCreated: TextDirectionTable.docx\n" );
      }
    }

    /// <summary>
    /// Load a document, gets its table and replace the default row with updated copies of it.
    /// </summary>
    public static void CreateRowsFromTemplate()
    {
      Console.WriteLine( "\tCreateRowsFromTemplate()" );

      // Load a document
      using( var document = DocX.Load( TableSample.TableSampleResourcesDirectory + @"DocumentWithTemplateTable.docx" ) )
      {
        // get the table with caption "GROCERY_LIST" from the document.
        var groceryListTable = document.Tables.FirstOrDefault( t => t.TableCaption == "GROCERY_LIST" );
        if( groceryListTable == null )
        {
          Console.WriteLine( "\tError, couldn't find table with caption GROCERY_LIST in current document." );
        }
        else
        {
          if( groceryListTable.RowCount > 1 )
          {
            // Get the row pattern of the second row.
            var rowPattern = groceryListTable.Rows[ 1 ];

            // Add items (rows) to the grocery list.
            TableSample.AddItemToTable( groceryListTable, rowPattern, "Banana" );
            TableSample.AddItemToTable( groceryListTable, rowPattern, "Strawberry" );
            TableSample.AddItemToTable( groceryListTable, rowPattern, "Chicken" );
            TableSample.AddItemToTable( groceryListTable, rowPattern, "Bread" );
            TableSample.AddItemToTable( groceryListTable, rowPattern, "Eggs" );
            TableSample.AddItemToTable( groceryListTable, rowPattern, "Salad" );

            // Remove the pattern row.
            rowPattern.Remove();
          }
        }

        document.SaveAs( TableSample.TableSampleOutputDirectory + @"CreateTableFromTemplate.docx" );
        Console.WriteLine( "\tCreated: CreateTableFromTemplate.docx\n" );
      }
    }

    /// <summary>
    /// Add a Table in a document where its columns will have a specific width. In addition,
    /// the left margin of the row cells will be removed for all rows except the first.
    /// Finally, a blank border will be set for the table's top and bottom borders.
    /// </summary>
    public static void ColumnsWidth()
    {
      Console.WriteLine( "\tColumnsWidth()" );

      // Create a document
      using( var document = DocX.Create( TableSample.TableSampleOutputDirectory + @"ColumnsWidth.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Columns width" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert a title paragraph.
        var p = document.InsertParagraph( "In the following table, the cell's left margin has been removed for rows 2-6 as well as the top/bottom table's borders." ).Bold();
        p.Alignment = Alignment.center;
        p.SpacingAfter( 40d );

        // Add a table in a document of 1 row and 3 columns.
        var columnWidths = new float[] { 100f, 300f, 200f };
        var t = document.InsertTable( 1, columnWidths.Length );

        // Set the table's column width and background 
        t.SetWidths( columnWidths );
        t.AutoFit = AutoFit.Contents;

        var row = t.Rows.First();

        // Fill in the columns of the first row in the table.
        for( int i = 0; i < row.Cells.Count; ++i )
        {
          row.Cells[ i ].Paragraphs.First().Append( "Data " + i );
        }

        // Add rows in the table.
        for( int i = 0; i < 5; i++ )
        {
          var newRow = t.InsertRow();

          // Fill in the columns of the new rows.
          for( int j = 0; j < newRow.Cells.Count; ++j )
          {
            var newCell = newRow.Cells[ j ];
            newCell.Paragraphs.First().Append( "Data " + i );
            // Remove the left margin of the new cells.
            newCell.MarginLeft = 0;
          }
        }

        // Set a blank border for the table's top/bottom borders.
        var blankBorder = new Border( BorderStyle.Tcbs_none, 0, 0f, Color.White );
        t.SetBorder( TableBorderType.Bottom, blankBorder );
        t.SetBorder( TableBorderType.Top, blankBorder );

        document.Save();
        Console.WriteLine( "\tCreated: ColumnsWidth.docx\n" );
      }
    }

    /// <summary>
    /// Add a table and merged some cells. Individual cells can also be removed by shifting their right neighbors to the left.
    /// </summary>
    public static void MergeCells()
    {
      Console.WriteLine( "\tMergeCells()" );

      // Create a document.
      using( var document = DocX.Create( TableSample.TableSampleOutputDirectory + @"MergeCells.docx" ) )
      {
        // Add a title.
        document.InsertParagraph( "Merge and delete cells" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert a table.
        var t1 = document.InsertTable( 3, 2 );

        // Add 4 columns in the table.
        t1.InsertColumn();
        t1.InsertColumn();
        t1.InsertColumn( t1.ColumnCount - 1, true );
        t1.InsertColumn( t1.ColumnCount - 1, true );

        // Merged Cells 1 to 4 in first row of the table.
        t1.Rows[ 0 ].MergeCells( 1, 4 );

        // Merged the last 2 Cells in the second row of the table.
        var columnCount = t1.Rows[ 1 ].ColumnCount;
        t1.Rows[ 1 ].MergeCells( columnCount - 2, columnCount - 1 );

        // Add text in each cell of the table.
        foreach( var r in t1.Rows )
        {
          for( int i = 0; i < r.Cells.Count; ++i )
          {
            var c = r.Cells[ i ];
            c.Paragraphs[ 0 ].InsertText( "Column " + i );
            c.Paragraphs[ 0 ].Alignment = Alignment.center;
          }
        }

        // Delete the second cell from the third row and shift the cells on its right by 1 to the left.
        t1.DeleteAndShiftCellsLeft( 2, 1 );

        document.Save();
        Console.WriteLine( "\tCreated: MergeCells.docx\n" );
      }
    }

    /// <summary>
    /// Get/Set shading pattern to a table or cells.
    /// </summary>
    public static void ShadingPattern()
    {
      Console.WriteLine( "\tShadingPattern()" );

      // Create a document.
      using( var document = DocX.Create( TableSample.TableSampleOutputDirectory + @"ShadingPattern.docx" ) )
      {
        // Add a title.
        document.InsertParagraph( "Set shading pattern" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        document.InsertParagraph( "Set shading pattern for cells: " ).FontSize( 13d );
        
        // Insert a table.
        var t1 = document.InsertTable( 3, 4 );

        // Set shading pattern 1.
        t1.Rows[ 0 ].Cells[ 0 ].ShadingPattern = new ShadingPattern()
        {
          Fill = Color.Red,
          Style = PatternStyle.Percent30,
          StyleColor = Color.Yellow
        };

        // Set shading pattern 2.
        t1.Rows[ 0 ].Cells[ 1 ].ShadingPattern = new ShadingPattern()
        {
          Fill = Color.Blue,
          Style = PatternStyle.Percent50,
          StyleColor = Color.White
        };

        // Set shading pattern 3.
        t1.Rows[ 2 ].Cells[ 1 ].ShadingPattern = new ShadingPattern()
        {
          Fill = Color.Green,
          Style = PatternStyle.DkHorizonal,
          StyleColor = Color.Brown
        };

        // Set shading pattern to table cells by properties.
        t1.Rows[ 0 ].Cells[ 2 ].ShadingPattern.Fill = Color.Coral;
        t1.Rows[ 0 ].Cells[ 2 ].ShadingPattern.Style = PatternStyle.LtGrid;

        // Set shading pattern to table cells by properties.
        t1.Rows[ 2 ].Cells[ 2 ].ShadingPattern.Fill = Color.Brown;
        t1.Rows[ 2 ].Cells[ 2 ].ShadingPattern.Style = PatternStyle.LtGrid;
        t1.Rows[ 2 ].Cells[ 2 ].ShadingPattern.StyleColor = Color.DarkRed;

        // Set shading pattern to paragraphs
        t1.Rows[ 1 ].Cells[ 3 ].Paragraphs[ 0 ].Append( "Mike" );
        t1.Rows[ 1 ].Cells[ 3 ].Paragraphs[ 0 ].InsertParagraphAfterSelf( "Tom" );
        t1.Rows[ 1 ].Cells[ 3 ].Paragraphs[ 0 ].InsertParagraphAfterSelf( "John" );

        var shadingPattern_paragraph_1 = new ShadingPattern()
        {
          Fill = Color.Green,
          Style = PatternStyle.DkHorizonal,
          StyleColor = Color.Brown
        };

        var shadingPattern_paragraph_2 = new ShadingPattern()
        {
          Fill = Color.Blue,
          Style = PatternStyle.DkUpDiagonal,
          StyleColor = Color.White
        };

        t1.Rows[ 1 ].Cells[ 3 ].Paragraphs[ 0 ].ShadingPattern(shadingPattern_paragraph_1, ShadingType.Paragraph);
        t1.Rows[ 1 ].Cells[ 3 ].Paragraphs[ 2 ].ShadingPattern( shadingPattern_paragraph_2, ShadingType.Paragraph );

        document.InsertParagraph( "Set shading pattern for a table: " ).FontSize( 13d ).SpacingBefore( 40d );

        // Insert a table.
        var t2 = document.InsertTable( 3, 2 );

        // Set shading pattern to a table.
        t2.ShadingPattern = new ShadingPattern()
        {
          Fill = Color.Red,
          Style = PatternStyle.Percent30,
          StyleColor = Color.Yellow
        };

        document.Save();
        Console.WriteLine( "\tCreated: ShadingPattern.docx\n" );
      }
    }

    #endregion

    #region Private Methods

    private static void AddItemToTable( Table table, Row rowPattern, string productName )
    {
      // Gets a random unit price and quantity.
      var unitPrice = Math.Round( rand.NextDouble(), 2 );
      var unitQuantity = rand.Next( 1, 10 );

      // Insert a copy of the rowPattern at the last index in the table.
      var newItem = table.InsertRow( rowPattern, table.RowCount - 1 );

      // Replace the default values of the newly inserted row.
      newItem.ReplaceText( new StringReplaceTextOptions() { SearchValue = "%PRODUCT_NAME%", NewValue = productName } );
      newItem.ReplaceText( new StringReplaceTextOptions() { SearchValue = "%PRODUCT_UNITPRICE%", NewValue = "$ " + unitPrice.ToString( "N2" ) } );
      newItem.ReplaceText( new StringReplaceTextOptions() { SearchValue = "%PRODUCT_QUANTITY%", NewValue = unitQuantity.ToString() } );
      newItem.ReplaceText( new StringReplaceTextOptions() { SearchValue = "%PRODUCT_TOTALPRICE%", NewValue = "$ " + ( unitPrice * unitQuantity ).ToString( "N2" ) } );
    }

    #endregion
  }
}
