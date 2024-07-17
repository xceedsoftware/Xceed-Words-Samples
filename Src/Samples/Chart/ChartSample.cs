/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Chart Sample Application
Copyright (c) 2009-2024 - Xceed Software Inc.
 
This application demonstrates how to create a chart when using the API 
from the Xceed Words for .NET.
 
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
  public class ChartSample
  {
    #region Private Members
    private const string ImageSampleResourcesDirectory = Program.SampleDirectory + @"Image\Resources\";
    private const string ChartSampleOutputDirectory = Program.SampleDirectory + @"Chart\Output\";
    private const string ChartSampleResourceDirectory = Program.SampleDirectory + @"Chart\Resources\";

    #endregion

    #region Constructors

    static ChartSample()
    {
      if( !Directory.Exists( ChartSample.ChartSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( ChartSample.ChartSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Add a Bar chart to a document.
    /// </summary>
    public static void BarChart()
    {
      Console.WriteLine( "\tBarChart()" );

      // Creates a document
      using( var document = DocX.Create( ChartSample.ChartSampleOutputDirectory + @"BarChart.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Bar Chart" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a bar chart.
        var c = document.AddChart<BarChart>();
        c.AddLegend( ChartLegendPosition.Left, false );
        c.BarDirection = BarDirection.Bar;
        c.BarGrouping = BarGrouping.Standard;
        c.GapWidth = 200;

#if !OPEN_SOURCE
        // Position Category axis's labels on the lower side.
        c.ValueAxis.LabelPosition = LabelPosition.Low;
        // Add titles to axis
        c.CategoryAxis.Title = "Categories";
        c.ValueAxis.Title = "Expenses";
        // Set Tick marks and units
        c.ValueAxis.MajorTickMarksUnits = 50;
        c.ValueAxis.MinorTickMarksUnits = 10;
        c.ValueAxis.MinorTicksMarksType = TickMarksTypes.inside;
        c.ValueAxis.MajorTicksMarksType = TickMarksTypes.outside;

#endif

        // Create the data.
        var canada = ChartData.CreateCanadaExpenses();
        var usa = ChartData.CreateUSAExpenses();
        var brazil = ChartData.CreateBrazilExpenses();

        // Create and add series
        var s1 = new Series( "Brazil" );
        s1.Color = Color.GreenYellow;
        s1.Bind( brazil, "Category", "Expenses" );
        c.AddSeries( s1 );

        var s2 = new Series( "USA" );
        s2.Color = Color.LightBlue;
        s2.Bind( usa, "Category", "Expenses" );
        c.AddSeries( s2 );

        var s3 = new Series( "Canada" );
        s3.Color = Color.Gray;
        s3.Bind( canada, "Category", "Expenses" );
        c.AddSeries( s3 );

        // Insert the chart into the document.
        document.InsertParagraph( "Expenses(M$) for selected categories per country" ).FontSize( 15 ).SpacingAfter( 10d );
        document.InsertChart( c, 350f, 550f );

        document.Save();
        Console.WriteLine( "\tCreated: BarChart.docx\n" );
      }
    }

    /// <summary>
    /// Add a Line chart to a document.
    /// </summary>
    public static void LineChart()
    {
      Console.WriteLine( "\tLineChart()" );

      // Creates a document
      using( var document = DocX.Create( ChartSample.ChartSampleOutputDirectory + @"LineChart.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Line Chart" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a line chart.
        var c = document.AddChart<LineChart>();
#if !OPEN_SOURCE        
        // Add titles to Category axis.
        c.CategoryAxis.Title = "Categories";

        // Add formatting for Category Axis title.
        c.CategoryAxis.TitleFormat = new TitleFormatting()
        {
          UnderlineColor = Color.Blue,
          FontColor = Color.Red,
          Size = 16,
          StrikeThrough = AxisTitleStrikeThrough.noStrike,
          UnderlineStyle = AxisTitleUnderlineStyle.dotDotDashHeavy,
          Bold = true,
          FontFamily = new Xceed.Document.NET.Font( "Comic Sans MS" ),
          Highlight = Color.Yellow,
          CharacterSpacing = 4
        };

        c.ValueAxis.Title = "Expenses";
#endif
        c.AddLegend(ChartLegendPosition.Left, false);

        // Create the data.
        var canada = ChartData.CreateCanadaExpenses();
        var usa = ChartData.CreateUSAExpenses();
        var brazil = ChartData.CreateBrazilExpenses();

        // Create and add series
        var s1 = new Series( "Brazil" );
        s1.Color = Color.Yellow;
        s1.Bind( brazil, "Category", "Expenses" );
#if !OPEN_SOURCE
        s1.Width = 12;

        // Set the markers for the whole series "Brazil".
        var image = document.AddImage(ChartSample.ImageSampleResourcesDirectory + @"balloon.jpg");
        s1.Marker = new Marker()
        {
          Image = image,
          Size = 25,
          SymbolType = MarkerSymbolType.Circle,
          OutlineColor = Color.Black
        };

        // Set the marker of the 2nd DataPoint from the series "Brazil".
        s1.DataPoints[1].Marker = new Marker()
        {
          SolidFill = Color.AliceBlue,
          Size = 30,
          SymbolType = MarkerSymbolType.Diamond,
          OutlineColor = Color.Yellow
        };
        
#endif
        c.AddSeries( s1 );

        var s2 = new Series( "USA" );
        s2.Color = Color.Blue;
#if !OPEN_SOURCE
        s2.Width = 4;
#endif
        s2.Bind( usa, "Category", "Expenses" );

        c.AddSeries( s2 );

        var s3 = new Series( "Canada" );
        s3.Color = Color.Red;
        s3.Bind( canada, "Category", "Expenses" );
        c.AddSeries( s3 );

        // Insert chart into document
        document.InsertParagraph( "Expenses(M$) for selected categories per country" ).FontSize( 15 ).SpacingAfter( 10d );
        document.InsertChart( c );
        document.Save();

        Console.WriteLine( "\tCreated: LineChart.docx\n" );
      }
    }

    /// <summary>
    /// Add a Pie chart to a document.
    /// </summary>
    public static void PieChart()
    {
      Console.WriteLine( "\tPieChart()" );

      // Creates a document
      using( var document = DocX.Create( ChartSample.ChartSampleOutputDirectory + @"PieChart.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Pie Chart" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a pie chart.
        var c = document.AddChart<PieChart>();
        c.AddLegend( ChartLegendPosition.Left, false );

        // Create the data.
        var brazil = ChartData.CreateBrazilExpenses();

        // Create and add series
        var s1 = new Series( "Brazil" );
        s1.Bind( brazil, "Category", "Expenses" );
        c.AddSeries( s1 );

        // Insert chart into document
        document.InsertParagraph( "Expenses(M$) for selected categories in Brazil" ).FontSize( 15 ).SpacingAfter( 10d );
        document.InsertChart( c );

        document.Save();
        Console.WriteLine( "\tCreated: PieChart.docx\n" );
      }
    }

    /// <summary>
    /// Add a 3D bar chart to a document.
    /// </summary>
    /// 
    public static void Chart3D()
    {
      Console.WriteLine( "\tChart3D()" );

      // Creates a document
      using( var document = DocX.Create( ChartSample.ChartSampleOutputDirectory + @"3DChart.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "3D Chart" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a 3D Bar chart.
        var c = document.AddChart<BarChart>();
        c.View3D = true;

        // Create the data.
        var brazil = ChartData.CreateBrazilExpenses();

        // Create and add series
        var s1 = new Series( "Brazil" );
        s1.Color = Color.GreenYellow;
        s1.Bind( brazil, "Category", "Expenses" );
        c.AddSeries( s1 );

        // Insert chart into document
        document.InsertParagraph( "Expenses(M$) for selected categories in Brazil" ).FontSize( 15 ).SpacingAfter( 10d );
        document.InsertChart( c );

        document.Save();
        Console.WriteLine( "\tCreated: 3DChart.docx\n" );
      }
    }

    public static void ModifyChartData()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tModifyChartData()" );

      // Loads a document.
      using( var document = DocX.Load( ChartSample.ChartSampleResourceDirectory + @"Report.docx" ) )
      {
        foreach( var p in document.Paragraphs )
        {
          // Gets the paragraph's charts.
          var charts = p.Charts;
          if( charts.Count > 0 )
          {
            // Gets the first chart's first serie's values.
            var numbers = charts[ 0 ].Series[ 0 ].Values;
            // Modify the third value from 2 to 6.
            numbers[ 2 ] = "6";
            // Add a new value.
            numbers.Add( "3" );
            // Update the first chart's first serie's values with the new one.
            charts[ 0 ].Series[ 0 ].Values = numbers;

            // Gets the first chart's first serie's categories.
            var categories = charts[ 0 ].Series[ 0 ].Categories;
            // Modify the second category from Canada to Russia.
            categories[ 1 ] = "Russia";
            // Add a new category.
            categories.Add( "Italia" );
            // Update the first chart's first serie's categories with the new one.
            charts[ 0 ].Series[ 0 ].Categories = categories;

            // Modify first chart's first serie's color from Blue to Gold.
            charts[ 0 ].Series[ 0 ].Color = Color.Gold;

            // Add a new Series
            var s1 = new Series( "Airplanes" );
            s1.Color = Color.Red;
            s1.Categories = categories;
            s1.Values = new List<double>() { 2, 5, 2, 3, 4 };
            charts[ 0 ].AddSeries( s1 );

            // Change the Values axis title
            charts[ 0 ].ValueAxis.Title = "Transportation Vehicles";

            // Remove the legend.
            charts[ 0 ].RemoveLegend();

            // Make the chart a 3D chart.
            charts[0].View3D = true;

            // Save the changes in the first chart.
            charts[ 0 ].Save();
          }
        }

        document.SaveAs( ChartSample.ChartSampleOutputDirectory + @"ModifyChartData.docx" );
        Console.WriteLine( "\tCreated: ModifyChartData.docx\n" );
      }
#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
#endif
    }

    public static void AddChartWithTextWrapping()
    {
#if !OPEN_SOURCE
      Console.WriteLine( "\tAddChartWithTextWrapping()" );

      // Create a document.
      using( var document = DocX.Create( ChartSample.ChartSampleOutputDirectory + @"AddChartWithTextWrapping.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Add Chart with Text Wrapping" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add a Pie Chart and set its wrapping as Square.
        var c = document.AddChart<PieChart>();
        c.AddLegend( ChartLegendPosition.Left, false );

        // Create the data.
        var brazil = ChartData.CreateBrazilExpenses();

        // Create and add series
        var s1 = new Series( "Brazil" );
        s1.Bind( brazil, "Category", "Expenses" );
        c.AddSeries( s1 );
        c.WrapStyle = WrapStyle.WrapSquare;
        c.WrapTextPosition = WrapText.bothSides;
        // Set horizontal alignment with Alignement centered on the page.
        c.HorizontalAlignment = WrappingHorizontalAlignment.CenteredRelativeToPage;
        // Set vertical alignement with an offset from top of paragraph.
        c.VerticalOffsetAlignmentFrom = WrappingVerticalOffsetAlignmentFrom.Paragraph;
        c.VerticalOffset = 22d;
        // Set a buffer on left and right of picture where no text will be drawn.
        c.DistanceFromTextLeft = 5d;
        c.DistanceFromTextRight = 5d;

        var p = document.InsertParagraph( "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum. Contrary to popular belief, Lorem Ipsum is not simply random text. It has roots in a piece of classical Latin literature from 45 BC, making it over 2000 years old. Richard McClintock, a Latin professor at Hampden-Sydney College in Virginia, looked up one of the more obscure Latin words, consectetur, from a Lorem Ipsum passage, and going through the cites of the word in classical literature, discovered the undoubtable source. Lorem Ipsum comes from sections 1.10.32 and 1.10.33 of \"de Finibus Bonorum et Malorum\" (The Extremes of Good and Evil) by Cicero, written in 45 BC. This book is a treatise on the theory of ethics, very popular during the Renaissance. The first line of Lorem Ipsum, \"Lorem ipsum dolor sit amet..\", comes from a line in section 1.10.32." );
        p.Alignment = Alignment.both;
        p.InsertChart( c, 0, 250f, 200f );
        p.SpacingAfter( 30 );

        document.Save();
        Console.WriteLine( "\tCreated: AddChartWithTextWrapping.docx\n" );
      }


#else
      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.

#endif
    }
    #endregion
  }
}
