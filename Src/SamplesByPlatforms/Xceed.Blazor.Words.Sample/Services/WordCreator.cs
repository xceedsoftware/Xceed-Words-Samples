using Microsoft.JSInterop;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace Xceed.Blazor.Words.Sample.Services
{
	public class WordCreator
	{
		private readonly IJSRuntime jsRuntime;

		public WordCreator( IJSRuntime _jsRuntime )
		{
			jsRuntime = _jsRuntime;
		}

		public async Task CreateSimpleDoc()
		{
			var doc = DocX.Create( "simple_doc.docx" );

			doc.InsertParagraph( "The History of Intel" )
				.FontSize( 18 )
				.Bold()
				.Alignment = Xceed.Document.NET.Alignment.center;

			doc.InsertParagraph( "Intel Corporation was founded in 1968 by Robert Noyce and Gordon Moore. Initially, it was known for its semiconductor products, especially DRAM, SRAM, and ROM chips. In 1971, Intel introduced the world's first microprocessor, the Intel 4004. This innovation paved the way for the modern computing era. Over the years, Intel has continued to lead in technology, driving the PC revolution in the 1980s with its x86 architecture. Today, Intel is a major player in the fields of AI, autonomous driving, and data centers. The company’s commitment to innovation remains steadfast, influencing technological advancements worldwide." )
				.FontSize( 12 );

			doc.InsertSectionPageBreak();

			doc.InsertParagraph( "The History of Apple" )
				.FontSize( 18 )
				.Bold()
				.Alignment = Xceed.Document.NET.Alignment.center;

			doc.InsertParagraph( "Apple Inc. was established in 1976 by Steve Jobs, Steve Wozniak, and Ronald Wayne. The company revolutionized personal computing with the introduction of the Apple II in 1977. In 1984, Apple launched the Macintosh, the first personal computer with a graphical user interface and a mouse. Despite some turbulent years, the return of Steve Jobs in 1997 marked a new era of innovation. Products like the iPod, iPhone, and iPad have transformed entire industries. Today, Apple continues to set trends in technology with its focus on design, user experience, and cutting-edge technology." )
				.FontSize( 12 );

			doc.InsertSectionPageBreak();

			doc.InsertParagraph( "The History of Tesla" )
				.FontSize( 18 )
				.Bold()
				.Alignment = Xceed.Document.NET.Alignment.center;

			doc.InsertParagraph( "Tesla, Inc. was founded in 2003 by Martin Eberhard and Marc Tarpenning, with Elon Musk joining shortly after. The company aims to accelerate the world’s transition to sustainable energy. Tesla’s first car, the Roadster, demonstrated that electric vehicles could be both high-performance and environmentally friendly. The subsequent releases of the Model S, Model X, and Model 3 have solidified Tesla's reputation in the automotive industry. Tesla's innovation extends beyond cars to energy solutions like the Powerwall and Solar Roof. The company continues to push the boundaries of technology and sustainability." )
				.FontSize( 12 );

			doc.Save();
			await DownloadFile( "simple_doc.docx" );
			doc.Dispose();
		}

		public async Task CreateListedDoc()
		{
			var doc = DocX.Create( "listed_doc.docx" );

			doc.InsertParagraph( "How to Make Cuban Moros y Cristianos" )
				.FontSize( 18 )
				.Bold()
				.Alignment = Xceed.Document.NET.Alignment.center;

			doc.InsertParagraph( "Moros y Cristianos, or 'Moors and Christians', is a classic Cuban dish consisting of black beans and rice. The name reflects the cultural blend that characterizes much of Cuban cuisine. This dish is a staple at family gatherings and celebrations. The combination of flavors and textures makes it a favorite among many." )
				.FontSize( 12 );

			doc.InsertParagraph( "\nIngredients:" )
				.FontSize( 14 )
				.Bold();

			doc.InsertParagraph( "1. 1 cup black beans\n2. 1 cup white rice\n3. 1 onion, finely chopped\n4. 1 green bell pepper, chopped\n5. 3 cloves garlic, minced\n6. 2 cups water\n7. 1 teaspoon cumin\n8. 1 bay leaf\n9. Salt to taste\n10. Pepper to taste" )
				.FontSize( 12 );

			doc.InsertParagraph( "\nSteps:" )
				.FontSize( 14 )
				.Bold();

			doc.InsertParagraph( "• Soak the black beans overnight, then drain and rinse.\n• In a large pot, sauté the onion, bell pepper, and garlic until tender.\n• Add the black beans, water, cumin, bay leaf, salt, and pepper. Bring to a boil, then simmer until the beans are tender.\n• Cook the white rice separately according to package instructions.\n• Mix the cooked rice with the black beans and serve." )
				.FontSize( 12 );

			doc.Save();
			await DownloadFile( "listed_doc.docx" );
			doc.Dispose();
		}

		public async Task CreateTableDoc()
		{
			var doc = DocX.Create( "table_doc.docx" );

			doc.InsertParagraph( "Technical Specifications of a High-Performance Laptop" )
				.FontSize( 18 )
				.Bold()
				.Alignment = Xceed.Document.NET.Alignment.center;

			doc.InsertParagraph( "This document provides the detailed technical specifications of a high-performance laptop. These specifications include key components and their respective features, ensuring that users understand the capabilities of the device." )
				.FontSize( 12 );

			var table = doc.AddTable( 17, 2 );
			table.Design = Xceed.Document.NET.TableDesign.LightListAccent1;
			table.Alignment = Xceed.Document.NET.Alignment.center;
			table.Rows[ 0 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "Component" ).Bold();
			table.Rows[ 0 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "Specification" ).Bold();

			string[] components = { "Processor", "RAM", "Storage", "Graphics Card", "Display", "Battery Life", "Operating System", "Weight", "Dimensions", "Ports", "Audio", "Keyboard", "Touchpad", "Wireless Connectivity", "Camera", "Warranty" };
			string[] specifications = { "Intel Core i7", "16GB DDR4", "512GB SSD", "NVIDIA GeForce RTX 3060", "15.6\" FHD", "Up to 10 hours", "Windows 10 Pro", "1.8 kg", "35.8 x 24.6 x 1.8 cm", "USB-C, USB-A, HDMI", "Stereo Speakers", "Backlit Keyboard", "Precision Touchpad", "Wi-Fi 6, Bluetooth 5.0", "720p HD Camera", "1 Year" };

			for( int i = 0; i < components.Length; i++ )
			{
				table.Rows[ i + 1 ].Cells[ 0 ].Paragraphs[ 0 ].Append( components[ i ] );
				table.Rows[ i + 1 ].Cells[ 1 ].Paragraphs[ 0 ].Append( specifications[ i ] );
			}

			doc.InsertTable( table );

			doc.Save();
			await DownloadFile( "table_doc.docx" );

			doc.Dispose();
		}

		private async Task DownloadFile( string fileName )
		{
			var bytes = await File.ReadAllBytesAsync( fileName );
			var base64 = Convert.ToBase64String( bytes );
			await jsRuntime.InvokeVoidAsync( "BlazorDownloadFile", fileName, base64 );
		}

	}
}
