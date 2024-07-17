![NuGet Downloads](https://img.shields.io/nuget/dt/Xceed.Words.NET) ![Static Badge](https://img.shields.io/badge/.Net_Framework-4.0%2B-blue) ![Static Badge](https://img.shields.io/badge/.Net-5.0%2B-blue) [![Learn More](https://img.shields.io/badge/Learn-More-blue?style=flat&labelColor=gray)](https://xceed.com/en/our-products/product/Words-for-net)

[![Xceed Words for .NET](./Resources/words_banner.png)](https://xceed.com/en/our-products/product/Words-for-net)

# Xceed Words for .NET - Examples

This repository contains a variety of sample applications to help you get started with using the Xceed Words for .NET in your own projects.

## Overview

Xceed Words for .NET allows you to create or manipulate Microsoft Word documents from your .NET applications, without requiring Word or Office to be installed. Convert Word documents to PDF (not all Word elements are supported; use the trial period to validate all required elements do get converted). Fast and lightweight. Widely used and backed by a responsive support and development team whose sole ambition is your complete satisfaction.

## About The Product

With its easy to use API, Xceed Words for .NET lets your application create new Microsoft Word .docx or PDF documents, or modify existing .docx documents. It gives you complete control over all content in a Word document, and lets you add or remove all commonly used element types, such as paragraphs, bulleted or numbered lists, images, tables, charts, headers and footers, sections, bookmarks, and more. Create PDF documents using the same API for creating Word documents.

You also get complete control over the document’s properties, including margins, page size, line spacing, page numbering, text direction and alignment, indentation, and more.

You can also quickly and easily set or modify any text’s formatting, fonts and font sizes, colors, boldness, underline, italics, strikethrough, highlighting, and more.

Search and replace text, add or remove password protection, join documents, copy documents, or apply templates – everything your application may need to do. It even supports modifying many Word files in parallel for greater speed. Key features include:

- **Create Word Documents**: Generate professional Word documents from scratch or based on templates.
- **Edit Existing Documents**: Modify content, format text, and update document properties programmatically.
- **Convert Documents**: Convert Word documents to various formats, including PDF, HTML, and more.
- **Rich Text Formatting**: Apply advanced text formatting, including styles, fonts, colors, and more.
- **Tables and Lists**: Create and manipulate tables, lists, and other complex document structures.
- **Headers and Footers**: Add and customize headers and footers, including page numbers and logos.
- **Images and Shapes**: Insert and manage images, shapes, and other graphical elements.
- **Merge and Split Documents**: Combine multiple documents into one or split a document into multiple parts.

For more information, please visit the [official product page](https://xceed.com/en/our-products/product/Words-for-net).

### Why Choose Xceed Words for .NET?

- Full featured. Latest releases add PDF capabilities.
- Supports .docx documents from Word 2007 and up
- Over 250,000 downloads
- Comprehensive documentation and sample applications included.
- Supports .NET 4.5, 5, 6 and 7

## Getting Started with Xceed Words for .NET.

To get started, clone this repository and explore the various sample projects provided. Each sample demonstrates different features and capabilities of Xceed Words for .NET.

### Requirements
- Visual Studio 2015 or later
- .NET Framework 4.0 or later
- .NET 5.0 or later

### 1. Installing the Xceed Words for .NET from nuget
To install the Xceed Words for .NET from NuGet, follow these steps:

1. **Open your project in Visual Studio.**
2. **Open the NuGet Package Manager Console** by navigating to `Tools > NuGet Package Manager > Package Manager Console`.
3. **Run the following command:**
```sh
   dotnet add package Xceed.Words.NET
```

4. Alternatively, you can use the NuGet Package Manager GUI:

1. Right-click on your project in the Solution Explorer.
2. Select Manage NuGet Packages.
3. Search for Xceed.Words.NET and click Install.

![Nuget library](./Resources/nuget_sample.png)

### 2. Refering Xceed Words for .NET library

1. **Add the reference with using statement at the top of the class**
   ```
   using Xceed.Words.NET;
   ```
   
2. **Use the classes and elements from the namespace**
   ```c#
    using Xceed.Words.NET;
    
    namespace BlazeDocX.Services
    {
      public class CVCreator
      {
        private readonly IJSRuntime jsRuntime;

          public CVCreator( IJSRuntime _jsRuntime )
          {
            jsRuntime = _jsRuntime;
          }
          ...
          public async Task CreateDoc( CV cv )
          {
            var docX = DocX.Create( "profile.docx" );
            var fullname = docX.InsertParagraph( cv.FullName.ToString().ToUpper() );
            fullname.Alignment = Alignment.center;
            fullname.Bold( true );
            fullname.FontSize( 14 );
            fullname.Font( "Swiss Light 10pt" );
      
            var occupation = docX.InsertParagraph( cv.Occupation.ToString() );
            occupation.Alignment = Alignment.center;
            occupation.Bold( true );
            occupation.FontSize( 10 );
            occupation.Font( "Arial" );
      
            var email = docX.InsertParagraph( cv.Email.ToString().ToLower() );
            email.Alignment = Alignment.center;
            email.Bold( true );
            email.FontSize( 10 );
            email.Font( "Arial" );
      
            email.InsertHorizontalLine( HorizontalBorderPosition.bottom, BorderStyle.Tcbs_single, size: 12, space: 10 );
      
            using MemoryStream memStream = new();
            docX.SaveAs( memStream );
            await jsRuntime.InvokeVoidAsync( "blazeDocX.downloadStream", memStream.GetBuffer(), $"CV - {cv.FullName}.docx" );
      
            docX.Dispose();
          }
           ...
       }
   }
   ```

   ### 3. How to License the Product Using the LicenseKey Property
To license the Xceed Forkbooks for .NET using the LicenseKey property, follow these steps:

1. **Obtain your license key** from Xceed. (Download the product from xceed.com or send us a request at support@xceed.com
2. **Set the LicenseKey property in your application startup code:**

   2.1 In case of WPF or Desktop app could be in the MainWindow
   ```csharp
   using System.Windows;

   public partial class MainWindow : Window
   {
       public MainWindow()
       {
           InitializeComponent();
           Xceed.Words.NET.Licenser.LicenseKey = "Your-Key-Here";
       }
   }
   ```
   2.2 In case of ASP.NET application must be in Program.cs class
   ```csharp
   using System.Net;
   using System.Text.Json;
   using System.Text.Json.Serialization;
   ...
   using Xceed.Document.NET;
   ...
   Xceed.Words.NET.Licenser.LicenseKey = "Your-Key-Here";
   ...
   var builder = WebAssemblyHostBuilder.CreateDefault(args);
   ```
4. Ensure the license key is set before any Words class, instance or similar control is instantiated.

## Sample Applications
### Basic Usage
A simple example showing how to create a document with 3 lines of text, a name, an occupation and underline email.

```csharp
    public static void CreateSimpleDoc()
    {
      using( var docX = DocX.Create( SomeDirectory + @"profile.docx" ) )
      {
        var fullname = docX.InsertParagraph( "John Doe" );
        fullname.Alignment = Alignment.center;
        fullname.Bold( true );
        fullname.FontSize( 14 );
        fullname.Font( "Swiss Light 10pt" );

        var occupation = docX.InsertParagraph( "Software Developer" );
        occupation.Alignment = Alignment.center;
        occupation.Bold( true );
        occupation.FontSize( 10 );
        occupation.Font( "Arial" );

        var email = docX.InsertParagraph( "doe@xceed.com" );
        email.Alignment = Alignment.center;
        email.Bold( true );
        email.FontSize( 10 );
        email.Font( "Arial" );
        email.InsertHorizontalLine( HorizontalBorderPosition.bottom, BorderStyle.Tcbs_single, size: 12, space: 10 );        
        docX.Save();
      }
    }

```

## Examples Overview

Below is a list of the examples available in this repository:

- **Bookmarks**: Demonstrates how to add and replace the text of the bookmarks.
- **Chart**: Shows how to insert and work with charts (Bar, Line, Pie, 3D, text wrapping).
- **Checkbox**: Provides a number of examples about how works with checkboxes (Add and modify).
- **Digital Signature**: Demonstrates how to add, verify and sign with Digital signature.
- **Document**: Shows how to use basics operations (Replace text, replace objects, add customs properties, apply template, append document, insert document, load with file name, stream or url, add html, rtf).
- **Equation**: Demonstrates how to insert an equation into a document.
- **FootnoteEndnote**: Provides examples how manage foot notes (Add foot notes, custom foot notes, End notes).
- **HeaderFooter**: Shows how to use headers and footers.
- **Hyperlink**: Demonstrates how to work and handle with hyperlinks.
- **Hyphenation**: Provides a number of examples about how works with hyphenations.
- **Image**: Demonstrates how to add, copy and modify images.
- **Line**: Shows how to insert an horizontal line.
- **List**: Demonstrates working with List (Add, modify, clone, numbered list, bulleted list, chapter).
- **Margin**: Shows how to handle with margins (Directions and identation).
- **Paragraph**: Demonstrates how to add, copy and modify paragraphs.
- **Pdf**: Provides examples how create as pdf.
- **Protection**: Provides a number of examples about how works with protections (Change password add password).
- **Section**: Shows how to add and set orientations.
- **Shape**: Provides a number of examples about how handle shapes.
- **Table**: Shows how to work with tables (add new, insert, add text wrapping, clone, text directions, columns width, merge cells).
- **TableOfContent**:Demonstrates how to add, copy and modify table of contents.

## Getting Started with the Samples

To get started with these examples, clone the repository and open the solution file in Visual Studio.

```bash
git clone https://github.com/your-repo/Xceed-Words-Samples.git
cd xceed-Words-samples
```
Open the solution file in Visual Studio and build the project to restore the necessary NuGet packages.
  
## Documentation

For more information on how to use the Xceed Words for .NET, please refer to the [official documentation](https://doc.xceed.com/xceed-document-libraries-for-net/webframe.html#rootWelcome.html).

## Licensing

To receive a license key, visit [xceed.com](https://xceed.com) and download the product, or contact us directly at [support@xceed.com](mailto:support@xceed.com) and we will provide you with a trial key.

## Contact

If you have any questions, feel free to open an issue or contact us at [support@xceed.com](mailto:support@xceed.com).

---

© 2024 Xceed Software Inc. All rights reserved.
