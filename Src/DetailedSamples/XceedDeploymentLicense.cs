/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Sample Application
Copyright (c) 2009-2023 - Xceed Software Inc.
 
This application demonstrates how to set a license when using the API 
from the Xceed Words for .NET.
 
This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

namespace Xceed.Words.NET.Examples
{
  public static class XceedDeploymentLicense
  {
    public static void SetLicense()
    {
      /* ================================
       * How to license Xceed components 
       * ================================
       *
       * To license (unlock) your component, set the LicenseKey property with your 
       * license key in the entry point of the application. This will ensure the component
       * is licensed before any of its methods are called.
       *  
       * If the component is used in a DLL project (no entry point is available), it is 
       * recommended that the LicenseKey property be set in a static constructor of a 
       * class that will be accessed systematically before any component is instantiated,
       * or you can simply set the LicenseKey property immediately BEFORE 
       * instantiation of the component. 
       * 
       * To deploy this sample, your license key should be set in the OnStartup() method.
       *
       * For more information, consult the "Licensing" topics in the product documentation. 
       */

      // Please uncomment the following line to set your own license key.
      // Xceed.Words.NET.Licenser.LicenseKey = "XXXXX-XXXXX-XXXXX-XXXX";	   
    }
  }
}


