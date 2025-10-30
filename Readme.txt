Setup Instructions

Open Visual Studio → Create new Console App (.NET Framework) project (not .NET Core).

Add COM Reference:

Right-click References → Add Reference → COM → Microsoft Visio 16.0 Type Library

Paste the code above into Program.cs.

Update:

string svgFolder = @"C:\Temp\AzureSVGs";
string stencilPath = @"C:\Temp\AzureStencil.vssx";


Run it.

Your stencil will be created at the specified path — one master per SVG, each centered and ready to use in Visio.
