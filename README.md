🧩 Azure Visio Stencil Builder

A lightweight C# utility that converts Azure SVG icons into categorized Visio stencil files (.vssx) — complete with proper scaling, labels, and organization.

It helps teams quickly generate stencils that look and behave like the official Microsoft Azure icon sets.

✨ Features

📂 Automatic categorization — creates one stencil per folder (e.g., AI, Compute, Storage).

🖼️ SVG import — imports vector icons directly into Visio.

📏 Consistent sizing — normalizes master dimensions for a clean look.

🏷️ Text labels — adds readable icon names below each shape.

🧱 Visio-native output — generates .vssx stencil files compatible with Visio 2016+.

🛠️ Prerequisites

Windows 10 or 11

Microsoft Visio 2019 (or later)

.NET Framework 4.8+

Microsoft.Office.Interop.Visio (COM reference)

Azure SVG icon pack (available from Microsoft Azure Architecture Icons)

🚀 Usage

Clone the repo:

git clone https://github.com/<yourname>/AzureStencilBuilder.git


Place your Azure SVG icons into category folders:

C:\Temp\AzureSVGs\
├── AI\
├── Compute\
├── Storage\
└── Networking\


Build and run the project:

dotnet run


The tool creates categorized stencils in:

C:\Temp\Stencils\
├── Azure-AI.vssx
├── Azure-Compute.vssx
├── Azure-Storage.vssx
└── Azure-Networking.vssx

🧰 Configuration

You can change these paths inside Program.cs:

string baseFolder = @"C:\Temp\AzureSVGs";   // Source SVG folders
string outputFolder = @"C:\Temp\Stencils";  // Output Visio stencils

💡 Tips

Run Visual Studio as Administrator to avoid COM permission issues.

For long names, you can tweak text font size or wrap width in the script.

For better experience with Visio, copy stencil files to C:\Users\<username>\Documents\My Shapes folder

Works great for internal design documentation or architecture diagrams.

📄 License

This project is licensed under the MIT License — feel free to use and modify it.

👨‍💻 Author

Shahzad Khan
Senior Azure Developer | Cloud & AI Engineer
🔗 shahzadblog.com