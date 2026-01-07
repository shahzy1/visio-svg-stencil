🧩 Azure Visio Stencil Builder

A lightweight engineering utility written in C# that converts Azure SVG icons into categorized Microsoft Visio stencil files (.vssx). The tool is intended to help engineers and architects create consistent, reusable architecture diagrams using standard Visio editions.

The project focuses on documentation enablement and workflow efficiency, allowing teams to generate Visio-native stencils that align with official Azure icon conventions while remaining compatible with standard Visio editions.

✨ Key Capabilities

Automatic categorization — generates one stencil per folder (e.g., AI, Compute, Storage) to keep diagrams organized and easy to maintain.

SVG import support — imports vector-based Azure icons directly into Visio using COM interop.

Consistent sizing — normalizes master dimensions to ensure clean, uniform diagrams.

Readable labels — adds standardized text labels beneath each shape for clarity.

Visio-native output — produces .vssx stencil files compatible with Visio 2016 and later.

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

This project is licensed under the MIT License — free to use, modify, and distribute.

👨‍💻 Author

Shahzad Khan
Azure Solutions Engineer | Cloud Platform & Integration
🔗 shahzadblog.com
