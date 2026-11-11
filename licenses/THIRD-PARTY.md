# Bundled components

The runtime loads only files included under `lib`. Package versions and download SHA-256 hashes are recorded in `dependencies.json`.

| Component | Version | License / source |
|---|---|---|
| Tabula Sharp | 1.0.1 | MIT; https://github.com/BobLd/tabula-sharp |
| PdfPig | 0.1.14 | Apache-2.0; https://github.com/UglyToad/PdfPig |
| ImportExcel | 7.8.10 | Apache-2.0; https://github.com/dfinke/ImportExcel |
| EPPlus bundled with ImportExcel | Assembly 4.5.3.2 | LGPL-2.1; https://github.com/JanKallman/EPPlus |
| Microsoft.Bcl.HashCode | 6.0.0 | MIT; official NuGet package license included |
| System.Memory / System.Buffers / System.Numerics.Vectors | 4.6.0 | MIT; official NuGet package licenses included |
| System.Runtime.CompilerServices.Unsafe | 6.1.0 | MIT; official NuGet package license included |
| System.ValueTuple | 4.5.0 | MIT; official NuGet package license included |

The assemblies are unmodified. NuGet package metadata and license texts are retained in the corresponding subdirectories. ImportExcel's runtime scripts and EPPlus assembly are copied from the unmodified 7.8.10 module. Upstream EPPlus 4.x source is available at https://github.com/JanKallman/EPPlus ; keep the LGPL license and this source reference when redistributing. The separate DLL can be replaced with a compatible version in accordance with its license.

The three user-supplied report PDFs are not distributed with this tool.
