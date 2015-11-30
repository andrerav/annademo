if not exist "./NuGet Packages" mkdir "./NuGet Packages"
del /Q "NuGet Packages\*"
call "./Tools/NuGet.exe" pack AnNa.SpreadsheetParser.EPPlus\AnNa.SpreadsheetParser.EPPlus.csproj -Build -IncludeReferencedProjects -Properties Configuration=Release -OutputDirectory "./NuGet Packages"
call "./Tools/NuGet.exe" pack AnNaSpreadsheetParser.SpreadsheetGear\AnNa.SpreadsheetParser.SpreadsheetGear.csproj -Build -IncludeReferencedProjects -Properties Configuration=Release -OutputDirectory "./NuGet Packages"

