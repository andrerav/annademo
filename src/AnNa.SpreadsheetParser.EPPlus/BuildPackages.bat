if not exist "./NuGet Packages" mkdir "./NuGet Packages"
del /Q "NuGet Packages\*"
call "./../Tools/NuGet.exe" pack AnNa.SpreadsheetParser.EPPlus.csproj -Build -IncludeReferencedProjects -Properties Configuration=Release -OutputDirectory "./NuGet Packages"
