call "./Tools/NuGet462.exe" setApiKey #####YOUR-API-KEY-HERE##### -Source https://www.nuget.org/
for %%f in ("NuGet Packages\*.nupkg") do (
	call "./Tools/NuGet462.exe" Push "%%f" -Source https://www.nuget.org/
)