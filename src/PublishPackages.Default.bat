call "./Tools/NuGet.exe" SetApiKey ####API-KEY-FROM-NUGET-GOES-HERE####
for %%f in ("NuGet Packages\*.nupkg") do (
	call "./Tools/NuGet.exe" Push "%%f"
)