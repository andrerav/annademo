call "./Tools/NuGet.exe" SetApiKey 9e527d6d-4bc1-4c59-b24f-f0a152a39348
for %%f in ("NuGet Packages\*.nupkg") do (
	call "./Tools/NuGet.exe" Push "%%f"
)