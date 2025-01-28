# EffectiveGroupPolicyExporter
A simple C# console app that exports all applied group policies from the registry

## Prerequisites
- Windows
- .NET 9 SDK
- To reliably export all keys set on the machine, you need to run the app (or the IDE of your choice that runs the app) as admin; the app skips the entries that it can't read because of permission problems

## Usage
Running the app simply writes all group policy entries that are applied to the current PC to the console for review. The logic is that the app iterates through all installed ADMX files (many are installed by default, but you can install more if you need) and parses them, finds the available registry keys and valuenames, and writes the values to the console.

You can run the app directly using `dotnet run`, or you can `dotnet publish` to compile to a win-x64 native self-contained executable (not AOT and not trimmed though, so it's about 70MB, XML serialization is pretty error-prone using trimming unfortunately).

You can supply the following parameters:
```
[--json [--pretty --include-explain]] [--culture {cultureCode}]

--json: writes the set policy entries to the standard output.
    --pretty: indents the JSON to make it easier to parse for humans.
    --include--explain: includes the policy explanation text (en-US). Note that this greatly increases the JSON payload size.
--culture {cultureCode}: sets the culture. This affects the display names and explain texts. Defaults to 'en-US' if not provided.
```

You can save or parse the output by using your preferred shell. For example, in PowerShell:
```ps1
$rawResults = (dotnet run --json)
Set-Content -Path "effective-group-policy.json" -Value $rawResults # Save the contents in a file.
```

Then do what you wish with the set policies:
```ps1
$results = (Get-Content "effective-group-policy.json" | ConvertFrom-Json | Select-Object -ExpandProperty PolicyResults | Select-Object -ExpandProperty Registry)
foreach ($result in $results) {
    if ('y' -eq (Read-Host "Do you want to set the following registry entry:\n $result.Key\t$result.ValueName = $result.Value ?"))
        Set-ItemProperty -Path $result.Key -Name $result.ValueName -Value $result.Value
}
```

## Troubleshooting
This is a small app that's provided for referece, as is, with no warranties whatsoever of any kind. You can clone, debug and modify it to your heart's content. There's no additional development planned for the tool, but feel free to reach out if you need help. Additionally, feel free to help others who reach out as well 😏