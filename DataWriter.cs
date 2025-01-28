using AdmxParser;
using AdmxParser.Serialization;
using Microsoft.Win32;
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PodNet.EffectiveGroupPolicyExporter;

public class DataWriter(TextWriter writer)
{
    public TextWriter Writer { get; } = writer;

    public void WriteLine(string? text = null) => Writer.WriteLine(text);

    public void WriteAll(IEnumerable<AdmxResult> results, string cultureCode)
    {
        var culture = CultureInfo.GetCultureInfo(cultureCode);
        foreach (var admxResult in results.Where(a => a.Policies.Any(p => p.PolicyQueryResults.Any(q => q.RegistryQueryResults.Any(r => r.Value is not null)))))
        {
            WriteLine($"\e[32m{admxResult.Admx.AdmxFilePath}\e[0m");
            foreach (var policyQuery in admxResult.Policies.Where(p => p.PolicyQueryResults.Any(q => q.RegistryQueryResults.Any(r => r.Value is not null))))
            {
                WriteLine($"  \e[33m{policyQuery.Policy.GetLocalizedDisplayName(admxResult.Admx.LoadedAdmlResources, culture)}\e[0m");
                foreach (var queryResults in policyQuery.PolicyQueryResults.Where(q => q.RegistryQueryResults.Any(r => r.Value is not null)))
                {
                    foreach (var result in queryResults.RegistryQueryResults.Where(r => r.Value is not null))
                    {
                        WriteLine($"    \e[34m{queryResults.Query.QueryType,-12} \e[35m{result.Key}\e[0m\\\e[36m{result.ValueName ?? "<null>"}\e[0m = \e[32m{(result.Value is null ? "<null>" : $"{result.Value} \e[33m({result.ValueKind})")}\e[0m");
                    }
                }
                WriteLine();
            }
        }
    }

    public void WriteAllJson(IEnumerable<AdmxJsonResult> results, bool pretty)
    {
#pragma warning disable CA1869 // Cache and reuse 'JsonSerializerOptions' instances | This is only used once per run
        var options = new JsonSerializerOptions(JsonSerializerOptions.Default)
        {
            WriteIndented = pretty,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
        };
#pragma warning restore CA1869 // Cache and reuse 'JsonSerializerOptions' instances
        options.Converters.Add(new JsonStringEnumConverter<RegistryValueKind>());
        options.Converters.Add(new JsonStringEnumConverter<PolicyClass>());
        WriteLine(JsonSerializer.Serialize(results, options));
    }
}
