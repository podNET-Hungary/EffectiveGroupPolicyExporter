using AdmxParser;
using AdmxParser.Serialization;
using Microsoft.Win32;
using System.Diagnostics;
using System.Globalization;
using System.Security;

namespace PodNet.EffectiveGroupPolicyExporter;

public class GroupPolicyExporter
{
    private static Dictionary<PolicyClass, (string Name, RegistryKey RegistryKey)[]> HivesMap { get; } = new()
    {
        [PolicyClass.User] = [("HKEY_CURRENT_USER", Registry.CurrentUser)],
        [PolicyClass.Machine] = [("HKEY_LOCAL_MACHINE", Registry.LocalMachine)],
        [PolicyClass.Both] = [("HKEY_CURRENT_USER", Registry.CurrentUser), ("HKEY_LOCAL_MACHINE", Registry.LocalMachine)],
    };
    public static IEnumerable<PolicyQuery> GetAllGroupPolicyRegistryQueries(PolicyDefinition policy)
    {
        if (policy.valueName is not null)
            yield return new(policy.@class, policy.key, policy.valueName, PolicyQueryType.MainValue);
        foreach (var item in policy.enabledList?.item ?? [])
            yield return new(policy.@class, item.key ?? policy.enabledList!.defaultKey, item.valueName, PolicyQueryType.EnabledList);
        foreach (var item in policy.disabledList?.item ?? [])
            yield return new(policy.@class, item.key ?? policy.disabledList!.defaultKey, item.valueName, PolicyQueryType.DisabledList);
        foreach (var element in policy.elements ?? [])
        {
            if (element is BooleanElement booleanElement)
                yield return new(policy.@class, booleanElement.key ?? policy.key, booleanElement.valueName, PolicyQueryType.Boolean);
            else if (element is DecimalElement decimalElement)
                yield return new(policy.@class, decimalElement.key ?? policy.key, decimalElement.valueName, PolicyQueryType.Decimal);
            else if (element is EnumerationElement enumerationElement)
            {
                yield return new(policy.@class, enumerationElement.key ?? policy.key, enumerationElement.valueName, PolicyQueryType.Enum);
                foreach (var item in enumerationElement.item)
                {
                    foreach (var listItem in (item.valueList?.item ?? []))
                        yield return new(policy.@class, listItem.key, listItem.valueName, PolicyQueryType.EnumList);
                }
            }
            else if (element is ListElement listElement)
                yield return new(policy.@class, listElement.key ?? policy.key, null, PolicyQueryType.List);
            else if (element is LongDecimalElement longDecimalElement)
                yield return new(policy.@class, longDecimalElement.key ?? policy.key, longDecimalElement.valueName, PolicyQueryType.LongDecimal);
            else if (element is multiTextElement multiTextElement)
                yield return new(policy.@class, multiTextElement.key ?? policy.key, multiTextElement.valueName, PolicyQueryType.MultiText);
            else if (element is TextElement textElement)
                yield return new(policy.@class, textElement.key ?? policy.key, textElement.valueName, PolicyQueryType.Text);
            else
                throw new InvalidOperationException($"Unknown element type: {element.GetType()}");
        }
    }


    public static IEnumerable<RegistryQueryResult> GetRegistryValuesForClass(PolicyClass @class, string key, string? valueName)
    {
        foreach (var (hiveName, hive) in HivesMap[@class])
        {
            RegistryKey? rKey;
            try
            {
                rKey = hive.OpenSubKey(key, RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.ReadKey);
            }
            catch (SecurityException ex)
            {
                Debug.WriteLine($"\e[41m\r\n{ex}\r\n\e[0m");
                continue;
            }

            if (valueName != null)
            {
                if (rKey == null)
                {
                    yield return new($@"{hiveName}\{key}", valueName, null, null);
                    continue;
                }
                var value = rKey.GetValue(valueName);
                yield return new($@"{hiveName}\{key}", valueName, value, value == null ? null : rKey.GetValueKind(valueName));
            }
            else
            {
                if (rKey == null)
                    continue;
                foreach (var subValueName in rKey.GetValueNames())
                {
                    var value = rKey.GetValue(valueName);
                    yield return new($@"{hiveName}\{key}", subValueName, value, value == null ? null : rKey.GetValueKind(valueName));
                }
            }
        }
    }

    private static HashSet<PolicyQueryType> AllQueryTypes { get; } = [.. Enum.GetValues<PolicyQueryType>()];
    public static IEnumerable<RegistryQueryResult> GetAllRegistryValues(PolicyQuery query)
    {
        if (AllQueryTypes.Contains(query.QueryType))
        {
            if (query.QueryType is not PolicyQueryType.List && query.ValueName is null)
                throw new InvalidOperationException("No value name provided. This is only allowed for group policy list elements.");
            foreach (var result in GetRegistryValuesForClass(query.Class, query.Key, query.ValueName))
                yield return result;
        }
        else
            throw new InvalidOperationException($"Invalid query type: {query.QueryType}");
        yield break;
    }

    public static async Task<IReadOnlyList<AdmxResult>> GetAllAdmxResultsAsync()
    {
        var allDirs = AdmxDirectory.GetInstalledMicrosoftPolicyTemplates().Prepend(AdmxDirectory.GetSystemPolicyDefinitions());
        await Task.WhenAll(allDirs.Select(d => d.LoadAsync()));

        return allDirs
            .SelectMany(dir => dir.LoadedAdmxContents)
            .Select(admx => new AdmxResult(admx,
                (admx.Policies ?? []).Select(p => new PolicyResult(p,
                    GetAllGroupPolicyRegistryQueries(p).Select(q => new PolicyQueryResult(q,
                        GetAllRegistryValues(q).ToArray())).ToList().AsReadOnly())
                ).ToList().AsReadOnly())
            ).ToList().AsReadOnly();
    }

    public static async Task<AdmxJsonResult[]> GetJsonAsync(bool includeExplain, string cultureCode)
    {
        var culture = CultureInfo.GetCultureInfo(cultureCode);
        return (await GetAllAdmxResultsAsync()).Where(a => a.Policies.Any(p => p.PolicyQueryResults.Any(q => q.RegistryQueryResults.Any(r => r.Value is not null)))).Select(a =>
        {
            var categories = a.Admx.Categories.ToDictionary(c => c.name);
            return new AdmxJsonResult(
                a.Admx.AdmxFilePath,
                a.Policies.Where(p => p.PolicyQueryResults.Any(q => q.RegistryQueryResults.Any(r => r.Value is not null)))
                    .Select(p => new PolicyJsonResult(
                        GetCategoryTree(p.Policy.parentCategory),
                        p.Policy.name,
                        AdmxResourceReference.Interpolate(p.Policy.displayName, a.Admx.LoadedAdmlResources, culture, true),
                        includeExplain ? AdmxResourceReference.Interpolate(p.Policy.explainText, a.Admx.LoadedAdmlResources, culture, true) : null,
                        p.Policy.@class,
                        p.PolicyQueryResults.SelectMany(pq => pq.RegistryQueryResults.Where(r => r.Value is not null)).Distinct().ToArray()
                        )).ToArray());
            string GetCategoryTree(CategoryReference categoryReference)
            {
                if (!categories.TryGetValue(categoryReference.@ref, out var category))
                    return categoryReference.@ref;
                var interpolated = AdmxResourceReference.Interpolate(category.displayName, a.Admx.LoadedAdmlResources, culture, true);
                if (category.parentCategory is { } parent)
                    return $@"{GetCategoryTree(parent)}\{interpolated}";
                return interpolated;
            }
        }).ToArray();
    }
}
