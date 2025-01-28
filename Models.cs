using AdmxParser;
using AdmxParser.Serialization;
using Microsoft.Win32;

namespace PodNet.EffectiveGroupPolicyExporter;

public record AdmxResult(AdmxContent Admx, IReadOnlyList<PolicyResult> Policies);
public record PolicyResult(PolicyDefinition Policy, IReadOnlyList<PolicyQueryResult> PolicyQueryResults);
public record PolicyQueryResult(PolicyQuery Query, RegistryQueryResult[] RegistryQueryResults);
public record PolicyQuery(PolicyClass Class, string Key, string? ValueName, PolicyQueryType QueryType);

public record AdmxJsonResult(string AdmxFilePath, PolicyJsonResult[] PolicyResults);
public record PolicyJsonResult(string Category, string PolicyName, string PolicyDisplayName, string? PolicyExplain, PolicyClass PolicyClass, RegistryQueryResult[] Registry);
public record RegistryQueryResult(string Key, string ValueName, object? Value, RegistryValueKind? ValueKind);

public enum PolicyQueryType
{
    MainValue,
    EnabledList,
    DisabledList,
    Boolean,
    Decimal,
    Enum,
    EnumList,
    List,
    LongDecimal,
    MultiText,
    Text
}
