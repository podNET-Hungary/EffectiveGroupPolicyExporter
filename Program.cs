using PodNet.EffectiveGroupPolicyExporter;

var writer = new DataWriter(Console.Out);
var cultureCode = args.SkipWhile(a => a != "--culture").Skip(1).FirstOrDefault() ?? "en-US";

if (args.Contains("--json"))
    writer.WriteAllJson(await GroupPolicyExporter.GetJsonAsync(args.Contains("--include-explain"), cultureCode), args.Contains("--pretty"));
else
    writer.WriteAll(await GroupPolicyExporter.GetAllAdmxResultsAsync(), cultureCode);