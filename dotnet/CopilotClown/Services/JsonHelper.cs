using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace CopilotClown.Services;

/// <summary>
/// Thin JSON wrapper using JavaScriptSerializer (built into .NET Framework 4.8).
/// Zero external NuGet dependencies.
/// </summary>
public static class JsonHelper
{
    private static readonly JavaScriptSerializer Serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

    public static string Serialize(object obj) => Serializer.Serialize(obj);

    public static T Deserialize<T>(string json) => Serializer.Deserialize<T>(json);

    public static Dictionary<string, object> Parse(string json) =>
        Serializer.Deserialize<Dictionary<string, object>>(json);

    /// <summary>
    /// Navigate a parsed JSON object by dot-separated path.
    /// Example: GetValue(dict, "choices.0.message.content")
    /// </summary>
    public static object GetValue(Dictionary<string, object> root, string path)
    {
        var parts = path.Split('.');
        object current = root;

        foreach (var part in parts)
        {
            if (current == null) return null;

            if (int.TryParse(part, out int index))
            {
                if (current is object[] arr && index < arr.Length)
                    current = arr[index];
                else if (current is System.Collections.ArrayList list && index < list.Count)
                    current = list[index];
                else
                    return null;
            }
            else if (current is Dictionary<string, object> dict)
            {
                current = dict.TryGetValue(part, out var val) ? val : null;
            }
            else
            {
                return null;
            }
        }

        return current;
    }

    public static string GetString(Dictionary<string, object> root, string path) =>
        GetValue(root, path)?.ToString() ?? "";

    public static int GetInt(Dictionary<string, object> root, string path)
    {
        var val = GetValue(root, path);
        if (val is int i) return i;
        if (val != null && int.TryParse(val.ToString(), out int parsed)) return parsed;
        return 0;
    }
}
