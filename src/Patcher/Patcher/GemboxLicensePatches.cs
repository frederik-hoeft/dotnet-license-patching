using GemBox.Spreadsheet;
using HarmonyLib;
using System.Diagnostics;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;

namespace Patcher;

// this class contains patches for the Gembox license checks
public static class GemboxLicensePatches
{
    public static void ApplyAll()
    {
        Console.WriteLine("Applying Gembox license patches...");
        // create a new harmony instance
        Harmony harmony = new("MyPatcher");

        // manually apply patches to allow Harmony to support obfuscated classes and methods
        // patch the license check
        Console.WriteLine("Installing postfix patch for Gembox license check to override result to true...");
        // we don't know the exact class and method names, as they are obfuscared and change between versions. So we have to search for them
        Console.WriteLine("Searching for potential entry points for the license check patch starting from the SpreadsheetInfo class...");
        // however we know the "path" to the license check method, so we can start at a known entry point
        // and search for members that match what we know about the license check class and method
        // So similarly to finding stable pointer paths in C++ reverse engineering for consistent patches on ASLR-enabled systems,
        // we search for stable reflection paths in C# for reliable patches between different versions of the library.
        FieldInfo[] entryPoints = typeof(SpreadsheetInfo).GetFields(BindingFlags.NonPublic | BindingFlags.Static)
            // we look for a SpreadsheetInfo field with a signature like this:
            // .field private static initonly class '\u0005\u000F\u000F\u0003' '\u0003'
            // with a DebuggerBrowsableAttribute(DebuggerBrowsableState.Never)
            .Where(f => f.IsPrivate
                && f.IsInitOnly
                && f.GetCustomAttribute<DebuggerBrowsableAttribute>() is DebuggerBrowsableAttribute { State: DebuggerBrowsableState.Never }
                && f.FieldType.IsClass
                // it's funny, because we can use the unprintable characters in the obfuscated member names to our advantage
                // to filter out unrelated things that have nothing to do with the licensing :)
                // as of now, the obfuscated member names consist of unprintable characters
                && f.Name.All(c => char.IsControl(c))
                && f.FieldType.Name.All(c => char.IsControl(c))) // same for the field type
            .ToArray();

        Console.WriteLine($"Found {entryPoints.Length} entry point(s) for potential reflection path(s) to the license check method:");
        foreach (FieldInfo field in entryPoints)
        {
            Console.WriteLine($"- {field.Name.Hexlify()} ({field.FieldType.Name.Hexlify()})");
        }

        // we use a known signature of the license check method to find the correct one
        // this signature can be obtained either trough dynamic analysis (e.g. debugging) or static analysis looking at the IL code of the assembly
        Type[] signature =
        [
            typeof(string),
            typeof(string),
            typeof(Action<bool>),
            typeof(int).MakeByRefType(),
            typeof(int).MakeByRefType(),
            typeof(int).MakeByRefType(),
        ];

        Console.WriteLine($"Searching {entryPoints.Length} entry point(s) for obfuscated method with signature: ({string.Join(", ", signature.Select(t => t.Name))}) --> bool ...");

        // now that we have a list of potential entry points, we can search for the license check method which has a known signature
        MethodInfo licenseCheckTarget = entryPoints.SelectMany(f => f.FieldType.GetMethods(BindingFlags.Public | BindingFlags.Instance)
            .Where(method => method.ReturnType == typeof(bool)
                && method.GetParameters()
                    .Select(p => p.ParameterType)
                    .SequenceEqual(signature)))
            // again, use the unprintable characters in the obfuscated member names to pinpoint the correct method
            .SingleOrDefault(m => m.Name.All(c => char.IsControl(c))) ?? throw new InvalidOperationException("Could not find the license check method.");

        Console.WriteLine($"Found (presumably) the license check method: {licenseCheckTarget.DeclaringType?.Name.Hexlify()}::{licenseCheckTarget.Name.Hexlify()} through matching signature! Applying patch...");

        // now, we can let Harmony do its magic to strip the license checks to simply "return true"
        HarmonyMethod licenseCheckPatch = new(typeof(GemboxLicensePatches).GetMethod(nameof(Postfix_OverrideLicenseCheckResult)!));
        harmony.Patch(licenseCheckTarget, postfix: licenseCheckPatch);
        Console.WriteLine($"Successfully applied {nameof(Postfix_OverrideLicenseCheckResult)} patch!");

        // patch the worksheet getter
        // this one is an "if (<not licensed>) throw new InvalidLicenseException()" check
        // but we can again use Harmony to just remove all IL instructions before and up to the throw
        Console.WriteLine("Installing transpiler patch for Gembox ExcelFile.Worksheets getter to strip license checks from IL code...");
        MethodInfo worksheetGetterTarget = typeof(ExcelFile).GetProperty(nameof(ExcelFile.Worksheets))!.GetGetMethod()!;
        HarmonyMethod worksheetGetterPatch = new(typeof(GemboxLicensePatches).GetMethod(nameof(Transpiler_RemoveWorksheetGetterILChecks)!));
        harmony.Patch(worksheetGetterTarget, transpiler: worksheetGetterPatch);

        // we should be done now
        Console.WriteLine($"Successfully applied {nameof(Transpiler_RemoveWorksheetGetterILChecks)} patch!");
        Console.WriteLine("Successfully patched Gembox license checks!");
    }

    // the postfix patch works as a simple detour that is executed after the original method body, but before the return
    public static void Postfix_OverrideLicenseCheckResult(ref bool __result) =>
        // Override the license check to always return true
        __result = true;

    // the transpiler patch allows us to modify the IL code of the original method
    public static IEnumerable<CodeInstruction> Transpiler_RemoveWorksheetGetterILChecks(IEnumerable<CodeInstruction> instructions)
    {
        // remove every IL instruction before and including the first throw (if (...) throw new invalid license exception)
        // see dotPeek or ILSpy to find the exact IL code to remove (done once and probably won't ever change)
        bool foundThrow = false;
        foreach (CodeInstruction instruction in instructions)
        {
            // skip all instructions before (and including) the first throw
            if (!foundThrow)
            {
                if (instruction.opcode == OpCodes.Throw)
                {
                    foundThrow = true;
                    continue;
                }
            }
            else
            {
                // the rest is part of the actual getter implementation
                yield return instruction;
            }
        }
        if (!foundThrow)
        {
            // we would have to re-investigate the IL code through static analysis to update our patch
            throw new InvalidOperationException("Patch operation failed: could not find the throw instruction in the ExcelFile.Worksheets getter. The IL code may have changed.");
        }
    }

    private static string Hexlify(this string str) => $"0x{Convert.ToHexString(Encoding.ASCII.GetBytes(str)).ToLower()}";
}