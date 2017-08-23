using Microsoft.VisualStudio.Setup.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace NuGetMsbuildChecker
{
    public static class MsbuildChecker
    {
        private readonly static string[] MSBuildVersions = new string[] { "14", "12", "4" };

        public static IEnumerable<MsBuildToolset> PrintMsbuildToolSet()
        {
            Console.WriteLine("Fetching all msbuild toolsets");
            var installedToolsets = new List<MsBuildToolset>();

            using (var projectCollection = LoadProjectCollection())
            {
                var installed = ((dynamic)projectCollection)?.Toolsets;
                if (installed != null)
                {
                    foreach (var item in installed)
                    {
                        installedToolsets.Add(new MsBuildToolset(version: item.ToolsVersion, path: item.ToolsPath));
                    }

                    installedToolsets = installedToolsets.ToList();
                }
            }

            var installedSxsToolsets = GetInstalledSxsToolsets();
            if (installedToolsets == null)
            {
                installedToolsets = installedSxsToolsets;
            }
            else if (installedSxsToolsets != null)
            {
                installedToolsets.AddRange(installedSxsToolsets);
            }

            Console.WriteLine($"Installed toolsets count: {installedToolsets.Count}");

            foreach (var toolset in installedToolsets)
            {
                Console.WriteLine($"toolset version: {toolset.Version}, toolset path: {toolset.Path}");
            }

            return installedToolsets;
        }

        public static void CheckVersion(string userVersion)
        {
            var installedToolsets = PrintMsbuildToolSet();
            var toolset = GetToolsetFromUserVersion(userVersion, installedToolsets);

            if (toolset != null)
            {
                Console.WriteLine("Found msbuild toolset");
                Console.WriteLine($"toolset version: {toolset.Version}, toolset path: {toolset.Path}");
            }
        }

        private static IDisposable LoadProjectCollection()
        {

            foreach (var version in MSBuildVersions)
            {
                try
                {
                    var msBuildTypesAssembly = Assembly.Load($"Microsoft.Build, Version={version}.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a");
                    var projectCollectionType = msBuildTypesAssembly.GetType("Microsoft.Build.Evaluation.ProjectCollection", throwOnError: true);
                    return Activator.CreateInstance(projectCollectionType) as IDisposable;
                }
                catch (Exception)
                {
                }
            }

            return null;
        }

        private static MsBuildToolset GetToolsetFromUserVersion(
            string userVersion,
            IEnumerable<MsBuildToolset> installedToolsets)
        {
            // Force version string to 1 decimal place
            var userVersionString = userVersion;
            decimal parsedVersion = 0;
            if (decimal.TryParse(userVersion, out parsedVersion))
            {
                var adjustedVersion = (decimal)(((int)(parsedVersion * 10)) / 10F);
                userVersionString = adjustedVersion.ToString("F1");
            }

            // First match by string comparison
            var selectedToolset = installedToolsets.FirstOrDefault(
                t => string.Equals(userVersionString, t.Version, StringComparison.OrdinalIgnoreCase));

            if (selectedToolset != null)
            {
                return selectedToolset;
            }

            // Then match by Major & Minor version numbers. And we want an actual parsing of t.ToolsVersion,
            // without the safe fallback to 0.0 built into t.ParsedToolsVersion.
            selectedToolset = installedToolsets.FirstOrDefault(t =>
            {
                Version parsedUserVersion;
                Version parsedToolsVersion;
                if (Version.TryParse(userVersionString, out parsedUserVersion) &&
                    Version.TryParse(t.Version, out parsedToolsVersion))
                {
                    return parsedToolsVersion.Major == parsedUserVersion.Major &&
                        parsedToolsVersion.Minor == parsedUserVersion.Minor;
                }

                return false;
            });

            if (selectedToolset == null)
            {
                Console.WriteLine("Can't find msbuild toolset");
                Console.WriteLine($"userVersionString: {userVersionString}");

                foreach (var toolset in installedToolsets)
                {
                    Console.WriteLine($"toolset version: {toolset.Version}, toolset path: {toolset.Path}");
                }
            }

            return selectedToolset;
        }

        private static List<MsBuildToolset> GetInstalledSxsToolsets()
        {
            ISetupConfiguration configuration;
            try
            {
                configuration = new SetupConfiguration() as ISetupConfiguration2;
            }
            catch (Exception)
            {
                return null; // No COM class
            }

            if (configuration == null)
            {
                return null;
            }

            var enumerator = configuration.EnumInstances();
            if (enumerator == null)
            {
                return null;
            }

            var setupInstances = new List<MsBuildToolset>();
            while (true)
            {
                var fetchedInstances = new ISetupInstance[3];
                int fetched;
                enumerator.Next(fetchedInstances.Length, fetchedInstances, out fetched);
                if (fetched == 0)
                {
                    break;
                }

                // fetched will return the value 3 even if only one instance returned
                var index = 0;
                while (index < fetched)
                {
                    if (fetchedInstances[index] != null)
                    {
                        setupInstances.Add(new MsBuildToolset(fetchedInstances[index]));
                    }

                    index++;
                }
            }

            if (setupInstances.Count == 0)
            {
                return null;
            }

            return setupInstances;
        }
    }
}
