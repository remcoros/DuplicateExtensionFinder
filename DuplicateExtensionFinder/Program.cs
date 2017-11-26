namespace DuplicateExtensionFinder
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Xml;
    using System.Xml.Serialization;

    class Program
    {
        static void Main(string[] args)
        {
            var paths = args.Where(arg => !arg.StartsWith("-")).ToArray();

            var invalidPaths = paths
                .Where(dir => !Directory.Exists(dir))
                .ToArray();

            if (invalidPaths.Any())
            {
                Console.WriteLine("Warning: Skipping non-exisiting folder(s): " + string.Join(", ", paths));
                Console.WriteLine();
            }

            var instances = !paths.Any() ? GetInstances() : new[] { new VisualStudioInstance("Custom", paths) };

            var onlyDupes = args.Any(x => string.Equals(x, "-dupes", StringComparison.OrdinalIgnoreCase));
            var doDelete = args.Any(x => string.Equals(x, "-delete", StringComparison.OrdinalIgnoreCase));

            foreach (var instance in instances)
            {
                var extensions = instance.Extensions;
                if (!extensions.Any())
                    continue;

                if (onlyDupes && !extensions.SelectMany(ext => ext).Any(ext => ext.IsDuplicate))
                    continue;

                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine(new string('*', 100));
                Console.WriteLine("******  {0} ({1})", instance.Name, string.Join(", ", instance.ExtensionDirectories));
                Console.WriteLine(new string('*', 100));
                Console.WriteLine();

                foreach (var group in extensions)
                {
                    if (onlyDupes && group.Count() <= 1)
                        continue;

                    Console.WriteLine("{0}", group.First().Name);

                    foreach (var extension in group.OrderBy(x => x.Version).ThenBy(x => x.CreationTime))
                    {
                        Console.WriteLine(" - {0} [{2}] ({1})", extension.Version, extension.Path, extension.IsDuplicate ? "DELETE" : "KEEP");

                        if (!extension.IsDuplicate)
                            continue;

                        if (!doDelete)
                            continue;

                        try
                        {
                            Directory.Delete(extension.Path, true);
                        }
                        catch (System.UnauthorizedAccessException)
                        {
                            Console.WriteLine();
                            Console.WriteLine("You must start as administrator to delete global extensions.");
                            Console.WriteLine("Press any key to continue...");
                            Console.ReadKey();
                            return;
                        }
                    }

                    Console.WriteLine();
                }
            }

            if (instances.SelectMany(inst => inst.Extensions).SelectMany(group => group).All(ext => !ext.IsDuplicate))
            {
                Console.WriteLine("No duplicates found.");
            }
            else if (!doDelete)
            {
                Console.WriteLine("Specify '-delete' to delete old extensions from disk.");
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        private static VisualStudioInstance[] GetInstances()
        {
            var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Microsoft\VisualStudio");
            var extensionDirs = new DirectoryInfo(root).EnumerateDirectories("Extensions", SearchOption.AllDirectories)
                .Where(dir => string.Equals(dir.Parent?.Parent?.FullName, root, StringComparison.OrdinalIgnoreCase));

            return extensionDirs.Select(d => new VisualStudioInstance(d.Parent.Name, new[] { d.FullName })).ToArray();
        }
    }

    public class VisualStudioInstance
    {
        private static readonly XmlSerializer _vsixSerializer = new XmlSerializer(typeof(Vsix));
        private static readonly XmlSerializer _packageSerializer = new XmlSerializer(typeof(PackageManifest));
        private static readonly string[] _manifestNames = new[] { "extension.vsixmanifest", "extension.vsixmanifest.deleteme" };

        private static readonly string _programFilesFolder = Environment.GetFolderPath(Environment.Is64BitOperatingSystem ? Environment.SpecialFolder.ProgramFilesX86 : Environment.SpecialFolder.ProgramFiles);
        private static readonly string _vsGlobalExtensionFolderFormat = Path.Combine(_programFilesFolder, @"Microsoft Visual Studio {0}.0\Common7\IDE\Extensions");

        public VisualStudioInstance(string name, IEnumerable<string> extensionDirectories)
        {
            Name = name;
            ExtensionDirectories = extensionDirectories.ToList();

            if ((name.Length >= 4) && double.TryParse(name.Substring(0, 4), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var version))
            {
                if (version <= 14)
                {
                    ExtensionDirectories.Add(string.Format(CultureInfo.InvariantCulture, _vsGlobalExtensionFolderFormat, (int)version));
                }
                else if ((int) version == 15)
                {
                    ExtensionDirectories.Add(FindVs15GlobalExtensionFolder());
                }

                ExtensionDirectories = ExtensionDirectories
                    .Where(item => item != null)
                    .ToArray();
            }

            Extensions = ReadExtensions(ExtensionDirectories).ToArray();
        }

        public string Name { get; }

        public IList<string> ExtensionDirectories { get; }

        public IList<IGrouping<string, Extension>> Extensions { get; } // = new IGrouping<string, Extension>[0];

        private static IEnumerable<IGrouping<string, Extension>> ReadExtensions(IEnumerable<string> pathNames)
        {
            var paths = pathNames.Select(path => new DirectoryInfo(path));

            var extensions = paths
                .Where(dir => dir.Exists)
                .SelectMany(dir => dir.GetDirectories("*.*", SearchOption.AllDirectories))
                .Where(dir => !dir.Attributes.HasFlag(FileAttributes.ReparsePoint)) // skip symlinks and junctions)
                .SelectMany(dir => _manifestNames.Select(name => Path.Combine(dir.FullName, name)).Select(manifest => ReadExtensionFile(manifest, dir)))
                .Where(ext => ext != null)
                .ToArray();

            var extensionsById = extensions
                .OrderBy(x => x.Name)
                .GroupBy(x => x.Id)
                .ToArray();

            var toDelete = extensionsById
                .SelectMany(group => group.OrderByDescending(x => x.Version).ThenByDescending(x => x.CreationTime).Skip(1));

            foreach (var extension in toDelete)
            {
                extension.IsDuplicate = true;
            }

            return extensionsById;
        }

        private static Extension ReadExtensionFile(string manifest, FileSystemInfo dir)
        {
            if (!File.Exists(manifest))
                return null;

            using (var reader = XmlReader.Create(manifest, new XmlReaderSettings { IgnoreComments = true }))
            {
                try
                {
                    if (_vsixSerializer.CanDeserialize(reader))
                    {
                        var vsix = (Vsix)_vsixSerializer.Deserialize(reader);
                        return new Extension
                        {
                            Id = vsix.Identifier.Id,
                            Name = vsix.Identifier.Name,
                            Version = new Version(vsix.Identifier.Version),
                            Path = dir.FullName,
                            CreationTime = dir.CreationTime
                        };
                    }

                    if (_packageSerializer.CanDeserialize(reader))
                    {
                        var package = (PackageManifest)_packageSerializer.Deserialize(reader);
                        return new Extension
                        {
                            Id = package.Metadata.Identity.Id,
                            Name = package.Metadata.DisplayName,
                            Version = new Version(package.Metadata.Identity.Version),
                            Path = dir.FullName,
                            CreationTime = dir.CreationTime
                        };
                    }
                }
                catch (XmlException)
                {
                    // invalid manifest, ignore...
                }
            }

            return null;
        }

        private static string FindVs15GlobalExtensionFolder()
        {
            var root = new DirectoryInfo(Path.Combine(_programFilesFolder, "Microsoft Visual Studio", "2017"));
            var folder = root.EnumerateDirectories("Extensions", SearchOption.AllDirectories)
                .FirstOrDefault(dir => string.Equals("IDE", dir.Parent?.Name, StringComparison.OrdinalIgnoreCase) 
                    && string.Equals("Common7", dir.Parent?.Parent?.Name, StringComparison.OrdinalIgnoreCase));

            return folder?.FullName;
        }
    }

    public class Extension
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public string Path { get; set; }

        public Version Version { get; set; }

        public DateTime CreationTime { get; set; }

        public bool IsDuplicate { get; set; }
    }

    [XmlRoot(ElementName = "Metadata", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2011")]
    public class Metadata
    {
        [XmlElement(ElementName = "Identity", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2011")]
        public Identity Identity { get; set; }

        [XmlElement(ElementName = "DisplayName", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2011")]
        public string DisplayName { get; set; }
    }

    [XmlRoot(ElementName = "Identity", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2011")]
    public class Identity
    {
        [XmlAttribute(AttributeName = "Id")]
        public string Id { get; set; }

        [XmlAttribute(AttributeName = "Version")]
        public string Version { get; set; }
    }

    [XmlRoot(ElementName = "PackageManifest", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2011")]
    public class PackageManifest
    {
        [XmlElement(ElementName = "Metadata", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2011")]
        public Metadata Metadata { get; set; }
    }

    [XmlRoot(ElementName = "Vsix", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2010")]
    public class Vsix
    {
        [XmlElement(ElementName = "Identifier", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2010")]
        public Identifier Identifier { get; set; }
    }

    [XmlRoot(ElementName = "Identifier", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2010")]
    public class Identifier
    {
        [XmlElement(ElementName = "Name", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2010")]
        public string Name { get; set; }

        [XmlElement(ElementName = "Version", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2010")]
        public string Version { get; set; }

        [XmlAttribute(AttributeName = "Id")]
        public string Id { get; set; }
    }
}
