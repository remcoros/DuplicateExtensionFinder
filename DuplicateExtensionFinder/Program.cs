namespace DuplicateExtensionFinder
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Xml;
    using System.Xml.Serialization;

    class Program
    {
        static void Main(string[] args)
        {
            var localAppDataDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var programFilesDir = Environment.GetFolderPath(Environment.Is64BitOperatingSystem ? Environment.SpecialFolder.ProgramFilesX86 : Environment.SpecialFolder.ProgramFiles);

            var paths = new[]
            {
                Path.Combine(localAppDataDir, @"Microsoft\VisualStudio\14.0\Extensions"),
                Path.Combine(programFilesDir, @"Microsoft Visual Studio 14.0\Common7\IDE\Extensions")
            };

            var onlyDupes = args.Any(x => string.Equals(x, "-dupes", StringComparison.OrdinalIgnoreCase));
            var doDelete = args.Any(x => string.Equals(x, "-delete", StringComparison.OrdinalIgnoreCase));

            var extensions = new List<Extension>();

            foreach (var path in paths)
            {
                var extensionDir = new DirectoryInfo(path);

                var vsixSerializer = new XmlSerializer(typeof(Vsix));
                var packageSerializer = new XmlSerializer(typeof(PackageManifest));

                var extDirs = extensionDir.GetDirectories("*.*", SearchOption.AllDirectories);

                foreach (var dir in extDirs)
                {
                    // skip symlinks and junctions
                    if (dir.Attributes.HasFlag(FileAttributes.ReparsePoint))
                    {
                        continue;
                    }

                    var manifest = Path.Combine(dir.FullName, "extension.vsixmanifest");
                    if (File.Exists(manifest))
                    {
                        using (var rdr = XmlReader.Create(manifest, new XmlReaderSettings { IgnoreComments = true }))
                        {
                            try
                            {
                                if (vsixSerializer.CanDeserialize(rdr))
                                {
                                    var vsix = (Vsix)vsixSerializer.Deserialize(rdr);
                                    extensions.Add(new Extension()
                                    {
                                        Id = vsix.Identifier.Id,
                                        Name = vsix.Identifier.Name,
                                        Version = new Version(vsix.Identifier.Version),
                                        Path = dir.FullName,
                                        CreationTime = dir.CreationTime
                                    });
                                }
                                else if (packageSerializer.CanDeserialize(rdr))
                                {
                                    var package = (PackageManifest)packageSerializer.Deserialize(rdr);
                                    extensions.Add(new Extension()
                                    {
                                        Id = package.Metadata.Identity.Id,
                                        Name = package.Metadata.DisplayName,
                                        Version = new Version(package.Metadata.Identity.Version),
                                        Path = dir.FullName,
                                        CreationTime = dir.CreationTime
                                    });
                                }
                            }
                            catch (XmlException)
                            {
                                // invalid manifest, ignore...
                            }
                        }
                    }
                }
            }

            var grouped = extensions.OrderBy(x => x.Name).GroupBy(x => x.Id).ToList();
            var toDelete = grouped.Where(x => x.Count() > 1).SelectMany(@group => @group.OrderByDescending(x => x.Version).ThenByDescending(x => x.CreationTime).Skip(1)).ToList();

            foreach (var group in grouped.Where(x => x.Count() > (onlyDupes ? 1 : 0)))
            {
                Console.WriteLine("{0}", @group.First().Name);
                foreach (var vsix in group.OrderBy(x => x.Version))
                {
                    var isDuplicate = toDelete.Contains(vsix);
                    Console.WriteLine(" - {0} [{2}] ({1})", vsix.Version, vsix.Path, isDuplicate ? "DELETE" : "KEEP");

                    if (isDuplicate && doDelete)
                    {
                        try
                        {
                            Directory.Delete(vsix.Path, true);
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
                }

                Console.WriteLine();
            }

            if (toDelete.Count == 0)
            {
                Console.WriteLine("No duplicates found.");
            }

            if (!doDelete)
            {
                Console.WriteLine("Specify '-delete' to delete old extensions from disk.");
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }
    }

    public class Extension
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public string Path { get; set; }

        public Version Version { get; set; }

        public DateTime CreationTime { get; set; }
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
