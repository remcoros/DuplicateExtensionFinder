using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DuplicateExtensionFinder
{
    using System.IO;
    using System.Xml;
    using System.Xml.Serialization;

    class Program
    {
        static void Main(string[] args)
        {
            bool onlyDupes = false;
            var paths = new[]
            {
                Path.Combine(Environment.GetEnvironmentVariable("LocalAppData"), @"Microsoft\VisualStudio\14.0\Extensions"),
                @"C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\Extensions"
            };
            if (args.Any(x => string.Equals(x, "-dupes", StringComparison.OrdinalIgnoreCase)))
            {
                onlyDupes = true;
            }

            var extensions = new List<Extension>();
            foreach (var path in paths)
            {
                var extensionDir = new DirectoryInfo(path);

                var vsixSerializer = new XmlSerializer(typeof(Vsix));
                var packageSerializer = new XmlSerializer(typeof(PackageManifest));

                var extDirs = extensionDir.GetDirectories();
                foreach (var dir in extDirs)
                {
                    // skip symlinks and junctions
                    if (dir.Attributes.HasFlag(FileAttributes.ReparsePoint))
                    {
                        continue;
                    }

                    string manifest = Path.Combine(dir.FullName, "extension.vsixmanifest");
                    if (File.Exists(manifest))
                    {
                        using (var file = File.OpenRead(manifest))
                        using (var rdr = new XmlTextReader(file))
                        {
                            if (vsixSerializer.CanDeserialize(rdr))
                            {
                                var vsix = (Vsix)vsixSerializer.Deserialize(rdr);
                                extensions.Add(new Extension()
                                {
                                    Id = vsix.Identifier.Id,
                                    Name = vsix.Identifier.Name,
                                    Version = new Version(vsix.Identifier.Version),
                                    Path = dir.FullName
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
                                    Path = dir.FullName
                                });
                            }
                        }
                    }
                }
            }

            var grouped = extensions.OrderBy(x => x.Name).GroupBy(x => x.Id);

            foreach (var group in grouped.Where(x => x.Count() > (onlyDupes ? 1 : 0)))
            {
                Console.WriteLine("{0}", @group.First().Name);
                foreach (var vsix in group.OrderBy(x => x.Version))
                {
                    Console.WriteLine(" - {0} ({1})", vsix.Version, vsix.Path);
                }

                Console.WriteLine();
            }
        }
    }

    public class Extension
    {
        public string Id { get; set; }
        public string Name { get; set; }

        public string Path { get; set; }

        public Version Version { get; set; }
    }

    [XmlRoot(ElementName = "Metadata", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2011")]
    public class Metadata
    {
        [XmlElement(ElementName = "Identity", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2011")]
        public Identity Identity { get; set; }

        [XmlElement(ElementName = "DisplayName", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2011")]
        public string DisplayName { get; set; }

        [XmlElement(ElementName = "Description", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2011")]
        public string Description { get; set; }
    }

    [XmlRoot(ElementName = "Identity", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2011")]
    public class Identity
    {
        [XmlAttribute(AttributeName = "Id")]
        public string Id { get; set; }

        [XmlAttribute(AttributeName = "Version")]
        public string Version { get; set; }

        [XmlAttribute(AttributeName = "Language")]
        public string Language { get; set; }

        [XmlAttribute(AttributeName = "Publisher")]
        public string Publisher { get; set; }
    }

    [XmlRoot(ElementName = "PackageManifest", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2011")]
    public class PackageManifest
    {
        [XmlElement(ElementName = "Metadata", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2011")]
        public Metadata Metadata { get; set; }

        [XmlIgnore]
        public string Path { get; set; }
    }

    [XmlRoot(ElementName = "Vsix", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2010")]
    public class Vsix
    {
        [XmlElement(ElementName = "Identifier", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2010")]
        public Identifier Identifier { get; set; }

        [XmlIgnore]
        public string Path { get; set; }
    }

    [XmlRoot(ElementName = "Identifier", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2010")]
    public class Identifier
    {
        [XmlElement(ElementName = "Name", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2010")]
        public string Name { get; set; }

        [XmlElement(ElementName = "Author", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2010")]
        public string Author { get; set; }

        [XmlElement(ElementName = "Version", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2010")]
        public string Version { get; set; }

        [XmlElement(ElementName = "Description", Namespace = "http://schemas.microsoft.com/developer/vsx-schema/2010")]
        public string Description { get; set; }

        [XmlAttribute(AttributeName = "Id")]
        public string Id { get; set; }
    }
}
