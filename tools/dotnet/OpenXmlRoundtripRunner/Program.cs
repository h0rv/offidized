using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

internal static class Program
{
    private static int Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.Error.WriteLine("usage: OpenXmlRoundtripRunner <input> <output>");
            return 2;
        }

        var inputPath = Path.GetFullPath(args[0]);
        var outputPath = Path.GetFullPath(args[1]);
        var extension = Path.GetExtension(inputPath).ToLowerInvariant();

        if (!File.Exists(inputPath))
        {
            Console.Error.WriteLine($"input file does not exist: {inputPath}");
            return 2;
        }

        try
        {
            var outputDirectory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrWhiteSpace(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            File.Copy(inputPath, outputPath, overwrite: true);

            switch (extension)
            {
                case ".docx":
                    using (var document = WordprocessingDocument.Open(outputPath, true))
                    {
                        SaveAllKnownPartRoots(document);
                    }
                    break;
                case ".xlsx":
                    using (var document = SpreadsheetDocument.Open(outputPath, true))
                    {
                        SaveAllKnownPartRoots(document);
                    }
                    break;
                case ".pptx":
                    using (var document = PresentationDocument.Open(outputPath, true))
                    {
                        SaveAllKnownPartRoots(document);
                    }
                    break;
                default:
                    Console.Error.WriteLine($"unsupported extension: {extension}");
                    return 2;
            }

            return 0;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine(exception.ToString());
            return 1;
        }
    }

    private static void SaveAllKnownPartRoots(OpenXmlPackage package)
    {
        var queue = new Queue<OpenXmlPartContainer>();
        var visitedParts = new HashSet<OpenXmlPart>();
        queue.Enqueue(package);

        while (queue.Count > 0)
        {
            var container = queue.Dequeue();
            foreach (var idPartPair in container.Parts)
            {
                var part = idPartPair.OpenXmlPart;
                part.RootElement?.Save();

                if (visitedParts.Add(part))
                {
                    queue.Enqueue(part);
                }
            }
        }
    }
}
