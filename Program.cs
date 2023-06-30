using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using Newtonsoft.Json;
using System.Xml.Linq;
using System.Xml;
using static PPT_Compare.StatData;
using System.Collections;
using DocumentFormat.OpenXml.Bibliography;
using System.Text.RegularExpressions;
using System.Configuration;
using DocumentFormat.OpenXml;
using System.IO.Compression;

namespace PPT_Compare
{
    public class StatData
    {
        public enum SlideState
        {
            New,
            Updated,
            Deleted,
            NotModified
        }

        public string FileName { get; set; }
        public SlideState[] Slides { get; set; }
    }

    class Program
    {
        private static List<StatData> stat = new List<StatData>();
        static void Main(string[] arg)
        {

            if (arg.Length != 2)
            {
                PrintHelp();
                return;
            }

            if (File.Exists(arg[0]) && File.Exists(arg[1]))
            {
                ProcessFile(arg[0], arg[1]);
            }
            else if (Directory.Exists(arg[0]) && Directory.Exists(arg[1]))
            {
                var d = new DirectoryInfo(arg[1]);
                var files = d.GetFiles("*.pptx");
                foreach (var f in files)
                {
                    var oldfile = Path.Combine(arg[0], f.Name);
                    if (File.Exists(oldfile))
                    {
                        try
                        {
                            ProcessFile(oldfile, f.FullName);

                        }
                        catch (OpenXmlPackageException ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Cyan;
                            Console.WriteLine($"Error during open {f.Name}");
                            Console.WriteLine($"Press key to continue");
                            Console.Read();
                            Console.ForegroundColor = ConsoleColor.White;
                        }


                    }
                    else
                        Console.WriteLine($"File {Path.GetFileName(arg[0])} not found in modified folder!");
                }

                PrintStat();

                Console.WriteLine("PPTX text exported to JSON files in the target folders. You can use text compare tool for find out change details.");
                Console.WriteLine("Press Any Key to finish");
                Console.Read();
            }
            else
            {
                Console.WriteLine("Provided Paths is not exists");
            }

        }


        private static void PrintStat()
        {
            int totalOrg = 0;
            int totalUpdated =0;
            Console.WriteLine("Slides map...................................");
            Console.WriteLine("legend: White - unchanged, Yellow - updated, Green - added, Red - deleted");
            foreach (var module in stat)
            {
                var org = module.Slides.Skip(1).Where(x => (x == SlideState.Deleted || x == SlideState.NotModified || x == SlideState.Updated)).Count();
                totalOrg += org;
                var upd = module.Slides.Skip(1).Where(x => (x == SlideState.New || x == SlideState.NotModified || x == SlideState.Updated)).Count();
                totalUpdated += upd;
                Console.Write($"{module.FileName}:({org}=>{upd})\t\t");
                PrintSlideStat(module.Slides);
                Console.WriteLine();
            }
            Console.BackgroundColor = ConsoleColor.White;
            Console.ForegroundColor = ConsoleColor.Black;
            Console.WriteLine($"Total changes ({totalOrg}=>{totalUpdated})");
            Console.BackgroundColor = ConsoleColor.Black;
            Console.ForegroundColor = ConsoleColor.White;
        }

        private static void PrintSlideStat(SlideState[] slides)
        {
         slides.Skip(1).ToList().ForEach(slide =>
         {
             switch (slide)
             {
                 case SlideState.Deleted:
                     Console.ForegroundColor = ConsoleColor.Red;
                     Console.Write("X");
                     break;
                 case SlideState.New:
                     Console.ForegroundColor = ConsoleColor.Green;
                     Console.Write("X");
                     break;
                 case SlideState.Updated:
                     Console.ForegroundColor = ConsoleColor.Yellow;
                     Console.Write("X");
                     break;
                 default:
                     Console.ForegroundColor = ConsoleColor.White;
                     Console.Write("X");
                     break;
             }

         });

        Console.ForegroundColor = ConsoleColor.White;
        }


        private static void PrintHelp()
        {
            Console.WriteLine(@"
___________________________________  ___               
\______   \______   \__    ___/\   \/  /               
 |     ___/|     ___/ |    |    \     /                
 |    |    |    |     |    |    /     \                
 |____|    |____|     |____|   /___/\  \               
                                                      
  ____  ____   _____ ___________ _______   ___________ 
_/ ___\/  _ \ /     \\____ \__  \\_  __ \_/ __ \_  __ \
\  \__(  <_> )  Y Y  \  |_> > __ \|  | \/\  ___/|  | \/
 \___  >____/|__|_|  /   __(____  /__|    \___  >__|   
     \/            \/|__|       \/            \/       
                
by Alex Ivanov - 2019 v2

Provide path to the original(old) and modified(new) version.
Comparing will be done ONLY for files with the SAME name.

example for folders: 
\tpptx-comparer.exe  c:\temp\old  c:\temp\new

example for files: 
\tpptx-comparer.exe  c:\temp\old\demo.pptx  c:\temp\new\demo.pptx
");

          
        }

        static void ProcessFile(string OrgfileNamePath, string ModfileNamePath)
        {
            StatData sData = new StatData() { FileName = Path.GetFileName(OrgfileNamePath) };
            stat.Add(sData);

            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine($"File: {Path.GetFileName(OrgfileNamePath)}");
            Console.ForegroundColor = ConsoleColor.White;

            var modifiedText = GetTextOnSlides(ModfileNamePath);
            var originalText = GetTextOnSlides(OrgfileNamePath);

            var modifiedNote = GetNoteOnSlides(ModfileNamePath);
            var originalNote = GetNoteOnSlides(OrgfileNamePath);

            var modifiedImg = GetImagesOnSlides(ModfileNamePath);
            var originalImg = GetImagesOnSlides(OrgfileNamePath);

            File.WriteAllText(Path.Combine(Path.GetDirectoryName(ModfileNamePath), Path.GetFileNameWithoutExtension(ModfileNamePath) +  ".json"), JsonConvert.SerializeObject(new { text = modifiedText, images = modifiedImg, notes = modifiedNote }, Newtonsoft.Json.Formatting.Indented));
            File.WriteAllText(Path.Combine(Path.GetDirectoryName(OrgfileNamePath), Path.GetFileNameWithoutExtension(OrgfileNamePath) + ".json"), JsonConvert.SerializeObject(new { text = originalText, images = originalImg, notes = originalNote }, Newtonsoft.Json.Formatting.Indented));

            sData.Slides = new SlideState[(modifiedText.Keys.Count > originalText.Keys.Count ? modifiedText.Keys.Count : originalText.Keys.Count )+ 1];

            // Open the presentation as read-only.
            int i = 0;
            foreach (uint sID in modifiedText.Keys)
            {
                i++;
                if (originalText.Keys.Contains(sID))
                {
                    var sMod = modifiedText[sID];
                    var sOrg = originalText[sID];

                    var nMod = modifiedNote[sID];
                    var nOrg = originalNote[sID];

                    var bModified = false;
                    var iModified = false;
                    var nModified = false;

                    for (int l = 0; l < sMod.Length; l++)
                    {
                        if (sOrg.Length < (l + 1) || sMod[l].ToLower() != sOrg[l].ToLower())
                        {
                            bModified = true;
                        }
                    }

                    for (int l = 0; l < nMod.Length; l++)
                    {
                        if (nOrg.Length < (l + 1))
                        {
                            nModified = true;
                            break;
                        }
                        // if the data contains exact data format it seems to come from the edit data. lets ignore that. 
                        if (nMod[l].ToLower() != nOrg[l].ToLower())
                        {
                            nModified = true; 
                        }
                    }

                    if (originalImg.Keys.Contains(sID))
                    {
                        var sIMod = modifiedImg[sID];
                        var sIOrg = originalImg[sID];                        
                        for (int l = 0; l < sIMod.Length; l++)
                        {
                            if (sIOrg.Length < (l + 1) || sIMod[l].GetHashCode() != sIOrg[l].GetHashCode())
                            {
                                iModified = true;
                            }
                        }
                        if (!iModified) iModified = sIOrg.Length != sIMod.Length;

                    }

                    if (bModified || iModified || nModified)
                    {
                        sData.Slides[i] = StatData.SlideState.Updated;
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        List<string> arr = new List<string>();
                        arr.Add($"\tSlide #{i}");
                        if (bModified)
                            arr.Add("text modified");
                        if (iModified)
                            arr.Add("pictures modified");
                        if (nModified)
                            arr.Add("notes modified");
                        Console.WriteLine(string.Join(", ", arr));
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                    else
                    {
                        sData.Slides[i] = StatData.SlideState.NotModified;
                    }

                }
                else
                {
                    sData.Slides[i] = StatData.SlideState.New;

                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"\tSlide #{i} Added");
                    Console.ForegroundColor = ConsoleColor.White;
                }
            }

            i = 0;
            foreach (uint sID in originalText.Keys)
            {
                i++;
                if (!modifiedText.Keys.Contains(sID))
                {
                    sData.Slides[i] = StatData.SlideState.Deleted;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"\tSlide #{i} Deleted");
                    Console.ForegroundColor = ConsoleColor.White;
                }

            }
        }

         private static Dictionary<uint, string[]> GetNoteOnSlides(string presentationFile)
        {
            var slides = new Dictionary<uint, string[]>();
            try
            {
                using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
                {

                    PresentationPart presentationPart = presentationDocument.PresentationPart;

                    // Verify that the presentation part and presentation exist.
                    if (presentationPart != null && presentationPart.Presentation != null)
                    {
                        // Get the Presentation object from the presentation part.
                        Presentation presentation = presentationPart.Presentation;

                        // Verify that the slide ID list exists.
                        if (presentation.SlideIdList != null)
                        {
                            // Get the collection of slide IDs from the slide ID list.
                            var slideIds = presentation.SlideIdList.ChildElements;

                            foreach (SlideId sld in slideIds)
                            {
                                string slidePartRelationshipId = sld.RelationshipId;

                                // Get the specified slide part from the relationship ID.
                                SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slidePartRelationshipId);

                                // Pass the slide part to the next method, and
                                // then return the array of strings that method
                                // returns to the previous method.
                                slides.Add(sld.Id, GetSlideNotes(slidePart) ?? new string[] { });
                            }
                        }
                    }
                }
                return slides;
            }
            catch (OpenXmlPackageException e)
            {
                if (e.ToString().Contains("Invalid Hyperlink"))
                {
                    using (FileStream fs = new FileStream(presentationFile, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        UriFixer.FixInvalidUri(fs, brokenUri => FixUri(brokenUri));
                    }
                    return GetTextOnSlides(presentationFile);
                }
                throw e;
            }
        }

        public static Dictionary<uint, string[]> GetImagesOnSlides(string presentationFile)
        {
            var slides = new Dictionary<uint, string[]>();
            try
            {
                using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
                {

                    PresentationPart presentationPart = presentationDocument.PresentationPart;

                    // Verify that the presentation part and presentation exist.
                    if (presentationPart != null && presentationPart.Presentation != null)
                    {
                        // Get the Presentation object from the presentation part.
                        Presentation presentation = presentationPart.Presentation;

                        // Verify that the slide ID list exists.
                        if (presentation.SlideIdList != null)
                        {
                            // Get the collection of slide IDs from the slide ID list.
                            var slideIds = presentation.SlideIdList.ChildElements;

                            foreach (SlideId sld in slideIds)
                            {
                                string slidePartRelationshipId = sld.RelationshipId;

                                // Get the specified slide part from the relationship ID.
                                SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slidePartRelationshipId);

                                // Pass the slide part to the next method, and
                                // then return the array of strings that method
                                // returns to the previous method.
                                slides.Add(sld.Id, GetAllImagesInSlide(slidePart) ?? new string[] { });
                            }
                        }
                    }
                }
                return slides;
            }
            catch (OpenXmlPackageException e)
            {
                if (e.ToString().Contains("Invalid Hyperlink"))
                {
                    using (FileStream fs = new FileStream(presentationFile, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        UriFixer.FixInvalidUri(fs, brokenUri => FixUri(brokenUri));
                    }
                    return GetTextOnSlides(presentationFile);
                }
                throw e;
            }

        }
        public static Dictionary<uint,string[]> GetTextOnSlides(string presentationFile)
        {
            var slides = new Dictionary<uint, string[]>();
            try
            {
                using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
                {

                    PresentationPart presentationPart = presentationDocument.PresentationPart;

                    // Verify that the presentation part and presentation exist.
                    if (presentationPart != null && presentationPart.Presentation != null)
                    {
                        // Get the Presentation object from the presentation part.
                        Presentation presentation = presentationPart.Presentation;

                        // Verify that the slide ID list exists.
                        if (presentation.SlideIdList != null)
                        {
                            // Get the collection of slide IDs from the slide ID list.
                            var slideIds = presentation.SlideIdList.ChildElements;

                            foreach (SlideId sld in slideIds)
                            {
                                string slidePartRelationshipId = sld.RelationshipId;

                                // Get the specified slide part from the relationship ID.
                                SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slidePartRelationshipId);

                                // Pass the slide part to the next method, and
                                // then return the array of strings that method
                                // returns to the previous method.
                                slides.Add(sld.Id, GetAllTextInSlide(slidePart) ?? new string[] { });
                            }
                        }
                    }
                }
                return slides;
            }
            catch (OpenXmlPackageException e)
            {
                if (e.ToString().Contains("Invalid Hyperlink"))
                {
                    using (FileStream fs = new FileStream(presentationFile, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        UriFixer.FixInvalidUri(fs, brokenUri => FixUri(brokenUri));
                    }
                    return GetTextOnSlides(presentationFile);
                }
                throw e;
            }

        }

        private static Uri FixUri(string brokenUri)
        {
            return new Uri("http://microsoft.com/");
        }

        public static string[] GetAllTextInSlide(PresentationDocument presentationDocument, int slideIndex)
        {
       
            // Get the presentation part of the presentation document.
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // Verify that the presentation part and presentation exist.
            if (presentationPart != null && presentationPart.Presentation != null)
            {
                // Get the Presentation object from the presentation part.
                Presentation presentation = presentationPart.Presentation;

                // Verify that the slide ID list exists.
                if (presentation.SlideIdList != null)
                {
                    // Get the collection of slide IDs from the slide ID list.
                    var slideIds = presentation.SlideIdList.ChildElements;

                    // If the slide ID is in range...
                    if (slideIndex < slideIds.Count)
                    {
                        // Get the relationship ID of the slide.
                        string slidePartRelationshipId = (slideIds[slideIndex] as SlideId).RelationshipId;

                        // Get the specified slide part from the relationship ID.
                        SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slidePartRelationshipId);

                        // Pass the slide part to the next method, and
                        // then return the array of strings that method
                        // returns to the previous method.
                        return GetAllTextInSlide(slidePart);
                    }
                }
            }
            // Else, return null.
            return null;
        }

        public static string[] GetAllTextInSlide(SlidePart slidePart)
        {
            // Verify that the slide part exists.
            if (slidePart == null)
            {
                throw new ArgumentNullException("slidePart");
            }

            // Create a new linked list of strings.
            LinkedList<string> texts = new LinkedList<string>();

            // If the slide exists...
            if (slidePart.Slide != null)
            {
                // Iterate through all the paragraphs in the slide.
                foreach (var paragraph in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
                {
                    // Create a new string builder.                    
                    StringBuilder paragraphText = new StringBuilder();

                    // Iterate through the lines of the paragraph.
                    foreach (var text in paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                    {
                        // Append each line to the previous lines.
                        paragraphText.Append(text.Text);
                    }

                    if (paragraphText.Length > 0)
                    {
                        // Add each paragraph to the linked list.
                        texts.AddLast(paragraphText.ToString().Trim().ToLower());
                    }
                }
            }

            if (texts.Count > 0)
            {
                // Return an array of strings.
                return texts.ToArray();
            }
            else
            {
                return null;
            }
        }

        public static string[] GetSlideNotes(SlidePart slidePart)
        {
            LinkedList<string> texts = new LinkedList<string>();
            // Verify that the slide part exists.
            if (slidePart == null)
            {
                throw new ArgumentNullException("slidePart");
            }

            if (slidePart.NotesSlidePart != null && slidePart.NotesSlidePart.NotesSlide != null)
            {
                foreach (Shape sp in slidePart.NotesSlidePart.NotesSlide.CommonSlideData.ShapeTree.Descendants<Shape>())
                {
                    if (sp.NonVisualShapeProperties.NonVisualDrawingProperties.Name.Value.StartsWith("Notes Placeholder"))
                    {

                        foreach (var paragraph in sp.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
                        {
                            // Create a new string builder.      
                            StringBuilder paragraphText = new StringBuilder();

                            // Iterate through the lines of the paragraph.
                            foreach (var txt in paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                            {
                                // Append each line to the previous lines.
                                paragraphText.Append(txt.Text);
                            }

                            if (paragraphText.Length > 0)
                            {
                                // Add each paragraph to the linked list.
                                texts.AddLast(paragraphText.ToString().Trim().ToLower());
                            }
                        }
                    }
                }
            }

            if (texts.Count > 0)
            {
                // Return an array of strings.
                return texts.ToArray();
            }
            else
            {
                return null;
            }
        }

        public static string[] GetAllImagesInSlide(SlidePart slidePart)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            // Verify that the slide part exists.
            if (slidePart == null)
            {
                throw new ArgumentNullException("slidePart");
            }
            try
            {
                // If the slide exists...
                if (slidePart.Slide != null)
                {
                    // Iterate through all the paragraphs in the slide.
                    foreach (var img in slidePart.Slide.Descendants<Picture>())
                    {
                        var theShape = img.NonVisualPictureProperties.Descendants<NonVisualDrawingProperties>().FirstOrDefault();

                        string name = img != null && theShape.Name != null && theShape.Name.HasValue ? theShape.Name.Value : null;
                        string desc = img != null && theShape.Description != null && theShape.Description.HasValue ? theShape.Description : null;

                        if (img.ShapeProperties.Transform2D != null)
                        {
                            // Console.WriteLine($"{name} = {img.ShapeProperties.Transform2D.Extents.Cx.Value}:{ img.ShapeProperties.Transform2D.Extents.Cy.Value} {img.ShapeProperties.Transform2D.Offset.X.Value}:{ img.ShapeProperties.Transform2D.Offset.Y.Value}");
                            result[name] = $"{ img.ShapeProperties.Transform2D.Extents.Cx.Value}:{img.ShapeProperties.Transform2D.Extents.Cy.Value} {img.ShapeProperties.Transform2D.Offset.X.Value}:{img.ShapeProperties.Transform2D.Offset.Y.Value}";
                        }

                        var gshapes = img.Parent.Descendants<GroupShape>();

                        if (gshapes != null)
                            foreach (var shape in gshapes)
                            {
                                if (shape.GroupShapeProperties.TransformGroup.Extents != null) //&& !result.ContainsKey(name))
                                {
                                    result[shape.NonVisualGroupShapeProperties.NonVisualDrawingProperties.Name] = $"{shape.GroupShapeProperties.TransformGroup.Extents.Cx.Value}:{ shape.GroupShapeProperties.TransformGroup.Extents.Cy.Value} {shape.GroupShapeProperties.TransformGroup.Offset.X.Value}:{ shape.GroupShapeProperties.TransformGroup.Offset.Y.Value}";
                                }
                            }
                    }

                }

                return result.Keys.Select(x => $"{x}:{result[x]}").ToArray();
            }
            catch
            {
                return null;
            }

        }


        public static class UriFixer
        {
            public static void FixInvalidUri(Stream fs, Func<string, Uri> invalidUriHandler)
            {
                XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
                using (ZipArchive za = new ZipArchive(fs, ZipArchiveMode.Update))
                {
                    foreach (var entry in za.Entries.ToList())
                    {
                        if (!entry.Name.EndsWith(".rels"))
                            continue;
                        bool replaceEntry = false;
                        XDocument entryXDoc = null;
                        using (var entryStream = entry.Open())
                        {
                            try
                            {
                                entryXDoc = XDocument.Load(entryStream);
                                if (entryXDoc.Root != null && entryXDoc.Root.Name.Namespace == relNs)
                                {
                                    var urisToCheck = entryXDoc
                                        .Descendants(relNs + "Relationship")
                                        .Where(r => r.Attribute("TargetMode") != null && (string)r.Attribute("TargetMode") == "External");
                                    foreach (var rel in urisToCheck)
                                    {
                                        var target = (string)rel.Attribute("Target");
                                        if (target != null)
                                        {
                                            try
                                            {
                                                Uri uri = new Uri(target);
                                            }
                                            catch (UriFormatException)
                                            {
                                                Uri newUri = invalidUriHandler(target);
                                                rel.Attribute("Target").Value = newUri.ToString();
                                                replaceEntry = true;
                                            }
                                        }
                                    }
                                }
                            }
                            catch (XmlException)
                            {
                                continue;
                            }
                        }
                        if (replaceEntry)
                        {
                            var fullName = entry.FullName;
                            entry.Delete();
                            var newEntry = za.CreateEntry(fullName);
                            using (StreamWriter writer = new StreamWriter(newEntry.Open()))
                            using (XmlWriter xmlWriter = XmlWriter.Create(writer))
                            {
                                entryXDoc.WriteTo(xmlWriter);
                            }
                        }
                    }
                }
            }
        }

    }
}
