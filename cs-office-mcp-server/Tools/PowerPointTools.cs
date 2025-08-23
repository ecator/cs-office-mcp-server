using ModelContextProtocol;
using ModelContextProtocol.Server;
using System;
using System.Buffers;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace OfficeServer.Tools;

[McpServerToolType]
public static class PowerPointTools

{

    [McpServerTool(Name = "powerpoint_get_slide_count"), Description("Get all the number of the slides of the specified PowerPoint file.")]
    public static string GetSlideCount([Description("The full path of the PowerPoint file.")] string fullName
        , [Description("The password of the PowerPoint file, if there is one.")] string? password = null)
    {
        var data = new StringBuilder();
        var count = 0;
        data.AppendLine();
        using (var session = new PowerPointSession())
        {
            var pr = session.OpenPresentation(fullName, true, password);

            count = session.GetSlideCount(pr);
        }
        data.Insert(0, $"Total `{count}` slides in the PowerPoint file `{fullName}`.");
        return data.ToString();
    }

    [McpServerTool(Name = "powerpoint_read"), Description("Get the text content of the specified PowerPoint file.")]
    public static Dictionary<string, Dictionary<string, object>> Read([Description("The full path of the PowerPoint file.")] string fullName
        , [Description("The starting slide number to read.")] int fromSlide = 1
        , [Description("The end slide number to read. If it's empty, then read up to the last slide.")] int? toSlide = null
        , [Description("The password of the PowerPoint file, if there is one.")] string? password = null)
    {
        var data = new Dictionary<string, Dictionary<string, object>>();
        var slideCount = 0;
        using (var session = new PowerPointSession())
        {
            var pr = session.OpenPresentation(fullName, true, password);
            slideCount = session.GetSlideCount(pr);
            if (toSlide.HasValue && toSlide > slideCount)
            {
                throw new McpException($"The end slide number {toSlide} cannot be greater than the total page count {slideCount}.");
            }
            if (!toSlide.HasValue)
            {
                toSlide = slideCount;
            }
            var slides = pr.Slides;
            session.RegisterComObject(slides);
            for(var i = fromSlide; i <= toSlide; i++)
            {
                var slideName = $"Slide{i}";
                var slide = slides[i];
                session.RegisterComObject(slide);
                var shapes = slide.Shapes;
                session.RegisterComObject(shapes);
                var shapesText = session.GetShapesText(shapes);
                data[slideName] = new Dictionary<string, object>();
                data[slideName]["page"] = shapesText;
                var notesPage = slide.NotesPage;
                session.RegisterComObject(notesPage);
                var notesShapes = notesPage.Shapes;
                session.RegisterComObject(notesShapes);
                var notesShapesText = session.GetShapesText(notesShapes);
                data[slideName]["notes"] = notesShapesText;
            }

        }
        return data;
    }

    [McpServerTool(Name = "powerpoint_find"), Description("Find value from PowerPoint files.")]
    public static string Find([Description("The list of full path of PowerPoint files that need to be searched for.")] string[] fullNameList
    , [Description(@"The value to be searched for.")] string searchValue
    , [Description("Match against any part of part of a larger word when true. Match against the entire words of the search text when false.")] bool matchPart = true
    , [Description("Ignoring lower case and upper case differences when true. Case insensitive when false.")] bool ignoreCase = true
    , [Description("The password of the PowerPoint files, if there is one and all are the same.")] string? password = null)
    {
        var data = new StringBuilder();
        var foundData = new StringBuilder();
        var line = new string[4];
        var totalCount = 0;
        var count = 0;
        if (fullNameList == null || fullNameList.Length == 0)
        {
            throw new McpException("The full path list of the PowerPoint file cannot be empty or null.");
        }
        data.AppendLine();
        data.AppendLine();
        using (var session = new PowerPointSession())
        {

            foreach (var fullName in fullNameList)
            {
                var pr = session.OpenPresentation(fullName, true, password);
                var slideCount = session.GetSlideCount(pr);
                count = 0;
                foundData.Clear();
                foundData.AppendLine();
                line[0] = "Slide";
                line[1] = "Area";
                line[2] = "Type";
                line[3] = "Content";
                foundData.AppendLine(string.Join("|", line));
                Array.Fill(line, "---");
                foundData.AppendLine(string.Join("|", line));
                var slides = pr.Slides;
                session.RegisterComObject(slides);
                for (var i = 1; i <= slideCount; i++)
                {
                    var slideName = $"Slide{i}";
                    var slide = slides[i];
                    session.RegisterComObject(slide);
                    var shapes = slide.Shapes;
                    session.RegisterComObject(shapes);
                    var foundText = session.FindShapesText(shapes, searchValue, matchPart, ignoreCase);
                    count += foundText.Count;
                    totalCount += foundText.Count;
                    foreach (var item in foundText)
                    {
                        line[0] =slideName;
                        line[1] = "page";
                        line[2] = session.EscapeMarkdownTableValue(item.Key);
                        line[3] = session.EscapeMarkdownTableValue(item.Value);
                        foundData.AppendLine(string.Join("|", line));
                    }
                    var notesPage = slide.NotesPage;
                    session.RegisterComObject(notesPage);
                    var notesShapes = notesPage.Shapes;
                    session.RegisterComObject(notesShapes);
                    var foundNotesText = session.FindShapesText(notesShapes, searchValue, matchPart, ignoreCase);
                    count += foundNotesText.Count;
                    totalCount += foundNotesText.Count;
                    foreach (var item in foundNotesText)
                    {
                        line[0] = slideName;
                        line[1] = "notes";
                        line[2] = session.EscapeMarkdownTableValue(item.Key);
                        line[3] = session.EscapeMarkdownTableValue(item.Value);
                        foundData.AppendLine(string.Join("|", line));
                    }

                }
                if (count > 0)
                {
                    foundData.Insert(0, $"`{count}` results in `{fullName}`:");
                    data.AppendLine(foundData.ToString());

                }
            }

        }
        data.Insert(0, $"Found a total of `{totalCount}` results for `{searchValue}` in all files.");
        return data.ToString();
    }

    [McpServerTool(Name = "powerpoint_run_macro"), Description("Run a macro of the specified PowerPoint file.")]
    public static string RunMacro([Description("The full path of the PowerPoint file.")] string fullName
    , [Description("The name of macro.")] string macroName
    , [Description("The parameters of macro. The maximum number is 30.")] string[]? macroParameters = null
    , [Description("Save the file after executing the macro.")] bool save = true
    , [Description("The password of the PowerPoint file, if there is one.")] string password = null)
    {
        var response = "";
        using (var session = new PowerPointSession())
        {
            var pr = session.OpenPresentation(fullName, false, password);
            try
            {
                response = session.RunMacro($"{pr.Name}!{macroName}", macroParameters);
                if (save)
                {
                    pr.Save();
                }
                pr.Close();
            }
            catch (Exception ex)
            {
                throw new McpException(ex.Message);
            }

        }
        if (string.IsNullOrEmpty(response))
        {
            response = $"{macroName} has been called.";
        }
        else
        {
            response = $"The result of {macroName}: {response}";
        }

        return response;
    }
}
