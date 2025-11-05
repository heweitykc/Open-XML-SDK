// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.ExcelAc;
using DocumentFormat.OpenXml.Office2021.Excel.NamedSheetViews;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AddNamedSheetView
{
    public class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Common.ExampleUtilities.ShowHelp(new string[]
                {
                    "NamedSheetView: ",
                    "Usage: NamedSheetView <filename> [jsonfile]",
                    "Where: <filename> is the .xlsx file in which to add a named sheet view.",
                    "       or .pptx file to copy slide 2 and insert at the end.",
                    "       [jsonfile] (optional) JSON file with PPT outline data to replace first slide placeholders.",
                });
            }
            else if (Common.ExampleUtilities.CheckIfFilesExist(args))
            {
                string filePath = args[0];
                string extension = Path.GetExtension(filePath).ToLower();

                if (extension == ".xlsx")
                {
                    InsertNamedSheetView(filePath);
                }
                else if (extension == ".pptx")
                {
                    string outputPath = filePath;

                    // 检查是否提供了 JSON 文件
                    if (args.Length >= 2 && File.Exists(args[1]))
                    {
                        string jsonContent = File.ReadAllText(args[1]);
                        
                        // 生成输出文件名
                        outputPath = GenerateOutputPath(filePath);
                        
                        // 复制源文件到新文件
                        File.Copy(filePath, outputPath, true);
                        Console.WriteLine($"Created new file: {outputPath}");
                        
                        // 在新文件上进行操作
                        ReplaceFirstSlideWithJson(outputPath, jsonContent);
                        ReplaceSecondSlideWithJson(outputPath, jsonContent);
                    }

                    CopyAndInsertSlide(outputPath);
                }
                else
                {
                    Console.WriteLine($"Unsupported file type: {extension}");
                    Console.WriteLine("Only .xlsx and .pptx files are supported.");
                }
            }
        }

        private static string GenerateOutputPath(string inputPath)
        {
            string directory = Path.GetDirectoryName(inputPath) ?? string.Empty;
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
            string extension = Path.GetExtension(inputPath);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            string outputFileName = $"{fileNameWithoutExt}_modified_{timestamp}{extension}";
            return Path.Combine(directory, outputFileName);
        }

        public static void InsertNamedSheetView(string xlsxPath)
        {
            if (xlsxPath == null)
            {
                throw new ArgumentNullException(nameof(xlsxPath));
            }

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(xlsxPath, true))
            {
                // 添加一个新的工作表 (Sheet)
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Any())
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                Sheet newSheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(newWorksheetPart),
                    SheetId = sheetId,
                    Name = "NewSheet" + sheetId
                };
                sheets.Append(newSheet);

                // 同时添加 NamedSheetView 到第一个工作表
                WorksheetPart firstWorksheetPart = workbookPart.GetPartsOfType<WorksheetPart>().First();
                NamedSheetViewsPart namedSheetViewsPart = firstWorksheetPart.AddNewPart<NamedSheetViewsPart>();

                NamedSheetView namedSheetView = new NamedSheetView();
                namedSheetView.Id = "{" + System.Guid.NewGuid().ToString().ToUpper() + "}";
                namedSheetView.Name = "testview";

                namedSheetViewsPart.NamedSheetViews = new NamedSheetViews(
                    namedSheetView);
                namedSheetViewsPart.NamedSheetViews.AddNamespaceDeclaration("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                workbookPart.Workbook.Save();

                Console.WriteLine($"New sheet '{newSheet.Name}' added successfully");
                Console.WriteLine("Named sheet view added to first sheet");
            }
        }

        public static void CopyAndInsertSlide(string pptxPath)
        {
            if (pptxPath == null)
            {
                throw new ArgumentNullException(nameof(pptxPath));
            }

            if (!File.Exists(pptxPath))
            {
                Console.WriteLine($"File not found: {pptxPath}");
                return;
            }

            using (PresentationDocument presentationDocument = PresentationDocument.Open(pptxPath, true))
            {
                PresentationPart presentationPart = presentationDocument.PresentationPart;
                
                if (presentationPart == null || presentationPart.Presentation == null)
                {
                    Console.WriteLine("Invalid presentation file");
                    return;
                }

                P.Presentation presentation = presentationPart.Presentation;
                P.SlideIdList slideIdList = presentation.SlideIdList;

                if (slideIdList == null || slideIdList.ChildElements.Count < 2)
                {
                    Console.WriteLine("Presentation must have at least 2 slides");
                    return;
                }

                // 获取第二页的 SlideId
                P.SlideId secondSlideId = slideIdList.ChildElements[1] as P.SlideId;
                if (secondSlideId == null)
                {
                    Console.WriteLine("Cannot find second slide");
                    return;
                }

                // 获取第二页的 SlidePart
                SlidePart secondSlidePart = presentationPart.GetPartById(secondSlideId.RelationshipId) as SlidePart;
                if (secondSlidePart == null)
                {
                    Console.WriteLine("Cannot get second slide part");
                    return;
                }

                // 列出第二页的所有元素
                Console.WriteLine("\n=== Slide 2 Elements ===");
                ListSlideElements(secondSlidePart);

                // 创建新的 SlidePart（复制）
                SlidePart newSlidePart = presentationPart.AddNewPart<SlidePart>();
                
                // 复制幻灯片内容
                newSlidePart.Slide = (P.Slide)secondSlidePart.Slide.CloneNode(true);

                // 给新幻灯片的所有文本添加时间戳
                AddTimestampToSlide(newSlidePart);

                // 复制所有关联的部分（图片、图表等）
                foreach (var part in secondSlidePart.Parts)
                {
                    string relationshipId = part.RelationshipId;
                    OpenXmlPart targetPart = part.OpenXmlPart;
                    
                    newSlidePart.AddPart(targetPart, relationshipId);
                }

                // 获取新的 SlideId
                uint maxSlideId = 256;
                foreach (P.SlideId slideId in slideIdList.ChildElements)
                {
                    if (slideId.Id > maxSlideId)
                    {
                        maxSlideId = slideId.Id;
                    }
                }

                // 创建新的 SlideId 并添加到最后
                P.SlideId newSlideId = new P.SlideId
                {
                    Id = maxSlideId + 1,
                    RelationshipId = presentationPart.GetIdOfPart(newSlidePart)
                };
                slideIdList.Append(newSlideId);

                // 保存更改
                presentation.Save();

                Console.WriteLine($"\nSuccessfully copied slide 2 and inserted at the end (position {slideIdList.ChildElements.Count})");
                Console.WriteLine($"Total slides: {slideIdList.ChildElements.Count}");
            }
        }

        private static void ListSlideElements(SlidePart slidePart)
        {
            if (slidePart?.Slide == null)
            {
                Console.WriteLine("Slide is empty");
                return;
            }

            P.Slide slide = slidePart.Slide;
            int elementCount = 0;

            // 统计各种形状
            if (slide.CommonSlideData?.ShapeTree != null)
            {
                var shapeTree = slide.CommonSlideData.ShapeTree;

                // 文本框和形状
                var shapes = shapeTree.Elements<P.Shape>();
                foreach (var shape in shapes)
                {
                    elementCount++;
                    string shapeName = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name ?? "Unnamed";
                    string shapeText = GetShapeText(shape);
                    
                    Console.WriteLine($"\n{elementCount}. Shape: {shapeName}");
                    if (!string.IsNullOrEmpty(shapeText))
                    {
                        Console.WriteLine($"   Text Content:");
                        Console.WriteLine($"   {new string('-', 60)}");
                        Console.WriteLine($"   {shapeText}");
                        Console.WriteLine($"   {new string('-', 60)}");
                    }
                    else
                    {
                        Console.WriteLine("   (No text content)");
                    }
                }

                // 图片
                var pictures = shapeTree.Elements<P.Picture>();
                foreach (var picture in pictures)
                {
                    elementCount++;
                    string pictureName = picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name ?? "Unnamed";
                    string description = picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description ?? "";
                    
                    Console.WriteLine($"{elementCount}. Picture: {pictureName}");
                    if (!string.IsNullOrEmpty(description))
                    {
                        Console.WriteLine($"   Description: {description}");
                    }
                }

                // 图表
                var graphicFrames = shapeTree.Elements<P.GraphicFrame>();
                foreach (var frame in graphicFrames)
                {
                    elementCount++;
                    string frameName = frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name ?? "Unnamed";
                    Console.WriteLine($"{elementCount}. Graphic Frame (Chart/Table/SmartArt): {frameName}");
                }

                // 组合形状
                var groupShapes = shapeTree.Elements<P.GroupShape>();
                foreach (var group in groupShapes)
                {
                    elementCount++;
                    string groupName = group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name ?? "Unnamed";
                    Console.WriteLine($"\n{elementCount}. Group Shape: {groupName}");

                    // 递归列出组内的子元素
                    int subCount = ListGroupShapeElements(group, "   ");
                    Console.WriteLine($"   Total sub-elements: {subCount}");
                }

                // 连接线
                var connectionShapes = shapeTree.Elements<P.ConnectionShape>();
                foreach (var conn in connectionShapes)
                {
                    elementCount++;
                    string connName = conn.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Name ?? "Unnamed";
                    Console.WriteLine($"{elementCount}. Connection Shape: {connName}");
                }
            }

            // 列出关联的部分（图片、图表等资源）
            Console.WriteLine("\n--- Associated Parts ---");
            int partCount = 0;
            foreach (var part in slidePart.Parts)
            {
                partCount++;
                string partType = part.OpenXmlPart.GetType().Name;
                string relationshipId = part.RelationshipId;
                Console.WriteLine($"{partCount}. {partType} (RelId: {relationshipId})");
            }

            // 背景和其他属性
            if (slide.CommonSlideData?.Background != null)
            {
                Console.WriteLine("\n--- Slide Properties ---");
                Console.WriteLine("Has custom background");
            }

            Console.WriteLine($"\nTotal elements on slide: {elementCount}");
            Console.WriteLine($"Total associated parts: {partCount}");
        }

        private static void ReplaceFirstSlideWithJson(string pptxPath, string jsonContent)
        {
            Console.WriteLine("\n=== Replacing First Slide Placeholders with JSON Data ===");

            try
            {
                using (PresentationDocument presentationDocument = PresentationDocument.Open(pptxPath, true))
                {
                    PresentationPart presentationPart = presentationDocument.PresentationPart;
                    if (presentationPart?.Presentation?.SlideIdList == null)
                    {
                        Console.WriteLine("Invalid presentation");
                        return;
                    }

                    // 解析 JSON
                    using (JsonDocument doc = JsonDocument.Parse(jsonContent))
                    {
                        JsonElement root = doc.RootElement;

                        // 提取替换值
                        var replacements = new Dictionary<string, string>
                        {
                            { "{ppt_title}", root.GetProperty("title").GetString() ?? string.Empty },
                            { "{ppt_subtitle}", root.GetProperty("subtitle").GetString() ?? string.Empty },
                            { "{ppt_author}", root.GetProperty("author").GetString() ?? string.Empty },
                            { "{ppt_website}", root.GetProperty("website").GetString() ?? string.Empty },
                        };

                        // 获取第一页
                        P.SlideId firstSlideId = presentationPart.Presentation.SlideIdList.ChildElements[0] as P.SlideId;
                        if (firstSlideId == null)
                        {
                            Console.WriteLine("Cannot find first slide");
                            return;
                        }

                        SlidePart firstSlidePart = presentationPart.GetPartById(firstSlideId.RelationshipId) as SlidePart;
                        if (firstSlidePart == null)
                        {
                            Console.WriteLine("Cannot get first slide part");
                            return;
                        }

                        // 替换所有形状中的占位符文本
                        int replacedCount = 0;
                        var shapes = firstSlidePart.Slide.Descendants<P.Shape>();
                        foreach (var shape in shapes)
                        {
                            replacedCount += ReplaceShapePlaceholders(shape, replacements);
                        }

                        presentationPart.Presentation.Save();
                        Console.WriteLine($"Successfully replaced {replacedCount} placeholders on first slide");
                    }
                }
            }
            catch (JsonException ex)
            {
                Console.WriteLine($"JSON parsing error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error replacing placeholders: {ex.Message}");
            }
        }

        private static void ReplaceSecondSlideWithJson(string pptxPath, string jsonContent)
        {
            Console.WriteLine("\n=== Replacing Second Slide (TOC) with JSON Data ===");

            try
            {
                using (PresentationDocument presentationDocument = PresentationDocument.Open(pptxPath, true))
                {
                    PresentationPart presentationPart = presentationDocument.PresentationPart;
                    if (presentationPart?.Presentation?.SlideIdList == null)
                    {
                        Console.WriteLine("Invalid presentation");
                        return;
                    }

                    if (presentationPart.Presentation.SlideIdList.ChildElements.Count < 2)
                    {
                        Console.WriteLine("Presentation does not have a second slide");
                        return;
                    }

                    // 解析 JSON
                    using (JsonDocument doc = JsonDocument.Parse(jsonContent))
                    {
                        JsonElement root = doc.RootElement;

                        // 获取 parts 数组
                        if (!root.TryGetProperty("parts", out JsonElement partsArray))
                        {
                            Console.WriteLine("JSON does not contain 'parts' array");
                            return;
                        }

                        // 提取所有 part titles
                        var partTitles = new List<string>();
                        foreach (var part in partsArray.EnumerateArray())
                        {
                            if (part.TryGetProperty("title", out JsonElement title))
                            {
                                partTitles.Add(title.GetString() ?? string.Empty);
                            }
                        }

                        Console.WriteLine($"Found {partTitles.Count} part titles in JSON");

                        // 获取第二页
                        P.SlideId secondSlideId = presentationPart.Presentation.SlideIdList.ChildElements[1] as P.SlideId;
                        if (secondSlideId == null)
                        {
                            Console.WriteLine("Cannot find second slide");
                            return;
                        }

                        SlidePart secondSlidePart = presentationPart.GetPartById(secondSlideId.RelationshipId) as SlidePart;
                        if (secondSlidePart == null)
                        {
                            Console.WriteLine("Cannot get second slide part");
                            return;
                        }

                        // 查找所有包含 {part_title_x} 占位符的形状
                        var shapes = secondSlidePart.Slide.Descendants<P.Shape>().ToList();
                        var placeholderShapes = new Dictionary<int, P.Shape>();

                        foreach (var shape in shapes)
                        {
                            string shapeText = GetShapeText(shape);
                            // 查找 {part_title_1}, {part_title_2}, ... 等占位符（索引从1开始）
                            for (int i = 1; i <= 10; i++) // 最多支持10个
                            {
                                string placeholder = $"{{part_title_{i}}}";
                                if (shapeText.Contains(placeholder))
                                {
                                    placeholderShapes[i] = shape;
                                    Console.WriteLine($"Found placeholder '{placeholder}' in shape");
                                    break;
                                }
                            }
                        }

                        Console.WriteLine($"Found {placeholderShapes.Count} placeholder shapes");

                        // 处理占位符和标题的匹配（索引从1开始）
                        int maxIndex = Math.Max(
                            placeholderShapes.Count > 0 ? placeholderShapes.Keys.Max() : 0,
                            partTitles.Count
                        );

                        for (int i = 1; i <= maxIndex; i++)
                        {
                            int titleIndex = i - 1; // 标题数组索引从0开始
                            
                            if (placeholderShapes.ContainsKey(i) && titleIndex < partTitles.Count)
                            {
                                // 有占位符且有对应的标题，进行替换
                                Console.WriteLine($"  Replacing {{part_title_{i}}} with: '{partTitles[titleIndex]}'");
                                ReplaceShapeContent(placeholderShapes[i], partTitles[titleIndex]);
                            }
                            else if (placeholderShapes.ContainsKey(i) && titleIndex >= partTitles.Count)
                            {
                                // 有占位符但没有对应的标题，删除形状
                                Console.WriteLine($"  Deleting extra placeholder shape: {{part_title_{i}}}");
                                DeleteShape(placeholderShapes[i]);
                            }
                            else if (!placeholderShapes.ContainsKey(i) && titleIndex < partTitles.Count)
                            {
                                // 有标题但没有占位符，跳过
                                Console.WriteLine($"  Skipping title (no placeholder): '{partTitles[titleIndex]}'");
                            }
                        }

                        presentationPart.Presentation.Save();
                        Console.WriteLine("Second slide updated successfully");
                    }
                }
            }
            catch (JsonException ex)
            {
                Console.WriteLine($"JSON parsing error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error replacing second slide: {ex.Message}");
            }
        }

        private static void ReplaceShapeContent(P.Shape shape, string newContent)
        {
            if (shape.TextBody == null)
            {
                return;
            }

            // 清空所有段落
            var paragraphs = shape.TextBody.Elements<A.Paragraph>().ToList();
            foreach (var para in paragraphs)
            {
                para.RemoveAllChildren();
            }

            // 如果没有段落，创建一个
            if (paragraphs.Count == 0)
            {
                var newParagraph = new A.Paragraph();
                shape.TextBody.AppendChild(newParagraph);
                paragraphs.Add(newParagraph);
            }

            // 在第一个段落中设置新文本
            var firstParagraph = paragraphs[0];
            var newRun = new A.Run();
            var newText = new A.Text(newContent);
            newRun.AppendChild(newText);
            firstParagraph.AppendChild(newRun);
        }

        private static void DeleteShape(P.Shape shape)
        {
            shape.Remove();
        }

        private static int ReplaceShapePlaceholders(P.Shape shape, Dictionary<string, string> replacements)
        {
            int count = 0;

            if (shape.TextBody == null)
            {
                return count;
            }

            // 首先检查整个形状的文本是否包含占位符
            string shapeFullText = GetShapeText(shape);

            foreach (var kvp in replacements)
            {
                if (shapeFullText.Contains(kvp.Key))
                {
                    // 找到占位符，用 JSON 值完全覆盖整个文本框内容
                    Console.WriteLine($"  Found placeholder '{kvp.Key}' in shape");
                    Console.WriteLine($"  Original text: '{shapeFullText.Trim()}'");
                    Console.WriteLine($"  Replacing entire content with: '{kvp.Value}'");

                    // 清空所有段落
                    var paragraphs = shape.TextBody.Elements<A.Paragraph>().ToList();
                    foreach (var para in paragraphs)
                    {
                        para.RemoveAllChildren();
                    }

                    // 如果没有段落，创建一个
                    if (paragraphs.Count == 0)
                    {
                        var newParagraph = new A.Paragraph();
                        shape.TextBody.AppendChild(newParagraph);
                        paragraphs.Add(newParagraph);
                    }

                    // 在第一个段落中设置新文本
                    var firstParagraph = paragraphs[0];
                    var newRun = new A.Run();
                    var newText = new A.Text(kvp.Value);
                    newRun.AppendChild(newText);
                    firstParagraph.AppendChild(newRun);

                    count++;
                    break; // 一个形状只替换一次
                }
            }

            return count;
        }

        private static void AddTimestampToSlide(SlidePart slidePart)
        {
            if (slidePart?.Slide?.CommonSlideData?.ShapeTree == null)
            {
                return;
            }

            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            Console.WriteLine($"\nAdding timestamp to all text: {timestamp}");

            var shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;

            // 处理普通形状
            var shapes = shapeTree.Descendants<P.Shape>();
            foreach (var shape in shapes)
            {
                AddTimestampToShape(shape, timestamp);
            }

            Console.WriteLine("Timestamp added successfully!");
        }

        private static void AddTimestampToShape(P.Shape shape, string timestamp)
        {
            if (shape.TextBody == null)
            {
                return;
            }

            // 获取所有段落
            var paragraphs = shape.TextBody.Elements<A.Paragraph>().ToList();

            // 遍历每个段落中的文本运行
            foreach (var paragraph in paragraphs)
            {
                var runs = paragraph.Elements<A.Run>().ToList();
                foreach (var run in runs)
                {
                    var textElement = run.Elements<A.Text>().FirstOrDefault();
                    if (textElement != null && !string.IsNullOrWhiteSpace(textElement.Text))
                    {
                        // 在原文本后添加时间戳
                        textElement.Text = $"{textElement.Text} [{timestamp}]";
                    }
                }
            }
        }

        private static int ListGroupShapeElements(P.GroupShape groupShape, string indent)
        {
            int count = 0;

            // 处理组内的普通形状
            var shapes = groupShape.Elements<P.Shape>();
            foreach (var shape in shapes)
            {
                count++;
                string shapeName = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name ?? "Unnamed";
                string shapeText = GetShapeText(shape);

                Console.WriteLine($"{indent}{count}. Shape: {shapeName}");
                if (!string.IsNullOrEmpty(shapeText))
                {
                    Console.WriteLine($"{indent}   Text Content:");
                    Console.WriteLine($"{indent}   {new string('-', 50)}");
                    // 对每一行添加缩进
                    var lines = shapeText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var line in lines)
                    {
                        Console.WriteLine($"{indent}   {line}");
                    }
                    Console.WriteLine($"{indent}   {new string('-', 50)}");
                }
                else
                {
                    Console.WriteLine($"{indent}   (No text content)");
                }
            }

            // 处理组内的图片
            var pictures = groupShape.Elements<P.Picture>();
            foreach (var picture in pictures)
            {
                count++;
                string pictureName = picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name ?? "Unnamed";
                Console.WriteLine($"{indent}{count}. Picture: {pictureName}");
            }

            // 递归处理嵌套的组合形状
            var nestedGroups = groupShape.Elements<P.GroupShape>();
            foreach (var nestedGroup in nestedGroups)
            {
                count++;
                string nestedGroupName = nestedGroup.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name ?? "Unnamed";
                Console.WriteLine($"{indent}{count}. Nested Group: {nestedGroupName}");
                int nestedCount = ListGroupShapeElements(nestedGroup, indent + "   ");
                Console.WriteLine($"{indent}   (Contains {nestedCount} sub-elements)");
            }

            return count;
        }

        private static string GetShapeText(P.Shape shape)
        {
            if (shape.TextBody == null)
            {
                return string.Empty;
            }

            var result = new System.Text.StringBuilder();
            var paragraphs = shape.TextBody.Elements<A.Paragraph>();

            foreach (var paragraph in paragraphs)
            {
                var runs = paragraph.Elements<A.Run>();
                foreach (var run in runs)
                {
                    var text = run.Elements<A.Text>().FirstOrDefault();
                    if (text != null)
                    {
                        result.Append(text.Text);
                    }
                }

                // 每个段落后添加换行
                result.AppendLine();
            }

            return result.ToString().TrimEnd();
        }
    }
}
