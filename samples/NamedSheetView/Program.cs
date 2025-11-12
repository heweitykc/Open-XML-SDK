// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.ExcelAc;
using DocumentFormat.OpenXml.Office2021.Excel.NamedSheetViews;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Globalization;
using System.Security.Cryptography;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

#nullable enable
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8603 // Possible null reference return.
#pragma warning disable CS8604 // Possible null reference argument.

namespace AddNamedSheetView
{
    public class Program
    {
        private static readonly HttpClient HttpClient = new HttpClient();
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions(JsonSerializerDefaults.Web);
        private static readonly string CacheDirectory = Path.Combine(AppContext.BaseDirectory, "aippt_tpl");
        private static readonly string AppDataDirectory = Path.Combine(AppContext.BaseDirectory, "appdata");
        private const string OutputDirectoryName = "ai_ppt_generator";
        private const string DefaultPrefix = "http://+:8050/";
        private static readonly HashSet<char> InvalidFileNameChars = new HashSet<char>(Path.GetInvalidFileNameChars());
        private static readonly ConcurrentDictionary<string, SemaphoreSlim> TemplateLocks = new ConcurrentDictionary<string, SemaphoreSlim>(StringComparer.OrdinalIgnoreCase);

        public static async Task Main(string[] args)
        {
            if (args.Length > 0 && string.Equals(args[0], "--server", StringComparison.OrdinalIgnoreCase))
            {
                string? prefix = args.Length > 1 ? args[1] : null;
                await RunServerAsync(prefix).ConfigureAwait(false);
                return;
            }

            if (args.Length < 2)
            {
                Common.ExampleUtilities.ShowHelp(new string[]
                {
                    "NamedSheetView: ",
                    "Usage: NamedSheetView <filename> [jsonfile] [outputPath]",
                    "Where: <filename> is the .xlsx file in which to add a named sheet view.",
                    "       or .pptx file to copy slide 2 and insert at the end.",
                    "       [jsonfile] (optional) JSON file with PPT outline data to replace first slide placeholders.",
                    "       [outputPath] (optional) Output file path. If omitted, a path is generated automatically.",
                });
                return;
            }

            try
            {
                string? outputPathArgument = args.Length > 2 ? args[2] : null;
                string outputPath = GeneratePresentation(args[0], args[1], outputPathArgument);
                Console.WriteLine(outputPath);
            }
            catch (Exception ex)
            {
                Log(() => $"Processing failed: {ex.Message}");
#if DEBUG
                Log(() => ex.ToString());
#endif
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

            public static string GeneratePresentation(string sourceFilePath, string jsonFilePath, string? outputFilePath = null)
            {
                if (string.IsNullOrWhiteSpace(sourceFilePath))
                {
                    throw new ArgumentException("Source file path cannot be null or empty.", nameof(sourceFilePath));
                }

                if (string.IsNullOrWhiteSpace(jsonFilePath))
                {
                    throw new ArgumentException("JSON file path cannot be null or empty.", nameof(jsonFilePath));
                }

                string jsonContent = File.ReadAllText(jsonFilePath);

                // 生成输出文件名
                string outputPath = ResolveOutputPath(sourceFilePath, outputFilePath);

                // 复制源文件到新文件
                File.Copy(sourceFilePath, outputPath, true);
                Log(() => $"Created new file: {outputPath}");

                using (PresentationDocument presentationDocument = PresentationDocument.Open(outputPath, true))
                {
                    PresentationPart? presentationPart = presentationDocument.PresentationPart;
                    if (presentationPart?.Presentation?.SlideIdList == null)
                    {
                        throw new InvalidOperationException("Invalid presentation");
                    }

                    PresentationPart ensuredPresentationPart = presentationPart!;

                    // 记录原始模板 PPT 的最后一页索引（在添加新 slides 之前）
                    int originalLastSlideIndex = ensuredPresentationPart.Presentation!.SlideIdList!.ChildElements.Count - 1;
                    Log(() => $"Original template has {originalLastSlideIndex + 1} slides");

                    // 解析 JSON
                    using (JsonDocument doc = JsonDocument.Parse(jsonContent))
                    {
                        // 在新文件上进行操作
                        Measure("ReplaceFirstSlideWithJson", () => ReplaceFirstSlideWithJson(ensuredPresentationPart, doc));
                        Measure("ReplaceSecondSlideWithJson", () => ReplaceSecondSlideWithJson(ensuredPresentationPart, doc));

                        // 生成各个 part 的 slides
                        Measure("GeneratePartSlidesFromJson", () => GeneratePartSlidesFromJson(ensuredPresentationPart, doc));

                        // 从原始模板复制最后一页并替换结束页占位符
                        Measure("CopyAndReplaceLastSlideFromTemplate", () => CopyAndReplaceLastSlideFromTemplate(ensuredPresentationPart, originalLastSlideIndex, doc));

                        //删除从索引 [2 - $originalLastSlideIndex] 的所有slides
                        Measure("DeleteSlidesFromIndex", () => DeleteSlidesFromIndex(ensuredPresentationPart, 2, originalLastSlideIndex));

                        // 媒体资源去重与清理
                        Measure("DeduplicateMediaResources", () => DeduplicateMediaResources(ensuredPresentationPart));
                        Measure("CleanupUnusedMediaResources", () => CleanupUnusedMediaResources(ensuredPresentationPart));
                    }

                    ensuredPresentationPart.Presentation!.Save();
                }

                Log(() => $"{outputPath} saved");
                return outputPath;
            }

        private static string ResolveOutputPath(string inputPath, string? providedOutputPath)
        {
            if (!string.IsNullOrWhiteSpace(providedOutputPath))
            {
                string? directory = Path.GetDirectoryName(providedOutputPath);
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                return providedOutputPath;
            }

            return GenerateOutputPath(inputPath);
        }

        private static string GenerateOutputPath(string inputPath)
        {
            string directory = Path.GetDirectoryName(inputPath) ?? string.Empty;
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
            string extension = Path.GetExtension(inputPath);

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
            string guid = Guid.NewGuid().ToString("N").Substring(0, 8);

            string outputDirectory = Path.Combine(directory, "output");
            Directory.CreateDirectory(outputDirectory);

            string outputFileName = $"{fileNameWithoutExt}_modified_{timestamp}_{guid}{extension}";
            return Path.Combine(outputDirectory, outputFileName);
        }

        private static void DeleteSlidesFromIndex(PresentationPart presentationPart, int startIndex, int endIndex)
        {
            Log(() => $"\n=== Deleting Original Template Slides from index {startIndex} to {endIndex} ===");

            var slideIdList = presentationPart.Presentation?.SlideIdList;
            if (slideIdList == null || slideIdList.ChildElements.Count == 0)
            {
                Log(() => "No slides found in presentation");
                return;
            }

            // 验证索引范围
            if (startIndex < 0 || startIndex >= slideIdList.ChildElements.Count)
            {
                Log(() => $"Invalid start index: {startIndex}");
                return;
            }

            if (endIndex < startIndex)
            {
                Log(() => "No slides to delete (endIndex < startIndex)");
                return;
            }

            // 确保 endIndex 不超过实际范围
            int actualEndIndex = Math.Min(endIndex, slideIdList.ChildElements.Count - 1);
            
            Log(() => $"Will delete slides from index {startIndex} to {actualEndIndex}");

            // 收集要删除的 SlideId（从后往前，避免索引变化问题）
            var slidesToDelete = new List<P.SlideId>();
            
            for (int i = actualEndIndex; i >= startIndex; i--)
            {
                if (i < slideIdList.ChildElements.Count)
                {
                    var slideId = slideIdList.ChildElements[i] as P.SlideId;
                    if (slideId != null)
                    {
                        slidesToDelete.Add(slideId);
                    }
                }
            }

            Log(() => $"Found {slidesToDelete.Count} slides to delete");

            // 删除 slides（已经按倒序排列，从后往前删）
            int deletedCount = 0;
            foreach (var slideId in slidesToDelete)
            {
                // 获取 SlidePart 并删除
                string? relationshipId = slideId.RelationshipId;
                if (string.IsNullOrEmpty(relationshipId))
                {
                    continue;
                }
                
                try
                {
                    // 删除 SlidePart
                    var slidePart = presentationPart.GetPartById(relationshipId) as SlidePart;
                    if (slidePart != null)
                    {
                        presentationPart.DeletePart(slidePart);
                    }

                    // 从 SlideIdList 中移除
                    slideId.Remove();
                    deletedCount++;
                    
                    Log(() => $"  Deleted slide with relationship ID: {relationshipId}");
                }
                catch (Exception ex)
                {
                    Log(() => $"  Failed to delete slide with relationship ID {relationshipId}: {ex.Message}");
                }
            }

            Log(() => $"Successfully deleted {deletedCount} slides");
        }

        private static void CopyAndReplaceLastSlideFromTemplate(PresentationPart presentationPart, int originalLastSlideIndex, JsonDocument doc)
        {
            Log(() => "\n=== Copying and Replacing Last Slide from Original Template ===");

            JsonElement root = doc.RootElement;

            // 获取 slide 列表
            var slideIdList = presentationPart.Presentation?.SlideIdList;
            if (slideIdList == null || slideIdList.ChildElements.Count == 0)
            {
                throw new ArgumentException("No slides found in presentation");
            }

            // 确保原始索引有效
            if (originalLastSlideIndex < 0 || originalLastSlideIndex >= slideIdList.ChildElements.Count)
            {                
                throw new ArgumentException($"Invalid original slide index: {originalLastSlideIndex}");
            }

            // 获取原始模板的最后一个 slide (使用记录的索引)
            P.SlideId templateLastSlideId = slideIdList.ChildElements[originalLastSlideIndex] as P.SlideId;
            if (templateLastSlideId == null)
            {
                throw new ArgumentException("Cannot find template's last slide");
            }

            string? templateLastRelationshipId = templateLastSlideId.RelationshipId;
            if (string.IsNullOrEmpty(templateLastRelationshipId))
            {
                throw new ArgumentException("Template last slide relationship ID is null");
            }

            SlidePart templateLastSlidePart = presentationPart.GetPartById(templateLastRelationshipId) as SlidePart;
            if (templateLastSlidePart == null)
            {                
                throw new ArgumentException("Cannot get template's last slide part");
            }

            Log(() => $"Found template's last slide (index {originalLastSlideIndex}), copying...");

            // 复制原始模板的最后一页
            SlidePart newSlidePart = CopySlide(presentationPart, templateLastSlidePart);
            if (newSlidePart == null)
            {                
                throw new ArgumentException("Failed to copy template's last slide");                
            }

            // 准备替换值
            var replacements = new Dictionary<string, string>();

            // 添加 end_title
            if (TryGetJsonString(root, "endtitle", out var endTitleValue))
            {
                replacements["{end_title}"] = endTitleValue;
                Log(() => $"  endtitle: {endTitleValue}");
            }

            if (TryGetJsonString(root, "author", out var authorValue))
            {
                replacements["{ppt_author}"] = authorValue;
                Log(() => $"  author: {authorValue}");
            }

            if (TryGetJsonString(root, "website", out var websiteValue))
            {
                replacements["{ppt_website}"] = websiteValue;
                Log(() => $"  website: {websiteValue}");
            }

            // 替换新 slide 中的占位符
            int replacedCount = 0;
            var shapes = newSlidePart.Slide.Descendants<P.Shape>();
            foreach (var shape in shapes)
            {
                replacedCount += ReplaceShapePlaceholders(shape, replacements);
            }

            Log(() => $"Successfully copied template's last slide and replaced {replacedCount} placeholders");
        }

        private static void GeneratePartSlidesFromJson(PresentationPart presentationPart, JsonDocument doc)
        {
            Log(() => "\n=== Generating Part Slides from JSON ===");

            JsonElement root = doc.RootElement;
            var chapter_title_search = "{part_title_";
            var chapter_subtitle_search = "{part_subtitle_";

            // 获取 parts 数组
            if (!root.TryGetProperty("parts", out JsonElement partsArray))
            {                
                throw new ArgumentException("JSON does not contain 'parts' array");
            }

            // 查找所有包含 {part_subtitle_ 的 slides
            var templateSlides = FindSlidesWithKeyword(presentationPart, chapter_subtitle_search);
            
            if (templateSlides.Count == 0)
            {
                throw new ArgumentException("No slides found containing '{part_subtitle_' placeholder");
            }

            Log(() => $"Found {templateSlides.Count} template slides containing '{{part_subtitle_}}'");

            var random = new Random();
            var partTemplateSelector = new TemplateSelector<SlidePart>(random);
            int partIndex = 1;

            // 遍历 JSON 中的每个 part
            foreach (var part in partsArray.EnumerateArray())
            {
                if (!part.TryGetProperty("title", out JsonElement title))
                {
                    throw new ArgumentException($"Part {partIndex} is missing required 'title' field.");
                }

                string partTitle = title.GetString() ?? string.Empty;
                Log(() => $"\n--- Processing Part {partIndex}: {partTitle} ---");

                // 从模板 slides 中随机选择一个
                var selectedTemplate = partTemplateSelector.GetNext(templateSlides);
                if (selectedTemplate == null)
                {
                    throw new ArgumentException("Failed to select template slide for part generation.");
                }
                Log(() => $"  Selected random template slide");

                // 复制 slide 并插入到最后
                var newSlidePart = CopySlide(presentationPart, selectedTemplate);
                if (newSlidePart == null)
                {
                    throw new ArgumentException($"Failed to copy slide for part {partIndex}.");
                }

                // 准备替换字典
                var replacements = new Dictionary<string, string>
                {
                    { chapter_title_search, partTitle }
                };

                // 处理 part_subtitle
                string partSubtitle = string.Empty;
                if (part.TryGetProperty("subtitle", out JsonElement subtitle))
                {
                    partSubtitle = subtitle.GetString() ?? string.Empty;
                }

                if (!string.IsNullOrEmpty(partSubtitle))
                {
                    // 有 subtitle，添加到替换字典
                    replacements[chapter_subtitle_search] = partSubtitle;
                    Log(() => $"  Part subtitle: {partSubtitle}");
                }

                // 替换占位符
                int replacedCount = 0;
                var shapes = newSlidePart.Slide.Descendants<P.Shape>().ToList();
                foreach (var shape in shapes)
                {
                    replacedCount += ReplaceShapePlaceholders(shape, replacements);
                }

                // 如果 subtitle 为空，删除包含 {part_subtitle} 的 shape
                if (string.IsNullOrEmpty(partSubtitle))
                {
                    var shapesToDelete = new List<P.Shape>();
                    foreach (var shape in shapes)
                    {
                        string shapeText = GetShapeText(shape);
                        if (shapeText.Contains(chapter_subtitle_search, StringComparison.OrdinalIgnoreCase))
                        {
                            shapesToDelete.Add(shape);
                        }
                    }

                    foreach (var shape in shapesToDelete)
                    {
                        Log(() => $"  Deleting shape containing '{{part_subtitle}}' (subtitle is empty)");
                        DeleteShape(shape);
                    }
                }

                Log(() => $"  Created new slide, replaced {replacedCount} placeholders");
                
                // 处理该 part 中的 chapters
                if (part.TryGetProperty("chapters", out JsonElement chaptersArray))
                {
                    GenerateChapterSlidesForPart(presentationPart, chaptersArray, partIndex);
                }
                
                partIndex++;
            }

            Log(() => $"\n=== Generated {partIndex - 1} part slides ===");
        }

        // 为一个 part 生成 chapter slides
        private static void GenerateChapterSlidesForPart(PresentationPart presentationPart, JsonElement chaptersArray, int partIndex)
        {
            Log(() => $"\n  === Generating Chapter Slides for Part {partIndex} ===");

            var random = new Random();
            var chapterTemplateSelector = new TemplateSelector<SlidePart>(random);
            int chapterIndex = 1;

            // 遍历每个 chapter
            foreach (var chapter in chaptersArray.EnumerateArray())
            {
                if (!chapter.TryGetProperty("title", out JsonElement chapterTitle))
                {
                    throw new ArgumentException($"Chapter {chapterIndex} missing 'title' field.");
                }

                string titleText = chapterTitle.GetString() ?? string.Empty;
                Log(() => $"\n    --- Processing Chapter {chapterIndex}: {titleText} ---");

                // 获取当前 chapter 的 sections（先声明变量）
                JsonElement sectionsArray = default;
                bool hasSections = chapter.TryGetProperty("sections", out sectionsArray);
                int sectionsCount = hasSections ? sectionsArray.GetArrayLength() : 0;
                
                if (hasSections)
                {
                    Log(() => $"      Chapter has {sectionsCount} sections");
                }

                // 查找适配当前 sections 数量的 slides（包含 {chapter_title} 占位符）
                var chapterTemplateSlides = FindSlidesWithKeyword(presentationPart, "{chapter_title}", sectionsCount);
                // for (int i = 0; i < chapterTemplateSlides.Count; i++)
                // {
                //     int templateIndex = GetSlideIndex(presentationPart, chapterTemplateSlides[i]);
                //     Log(() => $"      chapterTemplateSlides[{i}] found at template index {templateIndex}");
                // }
                
                if (chapterTemplateSlides.Count == 0)
                {                    
                    throw new ArgumentException($"No slides found matching sections count {sectionsCount}.");
                }

                Log(() => $"      Found {chapterTemplateSlides.Count} matching chapter template slides");

                // 从模板 slides 中随机选择一个
                var selectedTemplate = chapterTemplateSelector.GetNext(chapterTemplateSlides);
                if (selectedTemplate == null)
                {
                    throw new ArgumentException("Failed to select chapter template slide.");
                }
                int selectedTemplateIndex = GetSlideIndex(presentationPart, selectedTemplate);
                Log(() => $"      Selected random chapter template slide, at template index {selectedTemplateIndex}");

                // 复制 slide 并插入到最后
                var newSlidePart = CopySlide(presentationPart, selectedTemplate);
                if (newSlidePart == null)
                {
                    throw new ArgumentException($"Failed to copy slide for chapter {chapterIndex}.");
                }

                // 构建替换字典
                var replacements = new Dictionary<string, string>
                {
                    { "{chapter_title}", titleText }
                };

                // 处理 chapter subtitle
                string chapterSubtitleText = string.Empty;
                if (chapter.TryGetProperty("subtitle", out JsonElement chapterSubtitle))
                {
                    chapterSubtitleText = chapterSubtitle.GetString() ?? string.Empty;
                }

                if (!string.IsNullOrEmpty(chapterSubtitleText))
                {
                    replacements["{chapter_subtitle}"] = chapterSubtitleText;
                }

                // 处理 sections（使用之前获取的 sectionsArray）
                if (hasSections)
                {
                    int sectionIndex = 1;
                    foreach (var section in sectionsArray.EnumerateArray())
                    {
                        // 添加索引占位符 {s1}, {s2}, {s3}, ...
                        replacements[$"{{s{sectionIndex}}}"] = sectionIndex.ToString();

                        // 添加 section title
                        if (section.TryGetProperty("title", out JsonElement sectionTitle))
                        {
                            string sectionTitleText = sectionTitle.GetString() ?? string.Empty;
                            replacements[$"{{section_title_{sectionIndex}}}"] = sectionTitleText;
                        }

                        // 添加 section subtitle
                        if (section.TryGetProperty("subtitle", out JsonElement sectionSubtitle))
                        {
                            string sectionSubtitleText = sectionSubtitle.GetString() ?? string.Empty;
                            replacements[$"{{section_subtitle_{sectionIndex}}}"] = sectionSubtitleText;
                        }

                        // 处理 items（如果需要）
                        if (section.TryGetProperty("items", out JsonElement itemsArray))
                        {
                            int itemIndex = 1;
                            foreach (var item in itemsArray.EnumerateArray())
                            {
                                if (item.TryGetProperty("title", out JsonElement itemTitle))
                                {
                                    string itemTitleText = itemTitle.GetString() ?? string.Empty;
                                    replacements[$"{{item_{sectionIndex}_{itemIndex}}}"] = itemTitleText;
                                    
                                    // 如果 section 只有一个 item，也可以用 item 的 title 作为 section_subtitle
                                    if (itemIndex == 1 && !replacements.ContainsKey($"{{section_subtitle_{sectionIndex}}}"))
                                    {
                                        replacements[$"{{section_subtitle_{sectionIndex}}}"] = itemTitleText;
                                    }
                                }
                                itemIndex++;
                            }
                        }

                        sectionIndex++;
                    }
                }

                // 替换新 slide 中的占位符
                int replacedCount = 0;
                var shapes = newSlidePart.Slide.Descendants<P.Shape>().ToList();
                foreach (var shape in shapes)
                {
                    replacedCount += ReplaceShapePlaceholders(shape, replacements);
                }

                // 如果 chapter_subtitle 为空，删除包含该占位符的 shape
                if (string.IsNullOrEmpty(chapterSubtitleText))
                {
                    var shapesToDelete = new List<P.Shape>();
                    foreach (var shape in shapes)
                    {
                        string shapeText = GetShapeText(shape);
                        if (shapeText.Contains("{chapter_subtitle}", StringComparison.OrdinalIgnoreCase))
                        {
                            shapesToDelete.Add(shape);
                        }
                    }

                    foreach (var shape in shapesToDelete)
                    {
                        Log(() => $"      Deleting shape containing '{{chapter_subtitle}}' (subtitle is empty)");
                        DeleteShape(shape);
                    }
                }

                Log(() => $"      Created chapter slide, replaced {replacedCount} placeholders");
                chapterIndex++;
            }

            Log(() => $"  === Generated {chapterIndex - 1} chapter slides for Part {partIndex} ===");
        }

        // 复制一个 slide 并插入到 presentation 的最后
        private static SlidePart CopySlide(PresentationPart presentationPart, SlidePart sourceSlide)
        {
            // 创建新的 SlidePart
            SlidePart newSlidePart = presentationPart.AddNewPart<SlidePart>();

            // 复制 slide 内容
            newSlidePart.Slide = (P.Slide)sourceSlide.Slide.CloneNode(true);

            // 复制 SlideLayout 关系
            if (sourceSlide.SlideLayoutPart != null)
            {
                newSlidePart.AddPart(sourceSlide.SlideLayoutPart);
            }

            // 复制所有图片资源
            foreach (var imagePart in sourceSlide.ImageParts)
            {
                ImagePart newImagePart = newSlidePart.AddImagePart(imagePart.ContentType);
                using (var stream = imagePart.GetStream())
                {
                    newImagePart.FeedData(stream);
                }

                // 更新关系 ID
                string oldRelId = sourceSlide.GetIdOfPart(imagePart);
                string newRelId = newSlidePart.GetIdOfPart(newImagePart);

                // 替换 slide 中的 RelationshipId 引用
                ReplaceRelationshipId(newSlidePart.Slide, oldRelId, newRelId);
            }

            // 复制图表部分
            foreach (var chartPart in sourceSlide.ChartParts)
            {
                string oldRelId = sourceSlide.GetIdOfPart(chartPart);
                ChartPart newChartPart = newSlidePart.AddNewPart<ChartPart>();
                
                // 复制图表内容
                newChartPart.FeedData(chartPart.GetStream());
                
                string newRelId = newSlidePart.GetIdOfPart(newChartPart);
                ReplaceRelationshipId(newSlidePart.Slide, oldRelId, newRelId);
            }

            // 复制嵌入对象（如视频、音频等）
            foreach (var embeddedPackagePart in sourceSlide.EmbeddedPackageParts)
            {
                string oldRelId = sourceSlide.GetIdOfPart(embeddedPackagePart);
                EmbeddedPackagePart newEmbeddedPart = newSlidePart.AddEmbeddedPackagePart(embeddedPackagePart.ContentType);
                
                // 复制嵌入内容
                newEmbeddedPart.FeedData(embeddedPackagePart.GetStream());

                string newRelId = newSlidePart.GetIdOfPart(newEmbeddedPart);
                ReplaceRelationshipId(newSlidePart.Slide, oldRelId, newRelId);
            }

            // 获取当前最大的 SlideId
            var slideIdList = presentationPart.Presentation.SlideIdList;
            uint maxSlideId = 256;
            foreach (var slideId in slideIdList.Elements<P.SlideId>())
            {
                if (slideId.Id != null && slideId.Id.Value > maxSlideId)
                {
                    maxSlideId = slideId.Id.Value;
                }
            }

            // 将新 slide 添加到 presentation 的最后
            var newSlideId = new P.SlideId
            {
                Id = maxSlideId + 1,
                RelationshipId = presentationPart.GetIdOfPart(newSlidePart)
            };

            slideIdList.Append(newSlideId);

            return newSlidePart;
        }

        // 替换 slide 中的 RelationshipId 引用
        private static void ReplaceRelationshipId(OpenXmlElement element, string oldRelId, string newRelId)
        {
            if (element == null)
                return;

            // 定义关系命名空间
            const string relationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            
            // 查找所有可能包含关系 ID 的属性并替换
            var attributes = element.GetAttributes().ToList();
            
            foreach (var attr in attributes)
            {
                // 检查是否是关系命名空间的属性，并且值匹配
                if (attr.Value == oldRelId)
                {
                    // 检查是否是关系属性（r:embed, r:id, r:link 等）
                    if (attr.NamespaceUri == relationshipsNamespace)
                    {
                        element.SetAttribute(new OpenXmlAttribute(attr.Prefix, attr.LocalName, attr.NamespaceUri, newRelId));
                    }
                }
            }

            // 递归处理所有子元素
            foreach (var child in element.Elements())
            {
                ReplaceRelationshipId(child, oldRelId, newRelId);
            }
        }

        private static List<SlidePart> FindSlidesWithKeyword(PresentationPart presentationPart, string keyword, int sectionsCount = -1)
        {
            var matchedSlides = new List<SlidePart>();

            if (presentationPart.Presentation?.SlideIdList == null)
                return matchedSlides;

            foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<P.SlideId>())
            {
                var slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                if (slidePart == null)
                    continue;

                // 首先检查是否包含关键字
                bool containsKeyword = false;
                var allShapes = slidePart.Slide.Descendants<P.Shape>();
                foreach (var shape in allShapes)
                {
                    string text = GetShapeText(shape);
                    if (!string.IsNullOrEmpty(text) && text.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                    {
                        containsKeyword = true;
                        break;
                    }
                }

                if (!containsKeyword)
                    continue;

                // 如果指定了 sectionsCount，检查 slide 是否包含适配数量的 section 占位符
                if (sectionsCount > 0)
                {
                    // 获取整个 slide 的所有文本
                    string slideText = string.Empty;
                    foreach (var shape in allShapes)
                    {
                        slideText += GetShapeText(shape) + " ";
                    }

                    // 精确匹配：必须包含 {section_title_1} 到 {section_title_N}，且不能有 {section_title_N+1}
                    bool hasExactMatch = true;
                    
                    // 检查 1 到 sectionsCount 的所有占位符是否都存在
                    for (int i = 1; i <= sectionsCount; i++)
                    {
                        string sectionPlaceholder = $"{{section_title_{i}}}";
                        if (!slideText.Contains(sectionPlaceholder, StringComparison.OrdinalIgnoreCase))
                        {
                            hasExactMatch = false;
                            break;
                        }
                    }

                    // 检查 sectionsCount + 1 的占位符是否不存在（确保不会有多余的占位符）
                    if (hasExactMatch)
                    {
                        string nextSectionPlaceholder = $"{{section_title_{sectionsCount + 1}}}";
                        if (slideText.Contains(nextSectionPlaceholder, StringComparison.OrdinalIgnoreCase))
                        {
                            hasExactMatch = false;
                        }
                    }

                    if (!hasExactMatch)
                        continue;
                }

                matchedSlides.Add(slidePart);
            }

            return matchedSlides;
        }

        private static int GetSlideIndex(PresentationPart presentationPart, SlidePart slidePart)
        {
            var slideIdList = presentationPart.Presentation?.SlideIdList;
            if (slideIdList == null)
            {
                return -1;
            }

            string? relationshipId = presentationPart.GetIdOfPart(slidePart);
            if (string.IsNullOrEmpty(relationshipId))
            {
                return -1;
            }

            int index = 0;
            foreach (var slideId in slideIdList.Elements<P.SlideId>())
            {
                if (string.Equals(slideId.RelationshipId, relationshipId, StringComparison.Ordinal))
                {
                    return index;
                }
                index++;
            }

            return -1;
        }

        private static void ReplaceFirstSlideWithJson(PresentationPart presentationPart, JsonDocument doc)
        {
            Log(() => "\n=== Replacing First Slide Placeholders with JSON Data ===");
            
            JsonElement root = doc.RootElement;

            // 提取替换值
            var replacements = new Dictionary<string, string>
            {
                { "{ppt_title}", GetJsonString(root, "title") },
                { "{ppt_subtitle}", GetJsonString(root, "subtitle") },
                { "{ppt_author}", GetJsonString(root, "author") },
                { "{ppt_website}", GetJsonString(root, "website") },
            };

            Log(() => "  Replacing values: " + string.Join(", ", replacements.Select(kvp => $"{kvp.Key}='{kvp.Value}'")));

            // 获取第一页
            P.SlideId firstSlideId = presentationPart.Presentation.SlideIdList.ChildElements[0] as P.SlideId;
            if (firstSlideId == null)
            {
                throw new ArgumentException("Cannot find first slide");
            }

            SlidePart firstSlidePart = presentationPart.GetPartById(firstSlideId.RelationshipId) as SlidePart;
            if (firstSlidePart == null)
            {
                throw new ArgumentException("Cannot get first slide part");
            }

            // 替换所有形状中的占位符文本
            int replacedCount = 0;
            var shapes = firstSlidePart.Slide.Descendants<P.Shape>();
            foreach (var shape in shapes)
            {
                replacedCount += ReplaceShapePlaceholders(shape, replacements);
            }
            
            Log(() => $"Successfully replaced {replacedCount} placeholders on first slide");
        }

        private static void ReplaceSecondSlideWithJson(PresentationPart presentationPart, JsonDocument doc)
        {
            Log(() => "\n=== Replacing Second Slide (TOC) with JSON Data ===");

            JsonElement root = doc.RootElement;

            // 获取 parts 数组
            if (!root.TryGetProperty("parts", out JsonElement partsArray))
            {
                throw new ArgumentException("JSON does not contain 'parts' array");
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

            Log(() => $"Found {partTitles.Count} part titles in JSON");

            // 获取第二页
            P.SlideId secondSlideId = presentationPart.Presentation.SlideIdList.ChildElements[1] as P.SlideId;
            if (secondSlideId == null)
            {
                throw new ArgumentException("Cannot find second slide");
            }

            SlidePart secondSlidePart = presentationPart.GetPartById(secondSlideId.RelationshipId) as SlidePart;
            if (secondSlidePart == null)
            {
                throw new ArgumentException("Cannot get second slide part");
            }

            // 查找所有包含 {part_title_x} 占位符的形状
            var shapes = secondSlidePart.Slide.Descendants<P.Shape>().ToList();
            var placeholderShapes = new Dictionary<int, P.Shape>();

            var indexPlaceholderShapes = new Dictionary<int, P.Shape>();

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
                        Log(() => $"Found placeholder '{placeholder}' in shape");
                        break;
                    }
                }
                
                // 查找 {p1}, {p2}, {p3}, ... 等索引占位符
                for (int i = 1; i <= 10; i++)
                {
                    string indexPlaceholder = $"{{p{i}}}";
                    if (shapeText.Contains(indexPlaceholder))
                    {
                        indexPlaceholderShapes[i] = shape;
                        Log(() => $"Found index placeholder '{indexPlaceholder}' in shape");
                        break;
                    }
                }
            }

            Log(() => $"Found {placeholderShapes.Count} placeholder shapes");
            Log(() => $"Found {indexPlaceholderShapes.Count} index placeholder shapes");

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
                    Log(() => $"  Replacing {{part_title_{i}}} with: '{partTitles[titleIndex]}'");
                    ReplaceShapeContent(placeholderShapes[i], partTitles[titleIndex]);
                }
                else if (placeholderShapes.ContainsKey(i) && titleIndex >= partTitles.Count)
                {
                    // 有占位符但没有对应的标题，删除形状
                    Log(() => $"  Deleting extra placeholder shape: {{part_title_{i}}}");
                    DeleteShape(placeholderShapes[i]);
                }
                else if (!placeholderShapes.ContainsKey(i) && titleIndex < partTitles.Count)
                {
                    // 有标题但没有占位符，跳过
                    Log(() => $"  Skipping title (no placeholder): '{partTitles[titleIndex]}'");
                }
            }

            // 处理索引占位符 {p1}, {p2}, {p3}, ...
            int maxIndexPlaceholder = Math.Max(
                indexPlaceholderShapes.Count > 0 ? indexPlaceholderShapes.Keys.Max() : 0,
                partTitles.Count
            );

            for (int i = 1; i <= maxIndexPlaceholder; i++)
            {
                int titleIndex = i - 1; // 标题数组索引从0开始
                
                if (indexPlaceholderShapes.ContainsKey(i) && titleIndex < partTitles.Count)
                {
                    // 有索引占位符且有对应的标题，用索引数字替换
                    Log(() => $"  Replacing {{p{i}}} with: '{i}'");
                    ReplaceShapeContent(indexPlaceholderShapes[i], i.ToString());
                }
                else if (indexPlaceholderShapes.ContainsKey(i) && titleIndex >= partTitles.Count)
                {
                    // 有索引占位符但没有对应的标题，删除形状
                    Log(() => $"  Deleting extra index placeholder shape: {{p{i}}}");
                    DeleteShape(indexPlaceholderShapes[i]);
                }
            }
            
            Log(() => "Second slide updated successfully");
        }

        private static void ReplaceShapeContent(P.Shape shape, string newContent)
        {
            SetShapeText(shape, newContent);
        }

        private static void SetShapeText(P.Shape shape, string newContent)
        {
            if (shape?.TextBody == null)
            {
                return;
            }

            var paragraphs = shape.TextBody.Elements<A.Paragraph>().ToList();
            if (paragraphs.Count == 0)
            {
                var newParagraph = new A.Paragraph();
                shape.TextBody.AppendChild(newParagraph);
                paragraphs.Add(newParagraph);
            }

            var firstParagraph = paragraphs[0];

            var originalParagraphProperties = firstParagraph.Elements<A.ParagraphProperties>().FirstOrDefault();

            A.RunProperties originalRunProperties = null;
            var existingRun = firstParagraph.Elements<A.Run>().FirstOrDefault();
            if (existingRun != null)
            {
                originalRunProperties = existingRun.Elements<A.RunProperties>().FirstOrDefault();
            }

            var originalEndParaRPr = firstParagraph.Elements<A.EndParagraphRunProperties>().FirstOrDefault();

            firstParagraph.RemoveAllChildren();

            if (originalParagraphProperties != null)
            {
                firstParagraph.AppendChild((A.ParagraphProperties)originalParagraphProperties.CloneNode(true));
            }

            var newRun = new A.Run();

            if (originalRunProperties != null)
            {
                newRun.AppendChild((A.RunProperties)originalRunProperties.CloneNode(true));
            }

            var newText = new A.Text(newContent ?? string.Empty);
            newRun.AppendChild(newText);
            firstParagraph.AppendChild(newRun);

            if (originalEndParaRPr != null)
            {
                firstParagraph.AppendChild((A.EndParagraphRunProperties)originalEndParaRPr.CloneNode(true));
            }

            for (int i = 1; i < paragraphs.Count; i++)
            {
                paragraphs[i].Remove();
            }
        }

        private static void DeleteShape(P.Shape shape)
        {
            shape.Remove();
        }

        private static bool TryGetJsonString(JsonElement element, string propertyName, out string value)
        {
            value = string.Empty;

            if (element.ValueKind != JsonValueKind.Object)
            {
                return false;
            }

            if (!element.TryGetProperty(propertyName, out JsonElement propertyValue))
            {
                return false;
            }

            switch (propertyValue.ValueKind)
            {
                case JsonValueKind.String:
                    value = propertyValue.GetString() ?? string.Empty;
                    return true;
                case JsonValueKind.Number:
                case JsonValueKind.True:
                case JsonValueKind.False:
                    value = propertyValue.ToString();
                    return true;
                case JsonValueKind.Null:
                    value = string.Empty;
                    return true;
                default:
                    value = propertyValue.ToString();
                    return true;
            }
        }

        private static string GetJsonString(JsonElement element, string propertyName, string defaultValue = "")
        {
            return TryGetJsonString(element, propertyName, out var value) ? value : defaultValue;
        }

        private static void Measure(string label, Action action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            var stopwatch = Stopwatch.StartNew();
            try
            {
                action();
            }
            finally
            {
                stopwatch.Stop();
                Log(() => $"[Timing] {label} took {stopwatch.ElapsedMilliseconds} ms");
            }
        }

        [Conditional("DEBUG")]
        private static void Log(Func<string> messageFactory)
        {
            if (messageFactory != null)
            {
                Console.WriteLine(messageFactory());
            }
        }

        private sealed class TemplateSelector<T> where T : class
        {
            private readonly Random _random;
            private readonly Queue<T> _pool = new Queue<T>();
            private T? _last;

            public TemplateSelector(Random random)
            {
                _random = random ?? throw new ArgumentNullException(nameof(random));
            }

            public T? GetNext(IList<T> candidates)
            {
                if (candidates == null || candidates.Count == 0)
                {
                    return null;
                }

                var uniqueCandidates = Deduplicate(candidates);
                if (uniqueCandidates.Count == 0)
                {
                    return null;
                }

                var candidateSet = new HashSet<T>(uniqueCandidates);

                if (_last != null && !candidateSet.Contains(_last))
                {
                    _last = null;
                }

                if (_pool.Count == 0)
                {
                    RefillPool(uniqueCandidates, candidateSet);
                }

                var next = DequeueValid(candidateSet);
                if (next == null)
                {
                    _last = null;
                    RefillPool(uniqueCandidates, candidateSet);
                    next = DequeueValid(candidateSet);
                }

                if (next != null)
                {
                    _last = next;
                }

                return next;
            }

            private static List<T> Deduplicate(IList<T> source)
            {
                var result = new List<T>();
                var seen = new HashSet<T>();
                foreach (var item in source)
                {
                    if (item == null)
                    {
                        continue;
                    }

                    if (seen.Add(item))
                    {
                        result.Add(item);
                    }
                }

                return result;
            }

            private void RefillPool(List<T> candidates, HashSet<T> candidateSet)
            {
                if (candidates.Count == 0)
                {
                    return;
                }

                var shuffled = new List<T>(candidates);
                for (int i = shuffled.Count - 1; i > 0; i--)
                {
                    int swapIndex = _random.Next(i + 1);
                    (shuffled[i], shuffled[swapIndex]) = (shuffled[swapIndex], shuffled[i]);
                }

                if (_last != null && shuffled.Count > 1 && EqualityComparer<T>.Default.Equals(shuffled[0], _last))
                {
                    int swapIndex = _random.Next(1, shuffled.Count);
                    (shuffled[0], shuffled[swapIndex]) = (shuffled[swapIndex], shuffled[0]);
                }

                foreach (var item in shuffled)
                {
                    if (candidateSet.Contains(item))
                    {
                        _pool.Enqueue(item);
                    }
                }
            }

            private T? DequeueValid(HashSet<T> candidateSet)
            {
                while (_pool.Count > 0)
                {
                    var candidate = _pool.Dequeue();
                    if (candidate == null || !candidateSet.Contains(candidate))
                    {
                        continue;
                    }

                    if (_last != null && EqualityComparer<T>.Default.Equals(candidate, _last))
                    {
                        if (candidateSet.Count == 1)
                        {
                            return candidate;
                        }

                        _pool.Enqueue(candidate);
                        continue;
                    }

                    return candidate;
                }

                return null;
            }
        }

        private static void DeduplicateMediaResources(PresentationPart presentationPart)
        {
            Log(() => "\n=== Deduplicating Media Resources ===");

            if (presentationPart?.Presentation?.SlideIdList == null)
            {
                Log(() => "Presentation is missing or has no slides");
                return;
            }

            var imageLookup = new Dictionary<string, ImagePart>(StringComparer.OrdinalIgnoreCase);
            int deduplicatedCount = 0;

            foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<P.SlideId>())
            {
                if (slideId == null || string.IsNullOrEmpty(slideId.RelationshipId))
                {
                    continue;
                }

                if (presentationPart.GetPartById(slideId.RelationshipId) is not SlidePart slidePart)
                {
                    continue;
                }

                var imageParts = slidePart.ImageParts.ToList();
                foreach (var imagePart in imageParts)
                {
                    string relationshipId = slidePart.GetIdOfPart(imagePart);
                    string hashKey = GetImageHashKey(imagePart);

                    if (hashKey == null)
                    {
                        continue;
                    }

                    if (!imageLookup.TryGetValue(hashKey, out var canonicalPart))
                    {
                        imageLookup[hashKey] = imagePart;
                        continue;
                    }

                    if (canonicalPart == imagePart)
                    {
                        continue;
                    }

                    string newRelId = null;
                    try
                    {
                        newRelId = slidePart.GetIdOfPart(canonicalPart);
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        // canonicalPart 尚未与当前 slide 建立关系
                    }

                    if (string.IsNullOrEmpty(newRelId))
                    {
                        slidePart.AddPart(canonicalPart);
                        newRelId = slidePart.GetIdOfPart(canonicalPart);
                    }

                    if (string.IsNullOrEmpty(newRelId))
                    {
                        continue;
                    }

                    var sourceUri = imagePart.Uri?.ToString() ?? relationshipId ?? string.Empty;
                    var targetUri = canonicalPart.Uri?.ToString() ?? newRelId;

                    ReplaceRelationshipId(slidePart.Slide, relationshipId, newRelId);
                    slidePart.DeletePart(imagePart);
                    deduplicatedCount++;
                    Log(() => $"  Deduplicated image: {sourceUri} -> {targetUri}");
                }
            }

            Log(() => deduplicatedCount > 0
                ? $"Deduplicated {deduplicatedCount} duplicate images"
                : "No duplicate images detected");
        }

        private static string GetImageHashKey(ImagePart imagePart)
        {
            try
                {
                    // 1. 获取文件流并读取前 N 个字节
                    const int HeaderLength = 4096; // 选取 4KB 头部作为特征
                    
                    using var stream = imagePart.GetStream();
                    
                    if (stream.Length == 0)
                    {
                        return $"{imagePart.ContentType}:Empty";
                    }
                    
                    // 2. 构造一个包含 Size + Header 的字节数组
                    // Header bytes + 8 bytes (for Length)
                    byte[] buffer = new byte[Math.Min(HeaderLength, (int)stream.Length) + 8]; 
                    
                    // 存储文件长度（作为重要的区分特征）
                    byte[] lengthBytes = BitConverter.GetBytes(stream.Length);
                    Array.Copy(lengthBytes, 0, buffer, 0, 8);
                    
                    // 读取头部数据
                    stream.Read(buffer, 8, buffer.Length - 8);

                    // 3. 使用快速非加密哈希算法 (FNV-1a 或 MurmurHash)
                    // C# 内建库没有这些，这里使用简单的非加密哈希（快速但不安全）
                    
                    int hash = Fnv1aHash(buffer);
                    
                    // 4. 返回 Key
                    return $"{imagePart.ContentType}:{hash.ToString("X8")}";
                }
                catch (Exception ex)
                {
                    Log(() => $"    Failed to hash image {imagePart.Uri}: {ex.Message}");
                    return null;
                }
        }

        // FNV-1a Hash 辅助函数（用于快速非加密哈希）
        private static int Fnv1aHash(byte[] data)
        {
            const uint FnvPrime = 16777619;
            const uint FnvOffsetBasis = 2166136261;

            uint hash = FnvOffsetBasis;
            for (int i = 0; i < data.Length; i++)
            {
                hash ^= data[i];
                hash *= FnvPrime;
            }

            return (int)hash;
        }

        private static void CleanupUnusedMediaResources(PresentationPart presentationPart)
        {
            Log(() => "\n=== Cleaning Unused Media Resources ===");

            if (presentationPart?.Presentation?.SlideIdList == null)
            {
                Log(() => "Presentation is missing or has no slides");
                return;
            }

            foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<P.SlideId>())
            {
                if (slideId == null || string.IsNullOrEmpty(slideId.RelationshipId))
                {
                    continue;
                }

                if (presentationPart.GetPartById(slideId.RelationshipId) is not SlidePart slidePart)
                {
                    continue;
                }

                Log(() => $"  Checking slide {slidePart.Uri} for unused resources");
                CleanupUnusedPartsOnSlide(slidePart);
            }
        }

        private static void CleanupUnusedPartsOnSlide(SlidePart slidePart)
        {
            if (slidePart.Slide == null)
            {
                return;
            }

            var referencedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            CollectRelationshipIds(slidePart.Slide, referencedIds);

            int removedCount = 0;

            var childParts = slidePart.Parts.ToList();
            foreach (var child in childParts)
            {
                string relationshipId = child.RelationshipId;
                var part = child.OpenXmlPart;

                if (!IsMediaPart(part))
                {
                    continue;
                }

                if (!referencedIds.Contains(relationshipId))
                {
                    string? partUri = SafeGetPartUri(part);
                    slidePart.DeletePart(part);
                    removedCount++;
                    Log(() => partUri != null
                        ? $"    Removed unused media part {partUri}"
                        : $"    Removed unused media part (relationship {relationshipId})");
                }
            }

            var dataReferences = slidePart.DataPartReferenceRelationships.ToList();
            foreach (var dataReference in dataReferences)
            {
                if (!referencedIds.Contains(dataReference.Id))
                {
                    string? dataUri = dataReference.Uri?.ToString();
                    slidePart.DeleteReferenceRelationship(dataReference);
                    removedCount++;
                    Log(() => dataUri != null
                        ? $"    Removed unused data reference {dataUri}"
                        : $"    Removed unused data reference (relationship {dataReference.Id})");
                }
            }

            if (removedCount == 0)
            {
                Log(() => "    No unused media resources detected");
            }
        }

        private static void CollectRelationshipIds(OpenXmlElement element, HashSet<string> referencedIds)
        {
            if (element == null)
            {
                return;
            }

            const string relationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            foreach (var attribute in element.GetAttributes())
            {
                if (attribute.NamespaceUri == relationshipsNamespace && !string.IsNullOrEmpty(attribute.Value))
                {
                    referencedIds.Add(attribute.Value);
                }
            }

            foreach (var child in element.Elements())
            {
                CollectRelationshipIds(child, referencedIds);
            }
        }

        private static bool IsMediaPart(OpenXmlPart part)
        {
            return part is ImagePart
                || part is ChartPart
                || part is EmbeddedPackagePart
                || part is EmbeddedObjectPart
                || part is VmlDrawingPart;
        }

        private static string? SafeGetPartUri(OpenXmlPart? part)
        {
            if (part == null)
            {
                return null;
            }

            try
            {
                return part.Uri?.ToString();
            }
            catch (ObjectDisposedException)
            {
                return null;
            }
            catch (InvalidOperationException)
            {
                return null;
            }
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
                    Log(() => $"  Found placeholder '{kvp.Key}' in shape");
                    Log(() => $"  Original text: '{shapeFullText.Trim()}'");
                    Log(() => $"  Replacing entire content with: '{kvp.Value}'");

                    SetShapeText(shape, kvp.Value);

                    count++;
                    break; // 一个形状只替换一次
                }
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

        private static async Task RunServerAsync(string? prefix)
        {
            string listeningPrefix = string.IsNullOrWhiteSpace(prefix) ? DefaultPrefix : prefix;
            if (!listeningPrefix.EndsWith("/", StringComparison.Ordinal))
            {
                listeningPrefix += "/";
            }

            Directory.CreateDirectory(CacheDirectory);
            Directory.CreateDirectory(GetOutputDirectory());

            var listener = new HttpListener();
            listener.Prefixes.Add(listeningPrefix);

            try
            {
                listener.Start();
            }
            catch (HttpListenerException ex)
            {
                Console.Error.WriteLine($"Failed to start server on {listeningPrefix}: {ex.Message}");
                return;
            }

            Console.WriteLine($"AddNamedSheetView server listening on {listeningPrefix}");

            while (listener.IsListening)
            {
                HttpListenerContext context;
                try
                {
                    context = await listener.GetContextAsync().ConfigureAwait(false);
                }
                catch (HttpListenerException)
                {
                    break;
                }
                catch (ObjectDisposedException)
                {
                    break;
                }

                _ = Task.Run(() => HandleRequestAsync(context));
            }

            listener.Close();
        }

        private static async Task HandleRequestAsync(HttpListenerContext context)
        {
            try
            {
                string path = context.Request.Url?.AbsolutePath ?? "/";

                if (string.Equals(path, "/", StringComparison.Ordinal))
                {
                    await WriteJsonAsync(context.Response, new
                    {
                        code = 0,
                        data = new
                        {
                            message = "AddNamedSheetView server is running.",
                            endpoints = new[] { "/gen?tpl=...&uid=...&data=...", "/files/{path}" }
                        }
                    }).ConfigureAwait(false);
                    return;
                }

                if (path.StartsWith("/files/", StringComparison.OrdinalIgnoreCase))
                {
                    await ServeFileAsync(context).ConfigureAwait(false);
                    return;
                }

                if (path.StartsWith("/gen", StringComparison.OrdinalIgnoreCase))
                {
                    await HandleGenerateAsync(context).ConfigureAwait(false);
                    return;
                }

                await WriteJsonAsync(context.Response, new
                {
                    code = 404,
                    data = new { msg = "Not Found" }
                }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                PrintServerError("HandleRequestAsync", ex);
                await WriteJsonAsync(context.Response, new
                {
                    code = 500,
                    data = new { msg = ex.Message }
                }).ConfigureAwait(false);
            }
            finally
            {
                context.Response.Close();
            }
        }

        private static async Task HandleGenerateAsync(HttpListenerContext context)
        {
            var request = context.Request;

            if (!string.Equals(request.HttpMethod, "GET", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(request.HttpMethod, "POST", StringComparison.OrdinalIgnoreCase))
            {
                await WriteJsonAsync(context.Response, new
                {
                    code = 405,
                    data = new { msg = "Only GET or POST supported" }
                }).ConfigureAwait(false);
                return;
            }

            string? tplUrl = request.QueryString["tpl"] ?? request.QueryString["tpl_url"];
            if (string.IsNullOrWhiteSpace(tplUrl))
            {
                await WriteJsonAsync(context.Response, new
                {
                    code = 400,
                    data = new { msg = "tpl query parameter is required" }
                }).ConfigureAwait(false);
                return;
            }

            string outlineRaw = await ResolveOutlineAsync(request).ConfigureAwait(false);
            string outlineJson = NormalizeOutlinePayload(outlineRaw);
            Console.WriteLine($"[Server] Received outline (raw length={outlineRaw?.Length ?? 0}): {TruncateForLog(outlineRaw, 512)}");
            Console.WriteLine($"[Server] Outline decoded (length={outlineJson?.Length ?? 0}): {TruncateForLog(outlineJson, 512)}");

            if (string.IsNullOrWhiteSpace(outlineJson))
            {
                await WriteJsonAsync(context.Response, new
                {
                    code = 400,
                    data = new { msg = "outline data is required" }
                }).ConfigureAwait(false);
                return;
            }

            string uid = request.QueryString["uid"];
            if (string.IsNullOrWhiteSpace(uid))
            {
                uid = Guid.NewGuid().ToString("N");
            }

            var stopwatch = Stopwatch.StartNew();

            try
            {
                string templatePath = await GetCachedTemplateAsync(tplUrl).ConfigureAwait(false);
                string outlinePath = await PersistOutlineAsync(uid, outlineJson).ConfigureAwait(false);
                (string outputPath, string downloadUrl) = await GeneratePresentationAsync(uid, templatePath, outlinePath, context.Request.Url).ConfigureAwait(false);

                stopwatch.Stop();

                await WriteJsonAsync(context.Response, new
                {
                    code = 0,
                    data = new
                    {
                        ppt_path = outputPath,
                        ppt_url = downloadUrl,
                        elap = Math.Round(stopwatch.Elapsed.TotalMilliseconds, 2)
                    }
                }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                PrintServerError("HandleGenerateAsync", ex);
                stopwatch.Stop();
                await WriteJsonAsync(context.Response, new
                {
                    code = 500,
                    data = new { msg = ex.Message }
                }).ConfigureAwait(false);
            }
        }

        private static async Task ServeFileAsync(HttpListenerContext context)
        {
            string path = context.Request.Url?.AbsolutePath ?? string.Empty;
            string relativePath = path["/files/".Length..];

            if (string.IsNullOrWhiteSpace(relativePath))
            {
                await WriteJsonAsync(context.Response, new
                {
                    code = 400,
                    data = new { msg = "file path is required" }
                }).ConfigureAwait(false);
                return;
            }

            string safePath = relativePath.Replace('/', Path.DirectorySeparatorChar);
            string targetPath = Path.Combine(AppDataDirectory, safePath);
            string fullPath = Path.GetFullPath(targetPath);
            string appDataRoot = Path.GetFullPath(AppDataDirectory);

            if (!fullPath.StartsWith(appDataRoot, StringComparison.OrdinalIgnoreCase))
            {
                await WriteJsonAsync(context.Response, new
                {
                    code = 403,
                    data = new { msg = "access denied" }
                }).ConfigureAwait(false);
                return;
            }

            if (!File.Exists(fullPath))
            {
                await WriteJsonAsync(context.Response, new
                {
                    code = 404,
                    data = new { msg = "file not found" }
                }).ConfigureAwait(false);
                return;
            }

            context.Response.ContentType = GetContentType(fullPath);
            context.Response.StatusCode = 200;

            await using var fs = File.OpenRead(fullPath);
            await fs.CopyToAsync(context.Response.OutputStream).ConfigureAwait(false);
        }

        private static async Task<string> ResolveOutlineAsync(HttpListenerRequest request)
        {
            string? outline = request.QueryString["data"] ?? request.QueryString["outline"];
            if (!string.IsNullOrEmpty(outline))
            {
                return outline;
            }

            if (request.HasEntityBody)
            {
                using var reader = new StreamReader(request.InputStream, request.ContentEncoding ?? Encoding.UTF8);
                string body = await reader.ReadToEndAsync().ConfigureAwait(false);
                if (string.IsNullOrWhiteSpace(body))
                {
                    return string.Empty;
                }

                if (IsJson(request.ContentType))
                {
                    try
                    {
                        using var doc = JsonDocument.Parse(body);
                        if (doc.RootElement.TryGetProperty("outline", out JsonElement outlineElement))
                        {
                            return outlineElement.GetString() ?? outlineElement.ToString();
                        }

                        if (doc.RootElement.TryGetProperty("data", out JsonElement dataElement))
                        {
                            return dataElement.GetString() ?? dataElement.ToString();
                        }
                    }
                    catch (JsonException)
                    {
                        return body;
                    }
                }

                return body;
            }

            return string.Empty;
        }

        private static async Task<string> GetCachedTemplateAsync(string tplUrl)
        {
            if (string.IsNullOrWhiteSpace(tplUrl))
            {
                throw new ArgumentException("Template URL is required", nameof(tplUrl));
            }

            if (File.Exists(tplUrl))
            {
                return Path.GetFullPath(tplUrl);
            }

            if (!Uri.TryCreate(tplUrl, UriKind.RelativeOrAbsolute, out Uri? uri))
            {
                throw new InvalidOperationException($"Invalid template URL: {tplUrl}");
            }

            if (!uri.IsAbsoluteUri)
            {
                string localPath = Path.GetFullPath(tplUrl);
                if (!File.Exists(localPath))
                {
                    throw new FileNotFoundException($"Template file not found: {tplUrl}", localPath);
                }

                return localPath;
            }

            if (uri.IsFile)
            {
                string localFile = uri.LocalPath;
                if (!File.Exists(localFile))
                {
                    throw new FileNotFoundException($"Template file not found: {tplUrl}", localFile);
                }

                return localFile;
            }

            string tplHash = Convert.ToHexString(MD5.HashData(Encoding.UTF8.GetBytes(tplUrl)));
            string tplPath = Path.Combine(CacheDirectory, $"{tplHash}.pptx");

            if (File.Exists(tplPath))
            {
                return tplPath;
            }

            var semaphore = TemplateLocks.GetOrAdd(tplPath, _ => new SemaphoreSlim(1, 1));
            await semaphore.WaitAsync().ConfigureAwait(false);
            try
            {
                if (File.Exists(tplPath))
                {
                    return tplPath;
                }

                using HttpResponseMessage response = await HttpClient.GetAsync(uri).ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException($"Failed to download template: {(int)response.StatusCode} {response.ReasonPhrase}");
                }

                Directory.CreateDirectory(Path.GetDirectoryName(tplPath)!);
                string tempPath = tplPath + ".tmp_" + Guid.NewGuid().ToString("N");

                try
                {
                    await using (var fs = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true))
                    {
                        await response.Content.CopyToAsync(fs).ConfigureAwait(false);
                    }

                    File.Move(tempPath, tplPath, overwrite: true);
                }
                catch
                {
                    if (File.Exists(tempPath))
                    {
                        File.Delete(tempPath);
                    }
                    throw;
                }
            }
            finally
            {
                semaphore.Release();
            }

            return tplPath;
        }

        private static async Task<string> PersistOutlineAsync(string uid, string outlineJson)
        {
            string directory = GetOutputDirectory();
            string outlineFileName = CreateUniqueFileName(uid, "json");
            string outlinePath = Path.Combine(directory, outlineFileName);
            await File.WriteAllTextAsync(outlinePath, outlineJson).ConfigureAwait(false);
            return outlinePath;
        }

        private static Task<(string outputPath, string downloadUrl)> GeneratePresentationAsync(string uid, string templatePath, string outlinePath, Uri? requestUri)
        {
            string directory = GetOutputDirectory();
            string pptFileName = CreateUniqueFileName(uid, "pptx");
            string destination = Path.Combine(directory, pptFileName);

            GeneratePresentation(templatePath, outlinePath, destination);

            string downloadUrl = BuildDownloadUrl(requestUri, pptFileName);
            return Task.FromResult((destination, downloadUrl));
        }

        private static string GetOutputDirectory()
        {
            string output = Path.Combine(AppDataDirectory, OutputDirectoryName);
            Directory.CreateDirectory(output);
            return output;
        }

        private static string BuildDownloadUrl(Uri? requestUri, string fileName)
        {
            if (requestUri == null)
            {
                return Path.Combine(AppDataDirectory, OutputDirectoryName, fileName);
            }

            string baseUri = $"{requestUri.Scheme}://{requestUri.Authority}";
            return $"{baseUri}/files/{OutputDirectoryName}/{Uri.EscapeDataString(fileName)}";
        }

        private static async Task WriteJsonAsync(HttpListenerResponse response, object payload)
        {
            response.ContentType = "application/json";
            response.StatusCode = 200;
            byte[] buffer = JsonSerializer.SerializeToUtf8Bytes(payload, JsonOptions);
            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length).ConfigureAwait(false);
        }

        private static bool IsJson(string? contentType)
        {
            if (string.IsNullOrWhiteSpace(contentType))
            {
                return false;
            }

            return contentType.Contains("application/json", StringComparison.OrdinalIgnoreCase)
                || contentType.Contains("text/json", StringComparison.OrdinalIgnoreCase);
        }

        private static string GetContentType(string fullPath)
        {
            string extension = Path.GetExtension(fullPath).ToLowerInvariant();

            return extension switch
            {
                ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                ".json" => "application/json",
                _ => "application/octet-stream",
            };
        }

        private static string TruncateForLog(string? value, int maxLength)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            if (value.Length <= maxLength)
            {
                return value;
            }

            return value.Substring(0, maxLength) + "...";
        }

        private static string CreateUniqueFileName(string? prefix, string extension)
        {
            string safePrefix = SanitizeFileNameComponent(prefix);
            string timestamp = DateTime.UtcNow.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture);
            string random = Guid.NewGuid().ToString("N")[..8];
            string sanitizedExtension = string.IsNullOrWhiteSpace(extension) ? "dat" : extension.TrimStart('.');
            return $"{safePrefix}_{timestamp}_{random}.{sanitizedExtension}";
        }

        private static string SanitizeFileNameComponent(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "item";
            }

            var sb = new StringBuilder(value.Length);
            foreach (char c in value)
            {
                sb.Append(InvalidFileNameChars.Contains(c) ? '_' : c);
            }

            string result = sb.ToString().Trim('_');
            return string.IsNullOrWhiteSpace(result) ? "item" : result;
        }

        private static string NormalizeOutlinePayload(string? payload)
        {
            if (string.IsNullOrWhiteSpace(payload))
            {
                return string.Empty;
            }

            string trimmed = payload.Trim();
            if (IsLikelyJson(trimmed))
            {
                return trimmed;
            }

            if (trimmed.Contains('&') || trimmed.Contains('='))
            {
                string[] pairs = trimmed.Split('&', StringSplitOptions.RemoveEmptyEntries);
                foreach (string pair in pairs)
                {
                    string[] kv = pair.Split('=', 2);
                    if (kv.Length != 2)
                    {
                        continue;
                    }

                    if (kv[0].Equals("outline", StringComparison.OrdinalIgnoreCase) ||
                        kv[0].Equals("data", StringComparison.OrdinalIgnoreCase))
                    {
                        string decodedValue = WebUtility.UrlDecode(kv[1]);
                        if (!string.IsNullOrWhiteSpace(decodedValue))
                        {
                            decodedValue = decodedValue.Trim();
                            if (IsLikelyJson(decodedValue))
                            {
                                return decodedValue;
                            }

                            trimmed = decodedValue;
                        }
                    }
                }
            }

            if (trimmed.Contains('%') || trimmed.Contains('+'))
            {
                string decoded = WebUtility.UrlDecode(trimmed).Trim();
                if (!string.IsNullOrWhiteSpace(decoded))
                {
                    if (IsLikelyJson(decoded))
                    {
                        return decoded;
                    }

                    trimmed = decoded;
                }
            }

            return trimmed;
        }

        private static bool IsLikelyJson(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            char first = value.TrimStart()[0];
            return first == '{' || first == '[';
        }

        private static void PrintServerError(string context, Exception ex)
        {
            Console.Error.WriteLine($"[Server][Error][{context}] {ex}");
        }
    }
}
