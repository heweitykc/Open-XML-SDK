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
using System.Security.Cryptography;
using System.Diagnostics;

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
        public static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Common.ExampleUtilities.ShowHelp(new string[]
                {
                    "NamedSheetView: ",
                    "Usage: NamedSheetView <filename> [jsonfile]",
                    "Where: <filename> is the .xlsx file in which to add a named sheet view.",
                    "       or .pptx file to copy slide 2 and insert at the end.",
                    "       [jsonfile] (optional) JSON file with PPT outline data to replace first slide placeholders.",
                });
                return;
            }

                try
                {
                    string outputPath = GeneratePresentation(args[0], args[1]);
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

            public static string GeneratePresentation(string sourceFilePath, string jsonFilePath)
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
                string outputPath = GenerateOutputPath(sourceFilePath);

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
                        ReplaceFirstSlideWithJson(ensuredPresentationPart, doc);
                        ReplaceSecondSlideWithJson(ensuredPresentationPart, doc);

                        // 生成各个 part 的 slides
                        GeneratePartSlidesFromJson(ensuredPresentationPart, doc);

                        // 从原始模板复制最后一页并替换结束页占位符
                        CopyAndReplaceLastSlideFromTemplate(ensuredPresentationPart, originalLastSlideIndex, doc);

                        //删除从索引 [2 - $originalLastSlideIndex] 的所有slides
                        DeleteSlidesFromIndex(ensuredPresentationPart, 2, originalLastSlideIndex);

                        // 媒体资源去重与清理
                        DeduplicateMediaResources(ensuredPresentationPart);
                        CleanupUnusedMediaResources(ensuredPresentationPart);
                    }

                    ensuredPresentationPart.Presentation!.Save();
                }

                Log(() => $"{outputPath} saved");
                return outputPath;
            }

        private static string GenerateOutputPath(string inputPath)
        {
            string directory = Path.GetDirectoryName(inputPath) ?? string.Empty;
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
            string extension = Path.GetExtension(inputPath);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            // 创建 output 目录
            string outputDirectory = Path.Combine(directory, "output");
             if (!Directory.Exists(outputDirectory))
             {
                 Directory.CreateDirectory(outputDirectory);
                 Log(() => $"Created output directory: {outputDirectory}");
             }

            string outputFileName = $"{fileNameWithoutExt}_modified_{timestamp}{extension}";
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
                Log(() => "No slides found in presentation");
                return;
            }

            // 确保原始索引有效
            if (originalLastSlideIndex < 0 || originalLastSlideIndex >= slideIdList.ChildElements.Count)
            {
                Log(() => $"Invalid original slide index: {originalLastSlideIndex}");
                return;
            }

            // 获取原始模板的最后一个 slide (使用记录的索引)
            P.SlideId templateLastSlideId = slideIdList.ChildElements[originalLastSlideIndex] as P.SlideId;
            if (templateLastSlideId == null)
            {
                Log(() => "Cannot find template's last slide");
                return;
            }

            string? templateLastRelationshipId = templateLastSlideId.RelationshipId;
            if (string.IsNullOrEmpty(templateLastRelationshipId))
            {
                Log(() => "Template last slide relationship ID is null");
                return;
            }

            SlidePart templateLastSlidePart = presentationPart.GetPartById(templateLastRelationshipId) as SlidePart;
            if (templateLastSlidePart == null)
            {
                Log(() => "Cannot get template's last slide part");
                return;
            }

            Log(() => $"Found template's last slide (index {originalLastSlideIndex}), copying...");

            // 复制原始模板的最后一页
            SlidePart newSlidePart = CopySlide(presentationPart, templateLastSlidePart);
            if (newSlidePart == null)
            {
                Log(() => "Failed to copy template's last slide");
                return;
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
                Log(() => "JSON does not contain 'parts' array");
                return;
            }

            // 查找所有包含 {part_subtitle_ 的 slides
            var templateSlides = FindSlidesWithKeyword(presentationPart, chapter_subtitle_search);
            
            if (templateSlides.Count == 0)
            {
                Log(() => "No slides found containing '{part_subtitle_' placeholder");
                return;
            }

            Log(() => $"Found {templateSlides.Count} template slides containing '{{part_subtitle_}}'");

            var random = new Random();
            var partTemplatePool = new Queue<SlidePart>();
            SlidePart? lastPartTemplate = null;
            int partIndex = 1;

            // 遍历 JSON 中的每个 part
            foreach (var part in partsArray.EnumerateArray())
            {
                if (!part.TryGetProperty("title", out JsonElement title))
                {
                    Log(() => $"Part {partIndex} missing 'title' field, skipping");
                    partIndex++;
                    continue;
                }

                string partTitle = title.GetString() ?? string.Empty;
                Log(() => $"\n--- Processing Part {partIndex}: {partTitle} ---");

                // 从模板 slides 中随机选择一个
                var selectedTemplate = GetNextTemplate(templateSlides, partTemplatePool, random, ref lastPartTemplate);
                if (selectedTemplate == null)
                {
                    Log(() => "  Failed to select template slide, skipping");
                    partIndex++;
                    continue;
                }
                Log(() => $"  Selected random template slide");

                // 复制 slide 并插入到最后
                var newSlidePart = CopySlide(presentationPart, selectedTemplate);
                if (newSlidePart == null)
                {
                    Log(() => $"  Failed to copy slide for part {partIndex}");
                    partIndex++;
                    continue;
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
            var chapterTemplatePool = new Queue<SlidePart>();
            SlidePart? lastChapterTemplate = null;
            int chapterIndex = 1;

            // 遍历每个 chapter
            foreach (var chapter in chaptersArray.EnumerateArray())
            {
                if (!chapter.TryGetProperty("title", out JsonElement chapterTitle))
                {
                    Log(() => $"    Chapter {chapterIndex} missing 'title' field, skipping");
                    chapterIndex++;
                    continue;
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
                
                if (chapterTemplateSlides.Count == 0)
                {
                    Log(() => $"      No slides found matching sections count {sectionsCount}, trying without section filter");
                    // 如果没有找到适配的 slide，尝试不使用 section 筛选
                    chapterTemplateSlides = FindSlidesWithKeyword(presentationPart, "{chapter_title}");
                    
                    if (chapterTemplateSlides.Count == 0)
                    {
                        Log(() => "      No slides found containing '{chapter_title}' placeholder, skipping");
                        chapterIndex++;
                        continue;
                    }
                }

                Log(() => $"      Found {chapterTemplateSlides.Count} matching chapter template slides");

                // 从模板 slides 中随机选择一个
                var selectedTemplate = GetNextTemplate(chapterTemplateSlides, chapterTemplatePool, random, ref lastChapterTemplate);
                if (selectedTemplate == null)
                {
                    Log(() => "      Failed to select chapter template slide, skipping");
                    chapterIndex++;
                    continue;
                }
                Log(() => $"      Selected random chapter template slide");

                // 复制 slide 并插入到最后
                var newSlidePart = CopySlide(presentationPart, selectedTemplate);
                if (newSlidePart == null)
                {
                    Log(() => $"      Failed to copy slide for chapter {chapterIndex}");
                    chapterIndex++;
                    continue;
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
                Log(() => "Cannot find first slide");
                return;
            }

            SlidePart firstSlidePart = presentationPart.GetPartById(firstSlideId.RelationshipId) as SlidePart;
            if (firstSlidePart == null)
            {
                Log(() => "Cannot get first slide part");
                return;
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
                Log(() => "JSON does not contain 'parts' array");
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

            Log(() => $"Found {partTitles.Count} part titles in JSON");

            // 获取第二页
            P.SlideId secondSlideId = presentationPart.Presentation.SlideIdList.ChildElements[1] as P.SlideId;
            if (secondSlideId == null)
            {
                Log(() => "Cannot find second slide");
                return;
            }

            SlidePart secondSlidePart = presentationPart.GetPartById(secondSlideId.RelationshipId) as SlidePart;
            if (secondSlidePart == null)
            {
                Log(() => "Cannot get second slide part");
                return;
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

        [Conditional("DEBUG")]
        private static void Log(Func<string> messageFactory)
        {
            if (messageFactory != null)
            {
                Console.WriteLine(messageFactory());
            }
        }

        private static T? GetNextTemplate<T>(IList<T> source, Queue<T> pool, Random random, ref T? lastItem)
            where T : class
        {
            if (source == null || source.Count == 0)
            {
                return null;
            }

            if (pool.Count == 0)
            {
                var shuffled = ShuffleWithoutStartingRepeat(source, random, lastItem);
                foreach (var item in shuffled)
                {
                    pool.Enqueue(item);
                }
            }

            if (pool.Count == 0)
            {
                return null;
            }

            var next = pool.Dequeue();
            lastItem = next;
            return next;
        }

        private static List<T> ShuffleWithoutStartingRepeat<T>(IList<T> source, Random random, T? lastItem)
            where T : class
        {
            var shuffled = new List<T>(source);

            for (int i = shuffled.Count - 1; i > 0; i--)
            {
                int swapIndex = random.Next(i + 1);
                var temp = shuffled[i];
                shuffled[i] = shuffled[swapIndex];
                shuffled[swapIndex] = temp;
            }

            if (lastItem != null && shuffled.Count > 1 && EqualityComparer<T>.Default.Equals(shuffled[0], lastItem))
            {
                for (int i = 1; i < shuffled.Count; i++)
                {
                    if (!EqualityComparer<T>.Default.Equals(shuffled[i], lastItem))
                    {
                        var temp = shuffled[0];
                        shuffled[0] = shuffled[i];
                        shuffled[i] = temp;
                        break;
                    }
                }
            }

            return shuffled;
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
                using var sha256 = SHA256.Create();
                using var stream = imagePart.GetStream();
                var hashBytes = sha256.ComputeHash(stream);
                var hash = Convert.ToHexString(hashBytes);
                return $"{imagePart.ContentType}:{hash}";
            }
            catch (Exception ex)
            {
                Log(() => $"    Failed to hash image {imagePart.Uri}: {ex.Message}");
                return null;
            }
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
                    slidePart.DeletePart(part);
                    removedCount++;
                    Log(() => $"    Removed unused media part {part.Uri}");
                }
            }

            var dataReferences = slidePart.DataPartReferenceRelationships.ToList();
            foreach (var dataReference in dataReferences)
            {
                if (!referencedIds.Contains(dataReference.Id))
                {
                    slidePart.DeleteReferenceRelationship(dataReference);
                    removedCount++;
                    Log(() => $"    Removed unused data reference {dataReference.Uri}");
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
    }
}
