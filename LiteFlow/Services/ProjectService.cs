using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Text.Json;
using LiteFlow.Models;

namespace LiteFlow.Services
{
    public static class ProjectService
    {
        public static void SaveProject(string path, LiteFlowProjectData projectData, List<EvidenceItem> items)
        {
            projectData.Steps.Clear();
            foreach (var item in items)
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    item.Image.Save(ms, ImageFormat.Png);
                    projectData.Steps.Add(new EvidenceData { ImageDataBase64 = Convert.ToBase64String(ms.ToArray()), Note = item.Note, TextBelowImage = item.TextBelowImage });
                }
            }
            File.WriteAllText(path, JsonSerializer.Serialize(projectData));
        }

        public static LiteFlowProjectData? LoadProject(string path)
        {
            var json = File.ReadAllText(path);
            return JsonSerializer.Deserialize<LiteFlowProjectData>(json);
        }
    }
}