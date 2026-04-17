using LiteFlow.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace LiteFlow.Services
{
    public static class ExportService
    {
        private static string GetStateHash(LiteFlowProjectData project, List<EvidenceItem> items, Dictionary<string, string> tags)
        {
            using (var sha256 = SHA256.Create())
            {
                string rawData = string.Join("", tags.Values) + items.Count.ToString() + project.ReportLayout.ToString() + project.MobileColumns;
                foreach (var i in items) rawData += i.Note + (i.TextBelowImage ? "B" : "T");

                byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(rawData));
                return BitConverter.ToString(hashBytes).Replace("-", "").Substring(0, 16);
            }
        }

        private static string GetDynamicCachePath(string hash)
        {
            return Path.Combine(Path.GetTempPath(), $"liteflow_cache_{hash}.docx");
        }

        private static void CleanOldCacheFiles()
        {
            try
            {
                var tempDir = Path.GetTempPath();
                var cacheFiles = Directory.GetFiles(tempDir, "liteflow_cache_*.docx");
                foreach (var file in cacheFiles)
                {
                    if (File.GetLastWriteTime(file) < DateTime.Now.AddDays(-1)) { try { File.Delete(file); } catch { } }
                }
            }
            catch { }
        }

        public static void ExportToWord(LiteFlowProjectData project, List<EvidenceItem> items, string userChosenPath, Dictionary<string, string> tags)
        {
            string currentHash = GetStateHash(project, items, tags);
            string cachePath = GetDynamicCachePath(currentHash);

            if (!File.Exists(cachePath)) GenerateWordToCache(project, items, tags, cachePath);

            if (File.Exists(userChosenPath)) File.Delete(userChosenPath);
            File.Copy(cachePath, userChosenPath);
        }

        public static void ExportToPdf(LiteFlowProjectData project, List<EvidenceItem> items, string userChosenPath, Dictionary<string, string> tags)
        {
            string currentHash = GetStateHash(project, items, tags);
            string cachePath = GetDynamicCachePath(currentHash);

            if (!File.Exists(cachePath)) GenerateWordToCache(project, items, tags, cachePath);

            ConvertDocxToPdf(cachePath, userChosenPath);
        }

        private static void GenerateWordToCache(LiteFlowProjectData project, List<EvidenceItem> items, Dictionary<string, string> tags, string cachePath)
        {
            CleanOldCacheFiles();
            if (File.Exists(cachePath)) File.Delete(cachePath);

            WordDocumentEngine.PrepareDocument(project.TemplatePath, cachePath, tags);
            WordDocumentEngine.AppendAllEvidence(items, cachePath, project);
        }

        private static void ConvertDocxToPdf(string docxPath, string pdfPath)
        {
            try
            {
                Type wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType != null)
                {
                    dynamic wordApp = Activator.CreateInstance(wordType);
                    try
                    {
                        wordApp.Visible = false;
                        dynamic doc = wordApp.Documents.Open(docxPath);
                        doc.ExportAsFixedFormat(pdfPath, 17);
                        doc.Close(false);
                        return;
                    }
                    finally { wordApp.Quit(); }
                }
            }
            catch { }

            string[] loPaths = {
                @"C:\Program Files\LibreOffice\program\soffice.exe",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            };

            string sofficePath = null!;
            foreach (var path in loPaths) { if (File.Exists(path)) { sofficePath = path; break; } }

            if (sofficePath != null)
            {
                var proc = new System.Diagnostics.Process();
                proc.StartInfo.FileName = sofficePath;
                proc.StartInfo.Arguments = $"--headless --convert-to pdf \"{docxPath}\" --outdir \"{Path.GetDirectoryName(pdfPath)}\"";
                proc.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                proc.StartInfo.CreateNoWindow = true;
                proc.Start();
                proc.WaitForExit();

                string loOutput = Path.Combine(Path.GetDirectoryName(pdfPath)!, Path.GetFileNameWithoutExtension(docxPath) + ".pdf");
                if (loOutput != pdfPath && File.Exists(loOutput))
                {
                    if (File.Exists(pdfPath)) File.Delete(pdfPath);
                    File.Move(loOutput, pdfPath);
                }
                return;
            }
            throw new Exception("O LiteFlow necessita do Microsoft Word ou do LibreOffice instalados neste computador para garantir que o PDF sai com a formatação exata e perfeita do seu Template.");
        }
    }
}
