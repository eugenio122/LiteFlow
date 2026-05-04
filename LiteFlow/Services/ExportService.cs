using LiteFlow.Models;
using System;
using System.Collections.Generic;
using System.IO;

namespace LiteFlow.Services
{
    public static class ExportService
    {
        public static void ExportToWord(LiteFlowProjectData project, List<EvidenceItem> items, string userChosenPath, Dictionary<string, string> tags)
        {
            // Garante que não há conflito se o utilizador for sobrescrever um ficheiro existente
            if (File.Exists(userChosenPath))
            {
                File.Delete(userChosenPath);
            }

            // Gera o Word diretamente no caminho final (sem caches pelo meio)
            WordDocumentEngine.PrepareDocument(project.TemplatePath, userChosenPath, tags);
            WordDocumentEngine.AppendAllEvidence(items, userChosenPath, project);
        }

        public static void ExportToPdf(LiteFlowProjectData project, List<EvidenceItem> items, string userChosenPath, Dictionary<string, string> tags)
        {
            // 1. Gera um nome ÚNICO para o ficheiro temporário, garantindo que NUNCA usa cache
            string tempWordFile = Path.Combine(Path.GetTempPath(), $"ExportTemp_{Guid.NewGuid():N}.docx");

            try
            {
                // 2. FORÇA a criação de um Word fresquinho com os dados da UI mais recentes
                WordDocumentEngine.PrepareDocument(project.TemplatePath, tempWordFile, tags);
                WordDocumentEngine.AppendAllEvidence(items, tempWordFile, project);

                // 3. Converte o arquivo recém-gerado para PDF
                ConvertDocxToPdf(tempWordFile, userChosenPath);
            }
            finally
            {
                // 4. Limpeza imediata! Exclui o arquivo temporário independentemente de erro o sucesso
                if (File.Exists(tempWordFile))
                {
                    try { File.Delete(tempWordFile); } catch { }
                }
            }
        }

        // TORNADO PÚBLICO: Permite que a interface converta um DOCX já existente para poupar CPU
        public static void ConvertDocxToPdf(string docxPath, string pdfPath)
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