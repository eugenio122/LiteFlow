using System;
using System.Windows.Forms;

namespace LiteFlow
{
    internal static class Program
    {
        /// <summary>
        /// O ponto de entrada principal para a aplicação.
        /// Usado APENAS para testes isolados durante o desenvolvimento do plugin.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Inicialização padrão do Windows Forms
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // 1. Criar a Janela "Nave-Mãe" Standalone
            Form standaloneHostForm = new Form
            {
                Text = "LiteFlow - Editor de Evidências (Standalone)",
                Width = 1280,
                Height = 800,
                StartPosition = FormStartPosition.CenterScreen,
                Icon = SystemIcons.Application // Se tiver um .ico do LiteFlow, coloque aqui
            };

            // 2. Instanciar a UI principal (O núcleo da ferramenta)
            LiteFlowUI mainEditor = new LiteFlowUI
            {
                Dock = DockStyle.Fill
            };

            // 3. Adicionar o editor à janela e correr
            standaloneHostForm.Controls.Add(mainEditor);
            Application.Run(standaloneHostForm);
        }
    }
}