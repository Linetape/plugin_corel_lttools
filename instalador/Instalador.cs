// LT Tools - Instalador (Line Tape Iluminação e Locação)
using System;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("LT Tools - Instalador")]
[assembly: AssemblyProduct("LT Tools para CorelDRAW")]
[assembly: AssemblyCompany("Line Tape Iluminação e Locação")]
[assembly: AssemblyVersion("2.6.0.0")]

static class Instalador
{
    const string T = "LT Tools - Instalador";

    [STAThread]
    static void Main()
    {
        if (MessageBox.Show("Instalar o LT Tools 2.6 no CorelDRAW?\n\nO CorelDRAW precisa estar FECHADO.",
                            T, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;

        while (Process.GetProcessesByName("CorelDRW").Length > 0)
        {
            if (MessageBox.Show("O CorelDRAW ainda está aberto.\nFeche TODOS os CorelDRAW abertos e clique em OK.",
                                T, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) != DialogResult.OK) return;
        }

        try
        {
            string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string lt = Path.Combine(docs, @"Corel\Corel Content\LTTools");
            string ico = Path.Combine(lt, "icones");
            Directory.CreateDirectory(ico);

            Assembly asm = Assembly.GetExecutingAssembly();
            byte[] gms = null;
            foreach (string nome in asm.GetManifestResourceNames())
            {
                byte[] dados;
                using (Stream s = asm.GetManifestResourceStream(nome))
                {
                    dados = new byte[s.Length];
                    int pos = 0;
                    while (pos < dados.Length)
                    {
                        int lidos = s.Read(dados, pos, dados.Length - pos);
                        if (lidos <= 0) break;
                        pos += lidos;
                    }
                }
                string ext = Path.GetExtension(nome).ToLowerInvariant();
                if (ext == ".png" || ext == ".bmp" || ext == ".ico")
                    File.WriteAllBytes(Path.Combine(ico, nome), dados);
                else
                {
                    File.WriteAllBytes(Path.Combine(lt, nome), dados);
                    if (ext == ".gms") gms = dados;
                }
            }

            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(@"Software\VB and VBA Program Settings\LTTools\Config"))
                k.SetValue("Icones", ico + "\\");

            string corel = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Corel");
            int n = 0;
            string lista = "";
            if (Directory.Exists(corel) && gms != null)
            {
                foreach (string d in Directory.GetDirectories(corel))
                {
                    if (!Directory.Exists(Path.Combine(d, "Draw"))) continue;
                    string g = Path.Combine(d, @"Draw\GMS");
                    Directory.CreateDirectory(g);
                    foreach (string f in Directory.GetFiles(g, "*LTPlugin*.gms")) File.Delete(f);
                    File.WriteAllBytes(Path.Combine(g, "LTPlugin.gms"), gms);
                    n++;
                    lista += "\n   - " + Path.GetFileName(d);
                }
            }

            if (n == 0)
                MessageBox.Show("Não encontrei o CorelDRAW neste computador.\nAbra o CorelDRAW uma vez e rode o instalador de novo.",
                                T, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
                MessageBox.Show("LT Tools instalado!" + lista +
                                "\n\nAbra o CorelDRAW: a barra \"LT Tools\" aparece sozinha." +
                                "\nArraste a barra para junto das outras, no topo da tela.\nDica: deixe só um CorelDRAW aberto por vez.",
                                T, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show("Erro na instalação:\n" + ex.Message, T, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
