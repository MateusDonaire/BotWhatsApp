using System;
using System.Runtime.InteropServices;
using System.Threading;
using WindowsInput;
using WindowsInput.Native;

namespace RoboWhatsAppDesktop
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Iniciando Robô...");
            Program program = new Program();
            InputSimulator inputSimulator = new InputSimulator();
            var watch = new System.Diagnostics.Stopwatch();

            AbrirWhatsAppDesktop();
            try
            {
                watch.Start();

                EnviarImagensNoGrupoDeLeitura(program, inputSimulator);

                if (DateTime.Now.DayOfWeek == DayOfWeek.Sunday) 
                    EnviarImagensNoGrupoDeMaterial();

                watch.Stop();

                EnviarMensagemDoTempoDeExecucao(inputSimulator, watch);

            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                Console.WriteLine("Finalizando processo...");
                Thread.Sleep(2000);
                LeftMouseClick(1897, 13);
            }
        }

        private static void AbrirWhatsAppDesktop()
        {
            Console.WriteLine("Abrindo WhatsApp...");
            System.Diagnostics.Process.Start(@"C:\Users\GAMER\AppData\Local\WhatsApp\WhatsApp.exe");
            Console.WriteLine("Aguardando tempo de sincronização: O processo já irá ser iniciado...");
            Thread.Sleep(60000);
        }

        private static void EnviarMensagemDoTempoDeExecucao(InputSimulator inputSimulator, System.Diagnostics.Stopwatch watch)
        {
            var segundos = watch.ElapsedMilliseconds / 1000;
            var data = DateTime.Now.ToString("D");

            SelecionarPesquisa();

            inputSimulator.Keyboard.TextEntry("Eu");
            Thread.Sleep(1000);
            inputSimulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            Thread.Sleep(1000);

            if(segundos >= 60)
            {
                var minuto = 1;
                segundos = segundos - 60;
                inputSimulator.Keyboard.TextEntry($"*{data}* - *Mateus*, consegui enviar as imagens em *{minuto} minuto e {segundos} segundos*!!! Boa noite =D");
            }
            else
                inputSimulator.Keyboard.TextEntry($"*{data}* - *Mateus*, consegui enviar as imagens em *{segundos} segundos*!!! Boa noite =D");

            Thread.Sleep(500);
            inputSimulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        }

        private static void EnviarImagensNoGrupoDeLeitura(Program program, InputSimulator inputSimulator)
        {
            string planoLeitura = "";
            string sequencial = "LEITURA SEQUENCIAL";
            string robert = "LEITURA ROBERT ROBERTS";
            string diaLeitura, path, fullPath;

            for (int i = 0; i < 2; i++)
            {
                SelecionarPesquisa();
                if (i == 0)
                {
                    planoLeitura = sequencial;
                    SelecionarGrupo(inputSimulator, i, planoLeitura);
                }
                else
                {
                    planoLeitura = robert;
                    SelecionarGrupo(inputSimulator, i, planoLeitura);
                }
                SelecionarImagem();
                FormatarCaminhoDaImagem(program, planoLeitura, out diaLeitura, out path, out fullPath);
                EnviarImagem(inputSimulator, diaLeitura, path, fullPath);
            }
        }

        private static void FormatarCaminhoDaImagem(Program program, string planoLeitura, out string diaLeitura, out string path, out string fullPath)
        {
            diaLeitura = program.FormatarDiaLeitura();
            path = @"C:\Users\GAMER\Pictures\PlanoLeitura\" + planoLeitura + @"\";
            fullPath = path + diaLeitura;
        }

        private static void SelecionarImagem()
        {
            Console.WriteLine("Selecionando imagem...");
            LeftMouseClick(649, 1009);
            Thread.Sleep(2000);
            LeftMouseClick(646, 948);
            Thread.Sleep(2000);
        }

        private static void SelecionarGrupo(InputSimulator inputSimulator, int i, string planoLeitura)
        {
            Console.WriteLine("Iniciado " + (i + 1) + "º envio: Grupo " + planoLeitura);
            inputSimulator.Keyboard.TextEntry(planoLeitura);
            Thread.Sleep(500);
            inputSimulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            Thread.Sleep(500);
        }

        private static void EnviarImagem(InputSimulator inputSimulator, string diaLeitura, string path, string fullPath)
        {
            Console.WriteLine("Inserindo imagem... \n\rCaminho: " + path + "\n\rImagem: " + diaLeitura + ".png");
            Thread.Sleep(1000);
            inputSimulator.Keyboard.TextEntry(fullPath);
            Thread.Sleep(2000);
            inputSimulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);

            Console.WriteLine("Enviando imagem no grupo...");
            Thread.Sleep(500);
            inputSimulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            //inputSimulator.Keyboard.KeyPress(VirtualKeyCode.ESCAPE);
            Console.WriteLine("IMAGEM ENVIADA!");
        }

        private static void SelecionarPesquisa()
        {
            Thread.Sleep(500);
            LeftMouseClick(320, 108);
            Thread.Sleep(500);
        }

        private static void EnviarImagensNoGrupoDeMaterial()
        {
            Program program = new Program();
            InputSimulator inputSimulator = new InputSimulator();

            Console.WriteLine("Enviando imagens nos grupos de material...");
            for (int i = 0; i < 2; i++)
            {
                SelecionarPesquisa();

                string grupoMaterial = "";
                var planoLeitura = "";
                string diaLeitura, materialSemanal, path, fullPath;

                if (i == 0)
                {
                    planoLeitura = "LEITURA SEQUENCIAL";
                    grupoMaterial = "Material Sequencial";
                    SelecionarGrupo(inputSimulator, i, grupoMaterial);
                }
                else
                {
                    planoLeitura = "LEITURA ROBERT ROBERTS";
                    grupoMaterial = "Material Robert Roberts";
                    SelecionarGrupo(inputSimulator, i, grupoMaterial);
                }

                SelecionarImagem();
                FormatarCaminhoDasImagens(program, planoLeitura, out diaLeitura, out materialSemanal, out path, out fullPath);
                Console.WriteLine("Inserindo imagens... \n\rCaminho: " + path + "\n\rImagens: " + materialSemanal);
                EnviarImagem(inputSimulator, diaLeitura, path, fullPath);

                EnviarMensagem(inputSimulator);
            }
        }

        private static void EnviarMensagem(InputSimulator inputSimulator)
        {
            var data = DateTime.Now.ToString("D");
            var msg = $"Olá pessoal, Aqui está o material para caso eu não consiga enviar, boa semana a todos, *estamos juntos nessa =)*! - {data}";
            Thread.Sleep(2000);
            inputSimulator.Keyboard.TextEntry(msg);
            Thread.Sleep(2000);
            inputSimulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            Thread.Sleep(2000);
        }

        private static void FormatarCaminhoDasImagens(Program program, string planoLeitura, out string diaLeitura, out string materialSemanal, out string path, out string fullPath)
        {
            diaLeitura = program.FormatarDiaLeitura ();
            materialSemanal = program.SelecionarMaterialSemanal(diaLeitura);
            path = @"C:\Users\mateu\Pictures\PlanoLeitura\" + planoLeitura + @"\";
            fullPath = path + materialSemanal;
        }

        private string SelecionarMaterialSemanal(string diaLeitura)
        {
            Console.WriteLine("Identificando imagens que serão enviadas...");
            string materialSemanal = "";
            int diaLeituraInt = Convert.ToInt32(diaLeitura);
            diaLeitura = Convert.ToString(diaLeituraInt);
            for (int i = 0; i < 7; i++)
            {
                int img = i + 1;
                var novoDia = (Convert.ToInt32(diaLeitura) + img);
                Console.WriteLine((i + 1) + "º imagem selecionada!");
                if (i == 0)
                {
                    if (diaLeitura.Length == 1) materialSemanal += "\"" + "00" + novoDia + "\" ";
                    else if (diaLeitura.Length == 2) materialSemanal += "\"" + "0" + novoDia + "\" ";
                    else materialSemanal += "\"" + novoDia + "\" ";
                }
                else
                {
                    if (novoDia.ToString().Length == 1) materialSemanal += "\"" + "00" + novoDia + "\" ";
                    else if (novoDia.ToString().Length == 2) materialSemanal += "\"" + "0" + novoDia + "\" ";
                    else materialSemanal += "\"" + novoDia + "\" ";
                }

            }

            Console.WriteLine("Material selecionado com sucesso!");
            return materialSemanal;
        }

        private string FormatarDiaLeitura()
        {
            Console.WriteLine("Identificando o qual dia da leitura...");
            string diaFormatado;
            var dia = (DateTime.Now.DayOfYear - 9).ToString();
            
            if (dia.Length == 1) diaFormatado = "00" + dia;
            else if (dia.Length == 2) diaFormatado = "0" + dia;
            else diaFormatado = dia;

            
            return diaFormatado;

        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool SetCursorPos(int x, int y);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        public const int MOUSEEVENTF_LEFTDOWN = 0x02;
        public const int MOUSEEVENTF_LEFTUP = 0x04;

        public static void LeftMouseClick(int xpos, int ypos)
        {
            SetCursorPos(xpos, ypos);
            mouse_event(MOUSEEVENTF_LEFTDOWN, xpos, ypos, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, xpos, ypos, 0, 0);
        }

    }
}
