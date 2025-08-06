using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using Microsoft.Win32;


/*********************************************************************************************************
- Autor: Elton Lima da Silva
- Última atualização: 11/07/2025
- Versão: 1.0.0
- comando no prompt para build da aplicação: dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true


- Finalidade: Aplicação é destinada a realizar a troca do papel de parede do Windows de forma automática.
A cadas 10 minutos é checada a existencia de um novo arquivo de imagem, se encontrar, ele será 
definido como plano de fundo. Este arquivo fica localizado na pasta WallPaper no Servidor NAS 
no caminho \\10.172.0.11\public\wallpaper. 
A tarefa chamada WallPaperAutoSet deverá será criada de forma automática na máquina local.
C:\Users\{nome_do_usuario}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup

************************************************************************************************************/

class Program
{
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
    private const int SPI_SETDESKWALLPAPER = 20;
    private const int SPIF_UPDATEINIFILE = 0x01;
    private const int SPIF_SENDWININICHANGE = 0x02;

    static void Main()
    {
        Log("🔄 Iniciando rotina de papel de parede.");

        string exePath = Process.GetCurrentProcess().MainModule.FileName;
        string taskName = "TrocarPapelParede";
        string currentUser = WindowsIdentity.GetCurrent().Name;

        if (!TaskExists(taskName))
        {
            // string currentUser = Environment.UserName;
            // string schtasksCmd = $"schtasks /create /tn \"{taskName}\" /tr \"\\\"{exePath}\\\"\" /sc minute /mo 10 /ru \"{currentUser}\" /rl limited /f";            
            
            // bool agendado = RunSilentCommandWithOutput(schtasksCmd, out string retorno);

            // if (agendado && TaskExists(taskName))
            // {
            //     Log("✅ Tarefa agendada e confirmada com sucesso.");
            // }
            // else
            // {
            //     Log("❌ Falha ao agendar tarefa. Saída:");
            //     Log(retorno);
            // }
            CriarTarefaParaTodosUsuarios(exePath, taskName);
        }

        string rede = @"\\10.172.0.11\public\wallpaper";
        string usuario = @"10.172.0.11\papel_parede";
        string senha = "papel@2025";
        string localDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Wallpaper");
        string[] extensoesValidas = { ".jpg", ".jpeg", ".png", ".bmp" };

        try
        {
            RunSilentCommand($"net use {rede} /delete /yes");
            RunSilentCommand($"net use {rede} /user:{usuario} {senha}");

            var arquivos = Directory.GetFiles(rede)
                .Where(f => extensoesValidas.Contains(Path.GetExtension(f).ToLower()))
                .OrderByDescending(f => File.GetLastWriteTime(f))
                .ToList();

            if (!arquivos.Any())
            {
                Log("⚠ Nenhuma imagem válida encontrada no NAS.");
                return;
            }

            string imagemSelecionada = arquivos[0];
            Log($"🖼️ Imagem selecionada: {imagemSelecionada}");

            if (!Directory.Exists(localDir))
                Directory.CreateDirectory(localDir);

            string destino = Path.Combine(localDir, Path.GetFileName(imagemSelecionada));
            File.Copy(imagemSelecionada, destino, true);
            Log($"📁 Imagem copiada para: {destino}");

            ApplyWallpaper(destino);
            ApplyWallpaperToAllUsers(destino);
        }
        catch (Exception ex)
        {
            Log("❌ Erro geral: " + ex.Message);
        }
    }

    static void ApplyWallpaper(string imagem)
    {
        try
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop", true))
            {
                key?.SetValue("Wallpaper", imagem!);
                key?.SetValue("WallpaperStyle", "2");
                key?.SetValue("TileWallpaper", "0");
            }

            bool result = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, imagem, SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE);
            Log(result ? "✅ Papel de parede aplicado (usuário atual)." : "❌ Falha ao aplicar papel de parede.");
        }
        catch (Exception ex)
        {
            Log("❌ Erro ao aplicar papel de parede: " + ex.Message);
        }
    }

    static void ApplyWallpaperToAllUsers(string imagem)
    {
        string usersDir = @"C:\Users";
        foreach (var perfil in Directory.GetDirectories(usersDir))
        {
            string nomePerfil = Path.GetFileName(perfil);
            if (nomePerfil is "Default" or "Public") continue;

            string ntUserPath = Path.Combine(perfil, "NTUSER.DAT");
            if (!File.Exists(ntUserPath)) continue;

            string hiveName = $"TempHive_{nomePerfil}";
            try
            {
                RunSilentCommand($"reg load HKU\\{hiveName} \"{ntUserPath}\"");
                RunSilentCommand($"reg add \"HKU\\{hiveName}\\Control Panel\\Desktop\" /v Wallpaper /t REG_SZ /d \"{imagem}\" /f");
                RunSilentCommand($"reg add \"HKU\\{hiveName}\\Control Panel\\Desktop\" /v WallpaperStyle /t REG_SZ /d 2 /f");
                RunSilentCommand($"reg add \"HKU\\{hiveName}\\Control Panel\\Desktop\" /v TileWallpaper /t REG_SZ /d 0 /f");
                RunSilentCommand($"reg unload HKU\\{hiveName}");
                Log($"✔ Papel de parede aplicado para: {nomePerfil}");
            }
            catch (Exception ex)
            {
                Log($"⚠ Falha em {nomePerfil}: {ex.Message}");
            }
        }
    }

    static bool TaskExists(string taskName)
    {
        try
        {
            Process p = new Process();
            p.StartInfo.FileName = "cmd.exe";
            p.StartInfo.Arguments = $"/c schtasks /query /tn \"{taskName}\"";
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.CreateNoWindow = true;
            p.Start();

            string output = p.StandardOutput.ReadToEnd();
            p.WaitForExit();
            return output.Contains(taskName);
        }
        catch
        {
            return false;
        }
    }

    static bool RunSilentCommandWithOutput(string comando, out string output)
    {
        output = "";
        try
        {
            Process p = new Process();
            p.StartInfo.FileName = "cmd.exe";
            p.StartInfo.Arguments = "/c " + comando;
            p.StartInfo.CreateNoWindow = true;
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardError = true;
            p.Start();

            string stdOut = p.StandardOutput.ReadToEnd();
            string stdErr = p.StandardError.ReadToEnd();
            p.WaitForExit();

            output = stdOut + stdErr;
            return p.ExitCode == 0;
        }
        catch (Exception ex)
        {
            output = "Exception: " + ex.Message;
            return false;
        }
    }

    static void RunSilentCommand(string comando)
    {
        RunSilentCommandWithOutput(comando, out _);
    }

    static void Log(string mensagem)
    {
        try
        {
            string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log.txt");
            string linha = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {mensagem}";
            File.AppendAllText(logPath, linha + Environment.NewLine);
        }
        catch
        {
            // Silenciar falha de log
        }
    }

    static void CriarTarefaParaTodosUsuarios(string exePath, string taskName)
    {
        try
        {
            Process p = new Process();
            p.StartInfo.FileName = "cmd.exe";
            p.StartInfo.Arguments = "/c wmic useraccount where \"localaccount='true' and disabled='false'\" get name";
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.CreateNoWindow = true;
            p.Start();

            string output = p.StandardOutput.ReadToEnd();
            p.WaitForExit();

            var usuarios = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                .Where(u => !u.Contains("Name") && !string.IsNullOrWhiteSpace(u));

            foreach (var usuario in usuarios)
            {
                string schtasksCmd = $"schtasks /create /tn \"{taskName}\" /tr \"\\\"{exePath}\\\"\" /sc minute /mo 10 /ru \"{usuario.Trim()}\" /rl limited /f";
                bool agendado = RunSilentCommandWithOutput(schtasksCmd, out string retorno);

                if (agendado)
                    Log($"📝 Tarefa criada para o usuário: {usuario.Trim()}");
                else
                    Log($"⚠ Falha ao criar tarefa para {usuario.Trim()}: {retorno}");
            }
        }
        catch (Exception ex)
        {
            Log("❌ Erro ao criar tarefa para todos os usuários: " + ex.Message);
        }
    }
}
