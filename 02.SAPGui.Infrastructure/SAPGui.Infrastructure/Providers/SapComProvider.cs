using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace SapGui.Infrastructure.Providers
{
    /// <summary>
    /// Implementação de ISapComProvider usando P/Invoke e COM Interop.
    /// </summary>
    [SupportedOSPlatform("windows")]
    internal class SapComProvider : ISapComProvider
    {
        // P/Invoke para GetActiveObject
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject([MarshalAs(UnmanagedType.LPStruct)] Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        private static readonly Guid CLSID_SAP_GUI_APPLICATION = new Guid("B90F32AD-859E-4173-A04E-3F844F358CA2");
        private const string SapLogonProcessName = "saplogon.exe";
        private const int SapInitializationWaitMs = 7000; // Tempo de espera para inicialização

        private object? TryGetActiveObject(Guid clsid)
        {
            try
            {
                GetActiveObject(clsid, IntPtr.Zero, out object obj);
                return obj;
            }
            catch (COMException ex) when (ex.ErrorCode == -2147221021) // MK_E_UNAVAILABLE
            {
                return null;
            }
            // Outras exceções COM não são tratadas aqui, serão relançadas.
        }

        public object? GetApplication()
        {
            object? sapGuiApp = TryGetActiveObject(CLSID_SAP_GUI_APPLICATION);

            if (sapGuiApp == null)
            {
                Console.WriteLine("SAP GUI não encontrado via GetActiveObject. Tentando iniciar...");
                try
                {
                    using (Process? process = Process.Start(SapLogonProcessName))
                    {
                        if (process != null)
                        {
                            Console.WriteLine($"Processo {SapLogonProcessName} iniciado (PID: {process.Id}). Aguardando {SapInitializationWaitMs}ms...");
                            System.Threading.Thread.Sleep(SapInitializationWaitMs);
                        }
                        else
                        {
                             Console.WriteLine($"Falha ao iniciar o processo {SapLogonProcessName}.");
                        }
                    }
                }
                catch (Exception startEx)
                {
                     // Relança como exceção específica ou loga
                    throw new InvalidOperationException($"Falha ao iniciar {SapLogonProcessName}. Verifique se está no PATH ou forneça o caminho completo. Erro: {startEx.Message}", startEx);
                }
                sapGuiApp = TryGetActiveObject(CLSID_SAP_GUI_APPLICATION); // Tenta novamente
            }

            if (sapGuiApp == null)
            {
                // Poderia lançar uma exceção mais específica
                Console.WriteLine("Não foi possível obter ou iniciar o SAP GUI Scripting Engine após tentativas.");
            }
            return sapGuiApp;
        }

        public object? GetScriptingEngine(object guiApplication)
        {
            try
            {
                // Usar dynamic aqui ainda é a forma mais prática de interagir com COM desconhecido
                dynamic guiApp = guiApplication;
                return guiApp.GetScriptingEngine;
            }
            catch (Exception ex) // Captura RuntimeBinderException ou COMException
            {
                Console.WriteLine($"Erro ao obter ScriptingEngine: {ex.Message}");
                return null;
            }
        }

        public int GetConnectionCount(object scriptingEngine)
        {
            try { return ((dynamic)scriptingEngine).Connections.Count; } catch { return 0; }
        }

        public object? GetConnection(object scriptingEngine, int index)
        {
            try { return ((dynamic)scriptingEngine).Connections(index); } catch { return null; }
        }

        public int GetSessionCount(object connection)
        {
            try { return ((dynamic)connection).Sessions.Count; } catch { return 0; }
        }

        public object? GetSession(object connection, int index)
        {
            try { return ((dynamic)connection).Sessions(index); } catch { return null; }
        }

        public object? FindComComponentById(object session, string id)
        {
            try
            {
                dynamic dynSession = session;
                return dynSession.FindById(id);
            }
            catch (COMException ex)
            {
                bool isNotFound = ex.HResult == -2147352573 || // DISP_E_MEMBERNOTFOUND
                               ex.Message.Contains("could not be found") ||
                               ex.Message.Contains("element not found");
                if (isNotFound)
                {
                    Console.WriteLine($"Componente COM com ID '{id}' não encontrado via provider.");
                    return null; // Retorna null se não encontrado
                }
                // Relança outras exceções COM
                Console.WriteLine($"Erro COM não tratado em FindComComponentById('{id}'): {ex.Message}");
                throw;
            }
             catch (Exception ex) // Captura RuntimeBinderException, etc.
             {
                  Console.WriteLine($"Erro não COM em FindComComponentById('{id}'): {ex.Message}");
                 throw;
             }
        }

        public object? GetActiveWindow(object sessionOrConnection)
        {
            try { return ((dynamic)sessionOrConnection).ActiveWindow; } catch { return null; }
        }

        public void ReleaseComObject(object? obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                try
                {   // Abordagem simples: decrementa a contagem de referência uma vez.
                    Marshal.ReleaseComObject(obj);
                }
                catch (Exception ex)
                { Console.WriteLine($"Erro ao liberar objeto COM: {ex.Message}"); }
            }
        }
    }
} 