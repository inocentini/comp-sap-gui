using SAPGui.Core.Interfaces;
using SAPGui.Core.Interfaces.Components;
using SAPGui.Core.Models;
using SapGui.Infrastructure.Providers;
using SapGui.Infrastructure.Wrappers;
using System;
using System.Data;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace SapGui.Infrastructure.Services;

/// <summary>
/// Implementação concreta do ISAPService usando a API de Scripting do SAP GUI.
/// </summary>
[SupportedOSPlatform("windows")]
public class SAPService : ISAPService, IDisposable
{
    private readonly ISapComProvider _comProvider;
    private readonly ISapComponentWrapperFactory _wrapperFactory;

    // Estado interno mínimo
    private object? _sapSessionComObject; // Mantemos a referência ao objeto COM da sessão
    private ISAPWindow? _mainWindow;      // Referência à janela principal já "wrappada"
    private bool _disposed = false;

    // Injeção de Dependência via Construtor
    public SAPService(ISapComProvider comProvider, ISapComponentWrapperFactory wrapperFactory)
    {
        _comProvider = comProvider ?? throw new ArgumentNullException(nameof(comProvider));
        _wrapperFactory = wrapperFactory ?? throw new ArgumentNullException(nameof(wrapperFactory));
    }

    public ISAPWindow? OpenSAP()
    {
        EnsureNotDisposed();
        // Se já temos uma sessão, talvez reutilizar? Ou sempre buscar a primeira?
        // Por enquanto, segue a lógica original de buscar a primeira ativa.
        try
        {
            object? guiApp = _comProvider.GetApplication();
            if (guiApp == null) throw new InvalidOperationException("Não foi possível obter GuiApplication.");

            object? engine = _comProvider.GetScriptingEngine(guiApp);
            if (engine == null) throw new InvalidOperationException("Não foi possível obter ScriptingEngine.");

            if (_comProvider.GetConnectionCount(engine) > 0)
            {
                object? connection = _comProvider.GetConnection(engine, 0);
                if (connection != null && _comProvider.GetSessionCount(connection) > 0)
                {
                    _sapSessionComObject = _comProvider.GetSession(connection, 0);
                    if (_sapSessionComObject != null)
                    {
                        Console.WriteLine("Sessão SAP ativa encontrada e armazenada.");
                        object? mainWndComObject = _comProvider.FindComComponentById(_sapSessionComObject, "wnd[0]");
                        _mainWindow = _wrapperFactory.CreateWrapper<ISAPWindow>(mainWndComObject);
                        return _mainWindow;
                    }
                }
            }
            Console.WriteLine("Nenhuma conexão/sessão SAP ativa encontrada inicialmente.");
            // Limpar estado se não encontrou?
            _sapSessionComObject = null;
            _mainWindow = null;
            return null;
        }
        catch (Exception ex) // Captura exceções do provider ou InvalidOperationExceptions
        {   // Logar o erro
            Console.WriteLine($"Erro ao abrir/conectar ao SAP: {ex.Message}");
            // Relançar ou retornar null/lançar exceção específica do serviço?
            // Por simplicidade, relançamos encapsulada
            throw new Exception($"Erro ao abrir/conectar ao SAP: {ex.Message}", ex);
        }
        // Nota: Não liberamos objetos COM aqui, pois podem ser necessários para login.
        // A liberação ocorrerá no Dispose ou CloseSAP.
    }

    public ISAPWindow LoginSAP(SAPCredentials credentials)
    {
        EnsureNotDisposed();
        // 1. Garantir que temos uma sessão COM
        if (_sapSessionComObject == null)
        {
            // Tenta abrir/reconectar se não tiver uma sessão
            OpenSAP();
            if (_sapSessionComObject == null)
                throw new InvalidOperationException("Não foi possível estabelecer uma sessão SAP para login após tentativa de OpenSAP.");
        }

        try
        {
            // 2. Obter a Janela de Login (geralmente a janela ativa da sessão)
            object? loginWindowComObject = _comProvider.GetActiveWindow(_sapSessionComObject);
            ISAPWindow? loginWindow = _wrapperFactory.CreateWrapper<ISAPWindow>(loginWindowComObject);
            if (loginWindow == null) throw new InvalidOperationException("Não foi possível obter a janela de login SAP.");

            Console.WriteLine($"Tentando login na janela: {loginWindow.Text} (ID: {loginWindow.Id})");

            // 3. Encontrar e preencher campos usando FindInWindow (que agora usa o provider/factory)
            var userField = FindInWindow<ISAPTextField>(loginWindow, "usr/txtRSYST-BNAME");
            var passField = FindInWindow<ISAPTextField>(loginWindow, "usr/pwdRSYST-BCODE");
            var clientField = FindInWindow<ISAPTextField>(loginWindow, "usr/txtRSYST-MANDT");

            SetTextFieldText(userField, credentials.User);
            SetTextFieldText(passField, credentials.Password);
            SetTextFieldText(clientField, credentials.Client);

            try
            {
                var langField = FindInWindow<ISAPTextField>(loginWindow, "usr/txtRSYST-LANGU");
                SetTextFieldText(langField, credentials.Language);
            } catch { Console.WriteLine("Campo de idioma não encontrado ou erro ao definir."); }

            // 4. Enviar Enter
            loginWindow.SendVKey(0);
            Console.WriteLine("Dados de login enviados.");
            System.Threading.Thread.Sleep(1500); // Manter uma pausa pode ser necessário

            // 5. Verificar o resultado (obter a nova janela principal)
            // A sessão COM deve ter sido atualizada. Reobtemos a janela principal.
            // Precisamos reconfirmar a sessão ativa?
            object? mainWndComObject = _comProvider.FindComComponentById(_sapSessionComObject, "wnd[0]");
            _mainWindow = _wrapperFactory.CreateWrapper<ISAPWindow>(mainWndComObject);

            if (_mainWindow == null)
            {
                // Tentar obter status bar da janela de login antes de falhar?
                string errorStatus = GetStatusTextFromWindow(loginWindow) ?? "N/A";
                throw new Exception($"Falha no login. Janela principal não encontrada após envio. Status da janela anterior: {errorStatus}");
            }

            Console.WriteLine("Login possivelmente bem-sucedido. Janela principal obtida.");

            // 6. Verificar Status Bar da nova janela principal
            var statusBar = GetStatusBar(); // Usa _mainWindow internamente
            Console.WriteLine($"Status pós-login: {statusBar?.Text} (Tipo: {statusBar?.MessageType})");
            if (statusBar?.MessageType == "E" || statusBar?.MessageType == "A")
            {
                throw new Exception($"Erro SAP detectado na Status Bar após login: {statusBar.Text}");
            }

            return _mainWindow;
        }
        catch (Exception ex)
        {   // Logar o erro
            Console.WriteLine($"Erro durante o login SAP: {ex.Message}");
            // Limpar estado em caso de erro?
            // CloseSAP(); // Talvez não seja ideal limpar tudo automaticamente.
            throw new Exception($"Erro durante o login SAP: {ex.Message}", ex);
        }
    }

    public void AccessTransaction(string transactionCode)
    {
        EnsureNotDisposed();
        if (_mainWindow == null)
            throw new InvalidOperationException("Janela principal do SAP não está disponível. Faça o login primeiro ou chame OpenSAP.");

        try
        {
            var transactionField = FindInWindow<ISAPTextField>(_mainWindow, "okcd"); // Campo de comando OK Code
            if (transactionField != null)
            {
                transactionField.Text = "/n" + transactionCode;
                _mainWindow.SendVKey(0); // Enter
                Console.WriteLine($"Acessando transação: {transactionCode}");
                System.Threading.Thread.Sleep(500); // Pausa
                // Atualizar _mainWindow? A janela pode mudar após acessar transação.
                // Considerar re-buscar wnd[0] aqui.
            }
            else
            {
                throw new InvalidOperationException("Campo de transação (okcd) não encontrado na janela principal.");
            }
        }
        catch (Exception ex)
        {   Console.WriteLine($"Erro ao acessar transação '{transactionCode}': {ex.Message}");
            throw new Exception($"Erro ao acessar transação '{transactionCode}': {ex.Message}", ex);
        }
    }

    public T? FindById<T>(string id) where T : class, ISAPComponent
    {
        EnsureNotDisposed();
        if (_sapSessionComObject == null) throw new InvalidOperationException("Sessão SAP não iniciada para FindById.");

        try
        {
            object? comObject = _comProvider.FindComComponentById(_sapSessionComObject, id);
            return _wrapperFactory.CreateWrapper<T>(comObject);
        }
        catch (Exception ex)
        {   Console.WriteLine($"Erro em FindById para ID '{id}': {ex.Message}");
            // Lançar exceção ou retornar null?
            // O comportamento original parecia retornar null em alguns casos COM, mas lançava outros.
            // Retornar null parece mais seguro se o componente pode legitimamente não existir.
            return null;
            // throw new Exception($"Erro ao encontrar componente ID '{id}': {ex.Message}", ex); // Alternativa
        }
    }

    public ISAPComponent? FindById(string id)
    {   // Implementação não-genérica
        EnsureNotDisposed();
        if (_sapSessionComObject == null) throw new InvalidOperationException("Sessão SAP não iniciada para FindById.");
        try
        {
            object? comObject = _comProvider.FindComComponentById(_sapSessionComObject, id);
            return _wrapperFactory.CreateWrapper(comObject); // Usa a fábrica para criar wrapper genérico
        }
        catch (Exception ex)
        {   Console.WriteLine($"Erro em FindById (não genérico) para ID '{id}': {ex.Message}");
            return null;
        }
    }

    public ISAPStatusBar? GetStatusBar()
    {
        EnsureNotDisposed();
        // Tenta obter da mainWindow atual, se existir
        if (_mainWindow != null)
        {
            var statusBar = FindInWindow<ISAPStatusBar>(_mainWindow, "sbar");
            if (statusBar != null) return statusBar;
        }

        // Se não conseguiu via _mainWindow, tenta buscar wnd[0] novamente
        Console.WriteLine("Tentando obter status bar buscando wnd[0] novamente...");
        var mainWindow = FindById<ISAPWindow>("wnd[0]");
        if (mainWindow != null)
        {
            _mainWindow = mainWindow; // Atualiza a referência
            return FindInWindow<ISAPStatusBar>(_mainWindow, "sbar");
        }

        Console.WriteLine("Aviso: Não foi possível obter a janela principal ou a status bar.");
        return null;
    }

    // Método auxiliar para encontrar dentro de uma janela (usa FindById)
    private T? FindInWindow<T>(ISAPWindow window, string relativeId) where T : class, ISAPComponent
    {
        if (window == null) return null;
        // Lógica de construção do ID completo (pode ser simplificada ou melhorada)
        string fullId;
        if (relativeId.StartsWith("usr/") || relativeId.StartsWith("ssub/") || relativeId.StartsWith("tabs/") || relativeId.StartsWith("sbar") || relativeId.StartsWith("titl") || relativeId.StartsWith("okcd"))
        {
            fullId = $"{window.Id}/{relativeId}";
        }
        else if (relativeId.StartsWith("wnd["))
        {
            Console.WriteLine($"Aviso: Usando ID '{relativeId}' que parece absoluto dentro de FindInWindow.");
            fullId = relativeId;
        }
        else
        {
            fullId = $"{window.Id}/{relativeId}";
        }
        return FindById<T>(fullId);
    }

    // Método auxiliar para tentar obter o texto da status bar de uma janela específica
    private string? GetStatusTextFromWindow(ISAPWindow window)
    {
        try
        {
            return FindInWindow<ISAPStatusBar>(window, "sbar")?.Text;
        }
        catch { return null; }
    }

     // Método auxiliar para definir texto com segurança
    private void SetTextFieldText(ISAPTextField? field, string? text)
    {
        if (field != null)
        {
            field.Text = text ?? string.Empty;
        } else {
            // Logar ou lançar exceção?
            Console.WriteLine("Aviso: Tentativa de definir texto em campo nulo.");
            // Não lançar exceção aqui pode mascarar erros de ID no FindInWindow
            throw new InvalidOperationException("Campo de texto não encontrado para definir valor.");
        }
    }

    public void CloseSAP()
    {   // O ComProvider agora gerencia a liberação, mas precisamos limpar nosso estado.
        EnsureNotDisposed();
        Console.WriteLine("Limpando referências internas do SAPService...");
        _mainWindow = null;
        // Liberar o objeto COM da sessão que estávamos mantendo
        if (_sapSessionComObject != null)
        {
             _comProvider.ReleaseComObject(_sapSessionComObject);
            _sapSessionComObject = null;
        }
         // Deveríamos liberar GuiApplication e Connection também? Ou o provider faz isso?
         // Depende da implementação do provider. Assumindo que ele não mantém estado, nós não precisamos.
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposed) return;

        if (disposing)
        {
            // Liberar recursos gerenciados aqui (se houver)
        }

        // Liberar recursos não gerenciados (nossa referência ao COM da sessão)
        CloseSAP();

        _disposed = true;
    }

    private void EnsureNotDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(SAPService));
    }

     ~SAPService()
     {
         Dispose(false);
     }
}

// Adicionar exceção personalizada se ainda não existir
// public class ElementNotFoundException : Exception { ... } 