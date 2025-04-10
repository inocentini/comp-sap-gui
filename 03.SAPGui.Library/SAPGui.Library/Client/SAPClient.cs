using SAPGui.Core.Interfaces;
using SAPGui.Core.Interfaces.Components;
using SAPGui.Core.Models;
using SAPGui.Core.Exceptions;
using System;
using System.Data;
using SAPGui.Library; // Namespace da interface ISapClient

namespace SAPGui.Library.Client;

/// <summary>
/// Cliente de alto nível para interagir com o SAP.
/// Utiliza uma implementação de ISAPService para realizar as operações.
/// </summary>
internal class SAPClient : ISapClient // Implementa a interface, marcada como internal
{
    private readonly ISAPService _sapService;
    private bool _disposed = false;
    private ISAPWindow? _currentWindow; // Mantém referência à janela principal/ativa

    /// <summary>
    /// Cria uma nova instância do SAPClient.
    /// O construtor é internal para forçar o uso da SapClientFactory.
    /// </summary>
    /// <param name="sapService">A implementação do serviço SAP a ser utilizada.</param>
    /// <exception cref="ArgumentNullException">Lançada se sapService for null.</exception>
    internal SAPClient(ISAPService sapService) // Construtor internal
    {
        _sapService = sapService ?? throw new ArgumentNullException(nameof(sapService));
    }

    /// <summary>
    /// Abre a conexão com o SAP e obtém a janela principal.
    /// </summary>
    /// <exception cref="Exception">Se ocorrer um erro ao conectar.</exception>
    public void Open()
    {
        _currentWindow = _sapService.OpenSAP();
        if (_currentWindow == null) {
            // Tenta obter novamente após um tempo se OpenSAP iniciou o processo
            System.Threading.Thread.Sleep(1000);
             _currentWindow = _sapService.OpenSAP();
             if (_currentWindow == null) {
                 throw new Exception("Não foi possível obter a janela principal do SAP após a abertura.");
             }
        }
         Console.WriteLine($"Janela SAP principal obtida: {_currentWindow.Id} - {_currentWindow.Text}");
    }

    /// <summary>
    /// Realiza o login no SAP.
    /// </summary>
    /// <param name="user">Usuário SAP.</param>
    /// <param name="password">Senha SAP.</param>
    /// <param name="client">Mandante SAP.</param>
    /// <param name="language">Idioma SAP (ex: PT, EN).</param>
    public void Login(string user, string password, string client, string language)
    {
        var credentials = new SAPCredentials { User = user, Password = password, Client = client, Language = language };
        // O método LoginSAP agora retorna a janela principal após o login
        _currentWindow = _sapService.LoginSAP(credentials);
        Console.WriteLine($"Login realizado. Janela atual: {_currentWindow.Id} - {_currentWindow.Text}");
    }

    /// <summary>
    /// Acessa uma transação SAP.
    /// </summary>
    /// <param name="transactionCode">Código da transação.</param>
    public void GoToTransaction(string transactionCode) => _sapService.AccessTransaction(transactionCode);

    /// <summary>
    /// Obtém a janela principal atual (normalmente wnd[0]).
    /// </summary>
    /// <returns>A janela principal.</returns>
    /// <exception cref="InvalidOperationException">Se a janela não foi obtida (Open/Login não chamado ou falhou).</exception>
    public ISAPWindow GetMainWindow()
    {
        return _currentWindow ?? throw new InvalidOperationException("A janela principal do SAP não está disponível. Chame Open() ou Login() primeiro.");
    }

    // --- Métodos para encontrar componentes --- 

    /// <summary>
    /// Encontra um componente de qualquer tipo pelo ID.
    /// </summary>
    /// <param name="id">ID do componente.</param>
    /// <returns>O componente encontrado ou null.</returns>
    public ISAPComponent? FindComponentById(string id) => _sapService.FindById(id);

    /// <summary>
    /// Encontra um componente de um tipo específico pelo ID.
    /// </summary>
    /// <typeparam name="T">O tipo de ISAPComponent esperado (ex: ISAPButton, ISAPTextField).</typeparam>
    /// <param name="id">ID do componente.</param>
    /// <returns>O componente encontrado como tipo T, ou null.</returns>
    public T? FindComponentById<T>(string id) where T : class, ISAPComponent => _sapService.FindById<T>(id);

    // --- Métodos auxiliares fortemente tipados --- 

    /// <summary>
    /// Encontra um botão pelo ID.
    /// </summary>
    public ISAPButton? FindButtonById(string id) => FindComponentById<ISAPButton>(id);

    /// <summary>
    /// Encontra um campo de texto pelo ID.
    /// </summary>
    public ISAPTextField? FindTextFieldById(string id) => FindComponentById<ISAPTextField>(id);

    /// <summary>
    /// Encontra um GridView pelo ID.
    /// </summary>
    public ISAPGridView? FindGridViewById(string id) => FindComponentById<ISAPGridView>(id);

    /// <summary>
    /// Obtém a barra de status.
    /// </summary>
    public ISAPStatusBar? GetStatusBar() => _sapService.GetStatusBar();

    /// <summary>
    /// Encontra um CheckBox pelo ID.
    /// </summary>
    public ISAPCheckbox? FindCheckboxById(string id) => FindComponentById<ISAPCheckbox>(id);

    // --- Métodos de Interação --- 

    /// <summary>
    /// Define o texto de um campo de texto encontrado pelo ID.
    /// </summary>
    /// <param name="id">ID do campo de texto.</param>
    /// <param name="text">Texto a ser definido.</param>
    /// <exception cref="ElementNotFoundException">Se o campo não for encontrado.</exception>
    public void SetTextFieldValue(string id, string text)
    {
        var textField = FindTextFieldById(id);
        if (textField != null)
        {
            textField.Text = text;
        }
        else
        {
            throw new ElementNotFoundException($"Campo de texto com ID '{id}' não encontrado.", null); // Usar null para inner exception ou criar uma exceção específica
        }
    }

    /// <summary>
    /// Pressiona um botão encontrado pelo ID.
    /// </summary>
    /// <param name="id">ID do botão.</param>
    /// <exception cref="ElementNotFoundException">Se o botão não for encontrado.</exception>
    public void ClickButton(string id)
    {
        var button = FindButtonById(id);
        if (button != null)
        {
            button.Press();
        }
        else
        {
            throw new ElementNotFoundException($"Botão com ID '{id}' não encontrado.", null);
        }
    }

    /// <summary>
    /// Obtém o texto da barra de status.
    /// </summary>
    /// <returns>O texto da barra de status ou string vazia se não encontrada.</returns>
    public string GetStatusBarText()
    {
        return GetStatusBar()?.Text ?? string.Empty;
    }

    /// <summary>
    /// Envia uma tecla virtual (VKey) para a janela principal atual.
    /// </summary>
    /// <param name="vkeyCode">Código VKey.</param>
    public void SendVKey(int vkeyCode)
    {
        GetMainWindow().SendVKey(vkeyCode); // Usa a janela principal armazenada
    }

    /// <summary>
    /// Extrai dados de um GridView (encontrado pelo ID) para um DataTable.
    /// </summary>
    /// <param name="gridId">ID do GridView.</param>
    /// <returns>DataTable com os dados.</returns>
    /// <exception cref="ElementNotFoundException">Se o GridView não for encontrado.</exception>
    public DataTable GetGridData(string gridId)
    {
        var gridView = FindGridViewById(gridId);
        if (gridView != null)
        {
            return gridView.GetAsDataTable(); // Usa o método do wrapper
        }
        else
        {
            throw new ElementNotFoundException($"GridView com ID '{gridId}' não encontrado.", null);
        }
    }

    /// <summary>
    /// Fecha a conexão/sessão SAP.
    /// </summary>
    public void Close()
    {
        _sapService.CloseSAP();
        _currentWindow = null; // Limpa referência da janela
    }

    /// <summary>
    /// Libera os recursos utilizados pelo cliente e pelo serviço SAP subjacente.
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                // Libera recursos gerenciados (se houver)
            }
            _sapService?.Dispose(); // Chama Dispose do serviço injetado
            _currentWindow = null;
            _disposed = true;
        }
    }

    // O finalizador (~SAPClient) não é estritamente necessário aqui
    // porque a responsabilidade de liberar o objeto COM está no SAPService,
    // e o Dispose do SAPClient garante que o Dispose do SAPService seja chamado.
} 