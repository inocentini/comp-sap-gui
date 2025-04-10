using System.Data;
using SAPGui.Core.Interfaces.Components;
using SAPGui.Core.Models;

namespace SAPGui.Core.Interfaces;

/// <summary>
/// Define o contrato para serviços que interagem com o SAP GUI Scripting API.
/// </summary>
public interface ISAPService : IDisposable
{
    /// <summary>
    /// Abre uma nova conexão com o SAP e retorna a janela principal.
    /// </summary>
    /// <returns>A janela principal (wnd[0]) ou null se não puder conectar.</returns>
    ISAPWindow? OpenSAP();

    /// <summary>
    /// Realiza o login no SAP com as credenciais fornecidas.
    /// </summary>
    /// <param name="credentials">As credenciais de login.</param>
    /// <returns>A janela principal após o login.</returns>
    /// <exception cref="InvalidOperationException">Se a sessão não puder ser estabelecida.</exception>
    /// <exception cref="Exception">Se ocorrer um erro durante o login.</exception>
    ISAPWindow LoginSAP(SAPCredentials credentials);

    /// <summary>
    /// Acessa uma transação SAP específica.
    /// </summary>
    /// <param name="transactionCode">O código da transação.</param>
    void AccessTransaction(string transactionCode);

    /// <summary>
    /// Encontra um componente na tela SAP pelo seu ID.
    /// </summary>
    /// <typeparam name="T">O tipo de ISAPComponent esperado.</typeparam>
    /// <param name="id">O ID do componente.</param>
    /// <returns>O componente SAP encontrado como o tipo T, ou null se não encontrado ou tipo incorreto.</returns>
    T? FindById<T>(string id) where T : class, ISAPComponent;

    /// <summary>
    /// Encontra um componente na tela SAP pelo seu ID, sem checagem de tipo forte no retorno.
    /// </summary>
    /// <param name="id">O ID do componente.</param>
    /// <returns>O componente SAP encontrado, ou null se não encontrado.</returns>
    ISAPComponent? FindById(string id);

    // Removido FindElement(ElementLocator) - Buscar apenas por ID é mais robusto com Scripting API.
    // Se necessário, pode ser reintroduzido com lógica mais complexa.

    // Removido TryFindElementById - FindById<T> e FindById agora retornam null se não encontrado.

    // Removido Click(string elementId) - A ação será feita através do componente (ex: button.Press()).

    /// <summary>
    /// Obtém o componente da barra de status do SAP.
    /// </summary>
    /// <returns>O componente da barra de status ou null se não encontrado.</returns>
    ISAPStatusBar? GetStatusBar();

    // Removido SendKey(int vkeyCode) - A ação será feita através do componente ISAPWindow.

    // Removido GridToDataTable(string gridId) - A ação será feita através do componente ISAPGridView.

    /// <summary>
    /// Fecha a conexão/sessão atual com o SAP.
    /// </summary>
    void CloseSAP();
} 