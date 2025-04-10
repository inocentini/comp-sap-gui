using SAPGui.Core.Interfaces.Components;
using SAPGui.Core.Models;
using System;
using System.Data;

namespace SAPGui.Library
{
    /// <summary>
    /// Interface pública para interagir com o SAP GUI.
    /// Define as operações de alto nível disponíveis para o consumidor da biblioteca.
    /// </summary>
    public interface ISapClient : IDisposable
    {
        /// <summary>
        /// Abre a conexão com o SAP e obtém a janela principal.
        /// </summary>
        void Open();

        /// <summary>
        /// Realiza o login no SAP.
        /// </summary>
        void Login(string user, string password, string client, string language);

        /// <summary>
        /// Acessa uma transação SAP.
        /// </summary>
        void GoToTransaction(string transactionCode);

        /// <summary>
        /// Obtém a janela principal atual (normalmente wnd[0]).
        /// </summary>
        ISAPWindow GetMainWindow();

        /// <summary>
        /// Encontra um componente de qualquer tipo pelo ID.
        /// </summary>
        ISAPComponent? FindComponentById(string id);

        /// <summary>
        /// Encontra um componente de um tipo específico pelo ID.
        /// </summary>
        T? FindComponentById<T>(string id) where T : class, ISAPComponent;

        /// <summary>
        /// Encontra um botão pelo ID.
        /// </summary>
        ISAPButton? FindButtonById(string id);

        /// <summary>
        /// Encontra um campo de texto pelo ID.
        /// </summary>
        ISAPTextField? FindTextFieldById(string id);

        /// <summary>
        /// Encontra um GridView pelo ID.
        /// </summary>
        ISAPGridView? FindGridViewById(string id);

        /// <summary>
        /// Obtém a barra de status.
        /// </summary>
        ISAPStatusBar? GetStatusBar();

        /// <summary>
        /// Encontra um CheckBox pelo ID.
        /// </summary>
        ISAPCheckbox? FindCheckboxById(string id);

        /// <summary>
        /// Define o texto de um campo de texto encontrado pelo ID.
        /// </summary>
        void SetTextFieldValue(string id, string text);

        /// <summary>
        /// Pressiona um botão encontrado pelo ID.
        /// </summary>
        void ClickButton(string id);

        /// <summary>
        /// Obtém o texto da barra de status.
        /// </summary>
        string GetStatusBarText();

        /// <summary>
        /// Envia uma tecla virtual (VKey) para a janela principal atual.
        /// </summary>
        void SendVKey(int vkeyCode);

        /// <summary>
        /// Extrai dados de um GridView (encontrado pelo ID) para um DataTable.
        /// </summary>
        DataTable GetGridData(string gridId);

        /// <summary>
        /// Fecha a conexão/sessão SAP.
        /// </summary>
        void Close();
    }
} 