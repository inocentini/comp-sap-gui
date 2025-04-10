using System;

namespace SapGui.Infrastructure.Providers
{
    /// <summary>
    /// Interface para abstrair a interação direta com objetos COM do SAP GUI Scripting.
    /// </summary>
    public interface ISapComProvider
    {
        /// <summary>Garante que o processo SAP Logon está rodando e retorna o objeto GuiApplication.</summary>
        /// <returns>O objeto COM GuiApplication ou null se não puder ser obtido.</returns>
        object? GetApplication();

        /// <summary>Obtém o objeto ScriptingEngine a partir do objeto GuiApplication.</summary>
        /// <param name="guiApplication">O objeto COM GuiApplication.</param>
        /// <returns>O objeto COM ScriptingEngine ou null.</returns>
        object? GetScriptingEngine(object guiApplication);

        /// <summary>Obtém o número de conexões ativas no ScriptingEngine.</summary>
        /// <param name="scriptingEngine">O objeto COM ScriptingEngine.</param>
        /// <returns>O número de conexões.</returns>
        int GetConnectionCount(object scriptingEngine);

        /// <summary>Obtém uma conexão específica por índice.</summary>
        /// <param name="scriptingEngine">O objeto COM ScriptingEngine.</param>
        /// <param name="index">O índice da conexão.</param>
        /// <returns>O objeto COM GuiConnection ou null.</returns>
        object? GetConnection(object scriptingEngine, int index);

        /// <summary>Obtém o número de sessões dentro de uma conexão.</summary>
        /// <param name="connection">O objeto COM GuiConnection.</param>
        /// <returns>O número de sessões.</returns>
        int GetSessionCount(object connection);

        /// <summary>Obtém uma sessão específica por índice.</summary>
        /// <param name="connection">O objeto COM GuiConnection.</param>
        /// <param name="index">O índice da sessão.</param>
        /// <returns>O objeto COM GuiSession ou null.</returns>
        object? GetSession(object connection, int index);

        /// <summary>Encontra um objeto COM componente pelo seu ID absoluto dentro de uma sessão.</summary>
        /// <param name="session">O objeto COM GuiSession.</param>
        /// <param name="id">O ID completo do componente.</param>
        /// <returns>O objeto COM do componente encontrado ou null.</returns>
        object? FindComComponentById(object session, string id);

        /// <summary>Obtém a janela ativa de uma sessão ou conexão.</summary>
        /// <param name="sessionOrConnection">O objeto COM GuiSession ou GuiConnection.</param>
        /// <returns>O objeto COM da janela ativa ou null.</returns>
        object? GetActiveWindow(object sessionOrConnection);

        /// <summary>Libera um objeto COM de forma segura.</summary>
        /// <param name="obj">O objeto COM a ser liberado.</param>
        void ReleaseComObject(object? obj);
    }
} 