namespace SAPGui.Core.Interfaces.Components;

/// <summary>
/// Representa a barra de status do SAP (sbar).
/// </summary>
public interface ISAPStatusBar : ISAPComponent
{
    /// <summary>
    /// Obtém o tipo da mensagem atual (S=Success, E=Error, W=Warning, A=Abort, I=Information).
    /// </summary>
    string MessageType { get; }

    // A barra de status pode ter várias mensagens (message1, message2, etc.)
    // Poderíamos adicionar métodos para acessá-las se necessário.
} 