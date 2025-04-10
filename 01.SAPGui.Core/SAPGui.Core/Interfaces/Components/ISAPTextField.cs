namespace SAPGui.Core.Interfaces.Components;

/// <summary>
/// Representa um campo de texto SAP (GuiTextField, GuiCTextField, GuiPasswordField).
/// </summary>
public interface ISAPTextField : ISAPComponent
{
    // A propriedade Text já está na interface base ISAPComponent
    // string Text { get; set; }

    /// <summary>
    /// Obtém ou define a posição do cursor (caret) no campo.
    /// </summary>
    int CaretPosition { get; set; }

    /// <summary>
    /// Obtém o comprimento máximo do texto permitido no campo.
    /// </summary>
    int MaxLength { get; }
} 