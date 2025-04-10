namespace SAPGui.Core.Interfaces.Components;

/// <summary>
/// Representa um CheckBox SAP (GuiCheckBox).
/// </summary>
public interface ISAPCheckbox : ISAPComponent
{
    /// <summary>
    /// Obtém ou define o estado selecionado do CheckBox.
    /// </summary>
    bool Selected { get; set; }
} 