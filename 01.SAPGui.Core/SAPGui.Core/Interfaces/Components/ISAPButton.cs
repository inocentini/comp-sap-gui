namespace SAPGui.Core.Interfaces.Components;

/// <summary>
/// Representa um botão SAP (GuiButton).
/// </summary>
public interface ISAPButton : ISAPComponent
{
    /// <summary>
    /// Simula o pressionamento do botão.
    /// </summary>
    void Press();
} 