using System;

namespace SAPGui.Core.Interfaces.Components;

/// <summary>
/// Interface base para todos os componentes da interface SAP GUI.
/// </summary>
public interface ISAPComponent
{
    /// <summary>
    /// O ID único do componente na hierarquia SAP GUI.
    /// </summary>
    string Id { get; }

    /// <summary>
    /// O tipo do componente (ex: "GuiTextField", "GuiButton").
    /// </summary>
    string Type { get; }

    /// <summary>
    /// O nome do componente.
    /// </summary>
    string Name { get; }

    /// <summary>
    /// O texto associado ao componente (pode variar dependendo do tipo).
    /// </summary>
    string Text { get; set; } // Setter pode não ser aplicável a todos, mas é comum

    /// <summary>
    /// Define o foco para este componente.
    /// </summary>
    void SetFocus();

    // Poderia adicionar outros membros comuns como Left, Top, Width, Height, Tooltip, etc.
    // bool Changeable { get; } // Indica se o valor pode ser alterado
    // bool Visible { get; }
} 