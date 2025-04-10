namespace SAPGui.Core.Models;

/// <summary>
/// Define como localizar um elemento na interface SAP.
/// </summary>
public class ElementLocator
{
    /// <summary>
    /// O ID do elemento SAP.
    /// </summary>
    public string? Id { get; set; }

    /// <summary>
    /// O tipo do elemento SAP (ex: "GuiTextField", "GuiButton").
    /// </summary>
    public string? Type { get; set; }

    /// <summary>
    /// O texto associado ao elemento (ex: texto de um botão ou label).
    /// </summary>
    public string? Text { get; set; }

    // Adicione outros critérios de localização se necessário (ex: Nome, Tooltip)

    /// <summary>
    /// Cria um localizador por ID.
    /// </summary>
    public static ElementLocator ById(string id) => new ElementLocator { Id = id };

    /// <summary>
    /// Cria um localizador por Tipo e Texto (exemplo).
    /// </summary>
    public static ElementLocator ByTypeAndText(string type, string text) => new ElementLocator { Type = type, Text = text };

    // Adicione outros métodos fábrica conforme necessário
} 