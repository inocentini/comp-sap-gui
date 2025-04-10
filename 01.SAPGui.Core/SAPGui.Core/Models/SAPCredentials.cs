namespace SAPGui.Core.Models;

/// <summary>
/// Representa as credenciais de login do SAP.
/// </summary>
public class SAPCredentials
{
    public string User { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
    public string Client { get; set; } = string.Empty;
    public string Language { get; set; } = string.Empty; // Ex: "EN" ou "PT"
    // Adicione outras propriedades se necessário (ex: SNC Name)
} 