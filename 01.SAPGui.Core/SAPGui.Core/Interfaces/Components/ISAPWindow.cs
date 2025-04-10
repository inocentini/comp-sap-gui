namespace SAPGui.Core.Interfaces.Components;

/// <summary>
/// Representa uma janela SAP (wnd).
/// </summary>
public interface ISAPWindow : ISAPComponent
{
    /// <summary>
    /// Envia um código VKey para a janela.
    /// </summary>
    /// <param name="vKey">O código VKey.</param>
    void SendVKey(int vKey);

    /// <summary>
    /// Maximiza a janela.
    /// </summary>
    void Maximize();

    /// <summary>
    /// Fecha a janela.
    /// </summary>
    void Close();

    // Poderia ter métodos para encontrar filhos: FindById<T>(string id) where T : ISAPComponent
} 