using System.Data;

namespace SAPGui.Core.Interfaces.Components;

/// <summary>
/// Representa um GridView SAP (GuiGridView).
/// </summary>
public interface ISAPGridView : ISAPComponent // GridView é um tipo de GuiShell
{
    /// <summary>
    /// Obtém o número de linhas visíveis no grid.
    /// </summary>
    int VisibleRowCount { get; }

    /// <summary>
    /// Obtém o número total de linhas no grid.
    /// </summary>
    int RowCount { get; }

    /// <summary>
    /// Obtém o número de colunas no grid.
    /// </summary>
    int ColumnCount { get; }

    /// <summary>
    /// Obtém o valor de uma célula específica.
    /// </summary>
    /// <param name="row">Índice da linha (baseado em 0).</param>
    /// <param name="columnTechnicalName">Nome técnico da coluna.</param>
    /// <returns>O valor da célula como string.</returns>
    string GetCellValue(int row, string columnTechnicalName);

    /// <summary>
    /// Define o valor de uma célula específica (se editável).
    /// </summary>
    /// <param name="row">Índice da linha (baseado em 0).</param>
    /// <param name="columnTechnicalName">Nome técnico da coluna.</param>
    /// <param name="value">O novo valor.</param>
    void SetCellValue(int row, string columnTechnicalName, string value);

    /// <summary>
    /// Seleciona uma célula específica.
    /// </summary>
    /// <param name="row">Índice da linha (baseado em 0).</param>
    /// <param name="columnTechnicalName">Nome técnico da coluna.</param>
    void SetCurrentCell(int row, string columnTechnicalName);

    /// <summary>
    /// Simula um duplo clique na célula atual.
    /// </summary>
    void DoubleClickCurrentCell();

    /// <summary>
    /// Obtém os nomes técnicos de todas as colunas.
    /// </summary>
    /// <returns>Uma lista com os nomes técnicos das colunas.</returns>
    IReadOnlyList<string> GetColumnTechnicalNames();

    /// <summary>
    /// Obtém os títulos de todas as colunas.
    /// </summary>
    /// <returns>Uma lista com os títulos das colunas.</returns>
    IReadOnlyList<string> GetColumnTitles();

    /// <summary>
    /// Converte os dados do GridView para um DataTable.
    /// (Pode ser mantido aqui ou movido para um método de extensão/helper)
    /// </summary>
    /// <returns>Um DataTable com os dados do grid.</returns>
    DataTable GetAsDataTable();
} 