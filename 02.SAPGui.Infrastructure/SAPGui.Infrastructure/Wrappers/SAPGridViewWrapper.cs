using SAPGui.Core.Interfaces.Components;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;

namespace SapGui.Infrastructure.Wrappers;

internal class SAPGridViewWrapper : SAPComponentBase, ISAPGridView
{
    public SAPGridViewWrapper(object sapComObject) : base(sapComObject)
    {
        // Validação adicional para garantir que é um GridView
        if (base.Type != "GuiGridView")
        {
            throw new ArgumentException($"O objeto COM fornecido não é do tipo GuiGridView (Tipo: {base.Type}).", nameof(sapComObject));
        }
    }

    // GridView geralmente não tem uma propriedade 'Text' significativa.
    public override string Text { get => string.Empty; set { /* Ignorar */ } }

    public int VisibleRowCount => SapComObject.VisibleRowCount;
    public int RowCount => SapComObject.RowCount;
    public int ColumnCount => SapComObject.ColumnCount;

    public string GetCellValue(int row, string columnTechnicalName)
    {
        try
        {
            return SapComObject.GetCellValue(row, columnTechnicalName);
        }
        catch (COMException ex)
        {
            // Tratar erro específico (ex: célula inválida) ou relançar com mais contexto
            throw new Exception($"Erro ao obter valor da célula ({row}, {columnTechnicalName}): {ex.Message}", ex);
        }
    }

    public void SetCellValue(int row, string columnTechnicalName, string value)
    {
        try
        {
            SapComObject.SetCellValue(row, columnTechnicalName, value);
        }
        catch (COMException ex)
        {
            // Tratar erro específico ou relançar
            throw new Exception($"Erro ao definir valor da célula ({row}, {columnTechnicalName}): {ex.Message}", ex);
        }
    }

    public void SetCurrentCell(int row, string columnTechnicalName)
    {
        try
        {
            SapComObject.SetCurrentCell(row, columnTechnicalName);
        }
        catch (COMException ex)
        {
            throw new Exception($"Erro ao definir célula atual ({row}, {columnTechnicalName}): {ex.Message}", ex);
        }
    }

    public void DoubleClickCurrentCell()
    {
        try
        {
            SapComObject.DoubleClickCurrentCell();
        }
        catch (COMException ex)
        {
            throw new Exception($"Erro ao executar duplo clique na célula atual: {ex.Message}", ex);
        }
    }

    public IReadOnlyList<string> GetColumnTechnicalNames()
    {
        List<string> names = new List<string>();
        dynamic columnCollection = SapComObject.Columns;
        for (int i = 0; i < columnCollection.Count; i++)
        {
            names.Add(columnCollection.ElementAt(i).Name);
        }
        return names.AsReadOnly();
    }

    public IReadOnlyList<string> GetColumnTitles()
    {
        List<string> titles = new List<string>();
        dynamic columnCollection = SapComObject.Columns;
        for (int i = 0; i < columnCollection.Count; i++)
        {
            titles.Add(columnCollection.ElementAt(i).Title);
        }
        return titles.AsReadOnly();
    }

    public DataTable GetAsDataTable()
    {
        DataTable dataTable = new DataTable();
        var technicalNames = GetColumnTechnicalNames();
        var titles = GetColumnTitles();

        // Adiciona colunas ao DataTable usando títulos (e tratando duplicados)
        for (int j = 0; j < ColumnCount; j++)
        {
            string title = !string.IsNullOrEmpty(titles[j]) ? titles[j] : technicalNames[j];
            string dtColName = title;
            int suffix = 1;
            while (dataTable.Columns.Contains(dtColName))
            {
                dtColName = $"{title}_{suffix++}";
            }
            dataTable.Columns.Add(dtColName);
        }

        // Adiciona linhas
        int totalRows = RowCount;
        for (int i = 0; i < totalRows; i++)
        {
            DataRow dataRow = dataTable.NewRow();
            for (int j = 0; j < dataTable.Columns.Count; j++)
            {
                try
                {
                    dataRow[j] = GetCellValue(i, technicalNames[j]);
                }
                catch (Exception ex) // Captura erro de GetCellValue
                {
                    Console.WriteLine($"Aviso: Erro ao ler célula ({i}, {technicalNames[j]}): {ex.Message}");
                    dataRow[j] = DBNull.Value; // Ou string vazia
                }
            }
            dataTable.Rows.Add(dataRow);
        }
        return dataTable;
    }
} 