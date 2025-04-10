using SAPGui.Core.Interfaces.Components;

namespace SapGui.Infrastructure.Wrappers;

internal class SAPWindowWrapper : SAPComponentBase, ISAPWindow
{
    public SAPWindowWrapper(object sapComObject) : base(sapComObject) { }

    // Propriedade Text geralmente é o título da janela
    public override string Text { get => SapComObject.Text; set => SapComObject.Text = value; } // Janelas geralmente têm Text

    public void SendVKey(int vKey) => SapComObject.SendVKey(vKey);

    public void Maximize() => SapComObject.Maximize();

    public void Close() => SapComObject.Close();

    // Implementação futura: FindById<T> dentro da janela
} 