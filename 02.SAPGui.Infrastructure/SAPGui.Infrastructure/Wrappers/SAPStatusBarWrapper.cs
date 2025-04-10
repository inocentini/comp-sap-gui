using SAPGui.Core.Interfaces.Components;

namespace SapGui.Infrastructure.Wrappers;

internal class SAPStatusBarWrapper : SAPComponentBase, ISAPStatusBar
{
    public SAPStatusBarWrapper(object sapComObject) : base(sapComObject) { }

    // StatusBar tem 'Text' mas também 'MessageType'
    public override string Text { get => SapComObject.Text; set { /* Não faz sentido definir texto da status bar */ } }

    public string MessageType => SapComObject.MessageType;
} 