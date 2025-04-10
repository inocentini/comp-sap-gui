using SAPGui.Core.Interfaces.Components;

namespace SapGui.Infrastructure.Wrappers;

internal class SAPButtonWrapper : SAPComponentBase, ISAPButton
{
    public SAPButtonWrapper(object sapComObject) : base(sapComObject) { }

    // Text (rótulo do botão) é herdado.

    public void Press() => SapComObject.Press();
} 