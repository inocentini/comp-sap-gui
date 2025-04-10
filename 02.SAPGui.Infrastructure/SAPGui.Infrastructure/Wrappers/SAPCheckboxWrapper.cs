using SAPGui.Core.Interfaces.Components;
using System;

namespace SapGui.Infrastructure.Wrappers;

internal class SAPCheckboxWrapper : SAPComponentBase, ISAPCheckbox
{
    public SAPCheckboxWrapper(object sapComObject) : base(sapComObject)
    {
        if (base.Type != "GuiCheckBox")
        {
            throw new ArgumentException($"O objeto COM fornecido não é do tipo GuiCheckBox (Tipo: {base.Type}).", nameof(sapComObject));
        }
    }

    public bool Selected
    {
        get => SapComObject.Selected;
        set => SapComObject.Selected = value;
    }

    // Text (label do checkbox) é herdado.
}
 