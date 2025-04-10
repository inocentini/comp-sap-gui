using SAPGui.Core.Interfaces.Components;

namespace SapGui.Infrastructure.Wrappers;

internal class SAPTextFieldWrapper : SAPComponentBase, ISAPTextField
{
    public SAPTextFieldWrapper(object sapComObject) : base(sapComObject) { }

    // Propriedade Text é herdada de SAPComponentBase e funciona para TextField

    public int CaretPosition
    {
        get => SapComObject.CaretPosition;
        set => SapComObject.CaretPosition = value;
    }

    public int MaxLength => SapComObject.MaxLength;
} 