using NUnit.Framework;
using SapGui.Infrastructure.Wrappers; // Corrigido
using System;

namespace SapGui.Infrastructure.Tests.Wrappers;

// Fake COM object específico para TextField (herda do base)
public class FakeSapTextFieldComObject : FakeSapComObject // Usa a fake base
{
    public int CaretPosition { get; set; } = 0;
    public int MaxLength { get; set; } = 255;
}

[TestFixture]
public class SAPTextFieldWrapperTests
{
    private FakeSapTextFieldComObject _fakeComObject = null!;
    private SAPTextFieldWrapper _wrapper = null!;

    [SetUp]
    public void SetUp()
    {
        _fakeComObject = new FakeSapTextFieldComObject();
        // Instancia o wrapper real com o objeto fake
        _wrapper = new SAPTextFieldWrapper(_fakeComObject);
    }

    [Test]
    public void Text_Getter_DelegatesToComObject()
    {
        _fakeComObject.Text = "Initial Text";
        Assert.That(_wrapper.Text, Is.EqualTo(_fakeComObject.Text));
    }

    [Test]
    public void Text_Setter_DelegatesToComObject()
    {
        const string newText = "Set Text";
        _wrapper.Text = newText;
        Assert.That(_fakeComObject.Text, Is.EqualTo(newText));
    }

    [Test]
    public void CaretPosition_Getter_DelegatesToComObject()
    {
        _fakeComObject.CaretPosition = 10;
        Assert.That(_wrapper.CaretPosition, Is.EqualTo(_fakeComObject.CaretPosition));
    }

    [Test]
    public void CaretPosition_Setter_DelegatesToComObject()
    {
        const int newPosition = 5;
        _wrapper.CaretPosition = newPosition;
        Assert.That(_fakeComObject.CaretPosition, Is.EqualTo(newPosition));
    }

     [Test]
    public void MaxLength_Getter_DelegatesToComObject()
    {
        _fakeComObject.MaxLength = 50;
        Assert.That(_wrapper.MaxLength, Is.EqualTo(_fakeComObject.MaxLength));
    }
} 