using NUnit.Framework;
using SAPGui.Core.Exceptions;
using System;

namespace SAPGui.Core.Tests.Exceptions;

[TestFixture]
public class ElementNotFoundExceptionTests
{
    [Test]
    public void Constructor_Default_ShouldCreateInstance()
    {
        // Arrange & Act
        var exception = new ElementNotFoundException();

        // Assert
        Assert.That(exception, Is.Not.Null);
        Assert.That(exception.InnerException, Is.Null);
        // O Assert da mensagem pode variar dependendo da cultura do sistema,
        // então verificamos apenas se não é nula ou vazia.
        Assert.That(exception.Message, Is.Not.Null.Or.Empty);
    }

    [Test]
    public void Constructor_WithMessage_ShouldSetMessage()
    {
        // Arrange
        const string expectedMessage = "Elemento de teste não encontrado.";

        // Act
        var exception = new ElementNotFoundException(expectedMessage);

        // Assert
        Assert.That(exception, Is.Not.Null);
        Assert.That(exception.Message, Is.EqualTo(expectedMessage));
        Assert.That(exception.InnerException, Is.Null);
    }

    [Test]
    public void Constructor_WithMessageAndInnerException_ShouldSetBoth()
    {
        // Arrange
        const string expectedMessage = "Erro ao localizar elemento.";
        var innerException = new InvalidOperationException("Inner error");

        // Act
        var exception = new ElementNotFoundException(expectedMessage, innerException);

        // Assert
        Assert.That(exception, Is.Not.Null);
        Assert.That(exception.Message, Is.EqualTo(expectedMessage));
        Assert.That(exception.InnerException, Is.SameAs(innerException));
    }
} 