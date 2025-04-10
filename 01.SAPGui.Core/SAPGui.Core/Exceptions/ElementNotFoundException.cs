using System;

namespace SAPGui.Core.Exceptions;

/// <summary>
/// Exceção lançada quando um componente SAP esperado não é encontrado.
/// </summary>
public class ElementNotFoundException : Exception
{
    /// <summary>
    /// Cria uma nova instância de ElementNotFoundException.
    /// </summary>
    public ElementNotFoundException() { }

    /// <summary>
    /// Cria uma nova instância de ElementNotFoundException com uma mensagem específica.
    /// </summary>
    /// <param name="message">A mensagem que descreve o erro.</param>
    public ElementNotFoundException(string message)
        : base(message) { }

    /// <summary>
    /// Cria uma nova instância de ElementNotFoundException com uma mensagem e uma exceção interna.
    /// </summary>
    /// <param name="message">A mensagem que descreve o erro.</param>
    /// <param name="innerException">A exceção que é a causa da exceção atual.</param>
    public ElementNotFoundException(string message, Exception? innerException)
        : base(message, innerException) { }
} 