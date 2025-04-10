using SAPGui.Core.Interfaces.Components;

namespace SapGui.Infrastructure.Providers
{
    /// <summary>
    /// Interface para abstrair a criação de wrappers de componentes SAP.
    /// </summary>
    public interface ISapComponentWrapperFactory
    {
        /// <summary>
        /// Envolve um objeto COM em um wrapper fortemente tipado (T) ou genérico (ISAPComponent).
        /// </summary>
        /// <typeparam name="T">O tipo de wrapper específico desejado (herda de ISAPComponent).</typeparam>
        /// <param name="comObject">O objeto COM a ser envolvido.</param>
        /// <returns>Uma instância do wrapper T, ou null se o objeto COM for nulo, não for COM, ou o tipo não puder ser mapeado/convertido.</returns>
        T? CreateWrapper<T>(object? comObject) where T : class, ISAPComponent;

        /// <summary>
        /// Envolve um objeto COM em um wrapper genérico ISAPComponent apropriado baseado no seu tipo COM.
        /// </summary>
        /// <param name="comObject">O objeto COM a ser envolvido.</param>
        /// <returns>Uma instância de um wrapper que implementa ISAPComponent, ou null se o objeto COM for nulo, não for COM, ou o tipo não puder ser mapeado.</returns>
        ISAPComponent? CreateWrapper(object? comObject);
    }
} 