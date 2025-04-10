using Microsoft.Extensions.DependencyInjection;
using SAPGui.Core.Interfaces;
using System.Runtime.Versioning;
using SapGui.Infrastructure.Providers;
using SapGui.Infrastructure.Services;

namespace SapGui.Infrastructure;

/// <summary>
/// Métodos de extensão para configurar a injeção de dependência para a camada de Infrastructure.
/// </summary>
public static class DependencyInjection
{
    /// <summary>
    /// Adiciona os serviços da camada de Infrastructure ao contêiner de DI.
    /// </summary>
    /// <param name="services">A coleção de serviços.</param>
    /// <returns>A coleção de serviços configurada.</returns>
    /// <remarks>
    /// Registra ISAPService com SAPService usando um ciclo de vida Transiente.
    /// Isso garante que cada solicitação de ISAPService obtenha uma nova instância,
    /// o que é geralmente apropriado para gerenciar conexões e estado COM.
    /// </remarks>
    [SupportedOSPlatform("windows")] // Indica que os serviços registrados são específicos do Windows
    public static IServiceCollection AddInfrastructure(this IServiceCollection services)
    {
        // Registrar as novas abstrações e suas implementações
        // Geralmente Singleton ou Scoped para a factory e provider, dependendo do uso.
        // Singleton aqui parece razoável, pois eles não mantêm estado por si só.
        services.AddSingleton<ISapComProvider, SapComProvider>();
        services.AddSingleton<ISapComponentWrapperFactory, SapComponentWrapperFactory>();

        // SAPService agora depende das interfaces acima
        services.AddTransient<ISAPService, SAPService>();

        return services;
    }
} 