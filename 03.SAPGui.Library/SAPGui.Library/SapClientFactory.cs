using Microsoft.Extensions.DependencyInjection;
using SAPGui.Core.Interfaces;
using SAPGui.Library.Client; // Namespace da implementação SAPClient
using System.Runtime.Versioning;
using SapGui.Infrastructure;

namespace SAPGui.Library
{
    /// <summary>
    /// Fábrica para criar instâncias de ISapClient.
    /// Encapsula a configuração da injeção de dependência necessária.
    /// </summary>
    public static class SapClientFactory
    {
        private static ServiceProvider? _serviceProvider;
        private static readonly object _lock = new object();

        /// <summary>
        /// Cria e retorna uma nova instância de ISapClient.
        /// </summary>
        /// <returns>Uma instância de ISapClient pronta para uso.</returns>
        /// <exception cref="PlatformNotSupportedException">Se executado em um sistema operacional não Windows.</exception>
        /// <exception cref="InvalidOperationException">Se ocorrer um erro ao configurar ou resolver os serviços.</exception>
        [SupportedOSPlatform("windows")]
        public static ISapClient CreateClient()
        {
            // Inicialização thread-safe do ServiceProvider (lazy)
            if (_serviceProvider == null)
            {
                lock (_lock)
                {
                    if (_serviceProvider == null)
                    {
                        try
                        {
                            var services = new ServiceCollection();

                            // Chama o método de extensão da camada de Infrastructure para configurar a DI
                            services.AddInfrastructure();

                            // Adiciona o próprio SAPClient (a implementação) ao DI,
                            // pode ser Transient ou Scoped dependendo da necessidade.
                            // Transient garante uma instância nova a cada chamada à fábrica.
                            services.AddTransient<SAPClient>(); // Registra a classe concreta

                            _serviceProvider = services.BuildServiceProvider();
                        }
                        catch (Exception ex)
                        {
                            // Logar erro
                            Console.WriteLine($"Erro ao configurar o ServiceProvider na SapClientFactory: {ex.Message}");
                            throw new InvalidOperationException("Falha ao inicializar a fábrica do SAP Client.", ex);
                        }
                    }
                }
            }

            try
            {
                // Obtém a implementação SAPClient (que agora depende de ISAPService injetado)
                // Precisamos garantir que SAPClient implemente ISapClient e seu construtor seja acessível.
                // Ou podemos resolver ISAPService e instanciar SAPClient manualmente aqui.
                // Abordagem 1: Resolver SAPClient diretamente se ele estiver registrado
                // return _serviceProvider.GetRequiredService<SAPClient>();

                // Abordagem 2: Resolver ISAPService e instanciar SAPClient manualmente
                // Isso é mais flexível se o construtor de SAPClient for internal.
                var sapService = _serviceProvider.GetRequiredService<ISAPService>();
                return new SAPClient(sapService); // Certifique-se que o construtor é acessível (internal ou public)
            }
            catch (Exception ex)
            {
                // Logar erro
                Console.WriteLine($"Erro ao criar instância do SAPClient via fábrica: {ex.Message}");
                throw new InvalidOperationException("Falha ao criar a instância do SAP Client.", ex);
            }
        }

        /// <summary>
        /// Descarta o ServiceProvider interno (se aplicável, para cenários de longa duração ou teste).
        /// </summary>
        public static void DisposeProvider()
        {
            lock (_lock)
            {
                _serviceProvider?.Dispose();
                _serviceProvider = null;
            }
        }
    }
} 