# Biblioteca de Interação com SAP GUI para .NET

Esta biblioteca .NET fornece uma maneira estruturada e testável de interagir com o SAP GUI for Windows usando a API de Scripting COM, seguindo princípios de Clean Architecture.

**Plataforma:** .NET 8 (Windows)

## Visão Geral

A solução é dividida em camadas:

*   **`01.SAPGui.Core`**: Define as interfaces, modelos de domínio e exceções, sem dependências externas (exceto .NET base).
*   **`02.SAPGui.Infrastructure`**: Contém as implementações concretas das interfaces do Core, lidando com a interação COM com o SAP GUI Scripting API, configuração de injeção de dependência e wrappers para os objetos COM.
*   **`03.SAPGui.Library`**: Expõe a API pública e amigável para os consumidores da biblioteca, encapsulando a complexidade interna. Fornece a fábrica `SapClientFactory` e a interface `ISapClient`.

## Pré-requisitos

Antes de usar esta biblioteca, garanta que os seguintes pré-requisitos sejam atendidos no ambiente onde a aplicação será executada:

1.  **Sistema Operacional:** Windows (devido à dependência do COM Interop e SAP GUI for Windows).
2.  **.NET Runtime:** .NET 8 ou superior.
3.  **SAP GUI for Windows:** Instalado e funcional.
4.  **SAP GUI Scripting Habilitado:**
    *   **No Cliente:** Verifique as opções do SAP Logon (`Alt+F12` -> Options -> Accessibility & Scripting -> Scripting) e certifique-se de que o scripting está habilitado e que as notificações não estão ativadas (ou que sua aplicação saiba lidar com elas).
    *   **No Servidor SAP:** O administrador do SAP precisa habilitar o scripting no lado do servidor através do parâmetro de perfil `sapgui/user_scripting = TRUE`. **A biblioteca não funcionará sem isso.**

## Instalação

Você pode adicionar a biblioteca ao seu projeto .NET (Console App, etc.) usando o pacote NuGet.

```bash
dotnet add package Inocentini.SAPInteraction.Library --version 1.0.0
# (Ajuste a versão conforme necessário)
```

## Como Usar

A forma recomendada de consumir a biblioteca é através da fábrica `SapClientFactory` e da interface `ISapClient`.

```csharp
using SAPGui.Library; // Namespace da fábrica e interface
using System;
using System.Data;

public class Program
{
    public static void Main(string[] args)
    {
        // 1. Verificar se está no Windows
        if (!OperatingSystem.IsWindows())
        {
            Console.WriteLine("ERRO: Esta biblioteca só funciona no Windows com SAP GUI instalado.");
            return;
        }

        Console.WriteLine("--- Iniciando Automação SAP GUI ---");

        // Usar using para garantir o Dispose do cliente
        using ISapClient? sapClient = CreateSapClient();

        if (sapClient == null) return; // Fábrica falhou em criar o cliente

        try
        { 
            // 2. Abrir Conexão SAP
            Console.WriteLine("Abrindo SAP...");
            sapClient.Open(); // Tenta encontrar instância rodando ou iniciar saplogon.exe
            Console.WriteLine("SAP Aberto.");

            // 3. Fazer Login
            Console.WriteLine("Realizando login...");
            // !! Use credenciais válidas !!
            sapClient.Login("SEU_USUARIO", "SUA_SENHA", "SEU_MANDANTE", "PT");
            Console.WriteLine("Login realizado.");

            // 4. Acessar Transação
            string transaction = "VA01";
            Console.WriteLine($"Acessando transação {transaction}...");
            sapClient.GoToTransaction(transaction);
            System.Threading.Thread.Sleep(2000); // Pausa para a tela carregar (idealmente usar espera mais inteligente)
            Console.WriteLine($"Transação {transaction} acessada.");

            // 5. Interagir com Elementos
            string orderTypeFieldId = "wnd[0]/usr/ctxtVBAK-AUART"; // Exemplo de ID
            string orderType = "OR";
            Console.WriteLine($"Preenchendo campo '{orderTypeFieldId}' com '{orderType}'...");
            sapClient.SetTextFieldValue(orderTypeFieldId, orderType);

            Console.WriteLine("Enviando VKey Enter (0)...");
            sapClient.SendVKey(0); // Código para Enter
            System.Threading.Thread.Sleep(1000);

            // 6. Ler Informações
            Console.WriteLine("Lendo status bar...");
            string statusBarText = sapClient.GetStatusBarText();
            Console.WriteLine($"Status Bar: {statusBarText}");

            // 7. Exemplo de tratamento de erro ao buscar elemento
            try
            {
                Console.WriteLine("Tentando clicar em botão inexistente...");
                sapClient.ClickButton("wnd[0]/usr/btnInexistente");
            }
            catch (SAPGui.Core.Exceptions.ElementNotFoundException ex)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"Elemento não encontrado como esperado: {ex.Message}");
                Console.ResetColor();
            }

            Console.WriteLine("\n--- Automação Concluída com Sucesso (Exemplo) ---");

        }
        catch (Exception ex)
        {   // Captura exceções gerais (falha no login, erro COM, etc.)
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"\n--- ERRO DURANTE A EXECUÇÃO --- ");
            Console.WriteLine($"Mensagem: {ex.Message}");
            Console.WriteLine("Stack Trace:");
            Console.WriteLine(ex.StackTrace);
            Console.ResetColor();
        }
        finally
        {
            // 8. Fechar conexão (o using já chama Dispose, que chama Close)
            // Opcionalmente, pode chamar Close explicitamente se necessário antes do fim do bloco using.
            // sapClient?.Close();
            Console.WriteLine("Cliente SAP sendo descartado (Dispose chamado pelo using).");
        }

        Console.WriteLine("\nPressione Enter para sair.");
        Console.ReadLine();
    }

    // Função auxiliar para criar o cliente e tratar erro da fábrica
    private static ISapClient? CreateSapClient()
    {
        try
        {
            return SapClientFactory.CreateClient();
        }
        catch (InvalidOperationException factoryEx)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"ERRO FATAL: Falha ao inicializar a fábrica do SAP Client.");
            Console.WriteLine($"Mensagem: {factoryEx.Message}");
            if (factoryEx.InnerException != null) Console.WriteLine($"Inner: {factoryEx.InnerException.Message}");
            Console.ResetColor();
            return null;
        }
         catch (PlatformNotSupportedException platformEx)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"ERRO FATAL: {platformEx.Message}");
            Console.ResetColor();
            return null;
        }
    }
}
```

## API Pública (`ISapClient`)

A interface `ISapClient` (obtida via `SapClientFactory.CreateClient()`) expõe os seguintes métodos principais:

*   `Open()`: Tenta conectar a uma instância SAP GUI rodando ou inicia `saplogon.exe`.
*   `Login(...)`: Realiza o login com as credenciais fornecidas.
*   `GoToTransaction(string transactionCode)`: Acessa a transação especificada (ex: "VA01").
*   `FindComponentById<T>(string id)`: Encontra um componente SAP pelo seu ID e tipo wrapper (ex: `ISAPTextField`).
*   `FindButtonById`, `FindTextFieldById`, `FindGridViewById`, etc.: Atalhos para encontrar tipos específicos.
*   `SetTextFieldValue(string id, string text)`: Define o texto de um campo.
*   `ClickButton(string id)`: Pressiona um botão.
*   `GetStatusBarText()`: Retorna o texto atual da barra de status.
*   `GetStatusBar()`: Retorna o objeto `ISAPStatusBar`.
*   `GetGridData(string gridId)`: Extrai o conteúdo de um `GuiGridView` para um `DataTable`.
*   `SendVKey(int vkeyCode)`: Envia um código VKey (ex: 0 para Enter, 8 para F8).
*   `GetMainWindow()`: Retorna o wrapper da janela principal (`wnd[0]`).
*   `Close()`: Fecha a conexão SAP.
*   `Dispose()`: Libera os recursos (chamado automaticamente se usar `using`).

## Tratamento de Erros

A biblioteca pode lançar exceções, incluindo:

*   `PlatformNotSupportedException`: Ao tentar usar em SO não-Windows.
*   `InvalidOperationException`: Para erros na configuração da fábrica ou estado inválido (ex: tentar interagir antes de `Open`/`Login`).
*   `SAPGui.Core.Exceptions.ElementNotFoundException`: Quando um elemento (botão, campo) não é encontrado pelo ID fornecido.
*   `System.Exception`: Pode encapsular erros de COM Interop, falhas de login (verificar mensagem/inner exception), ou outros erros inesperados.

É recomendado usar blocos `try...catch` para lidar com essas situações.

## Arquitetura

A biblioteca utiliza Clean Architecture para separar responsabilidades:

*   **Core:** Define contratos (interfaces) e dados (modelos) sem saber como serão implementados.
*   **Infrastructure:** Implementa os contratos usando tecnologias específicas (COM Interop) e configura a Injeção de Dependência.
*   **Library:** Fornece uma fachada simples (`ISapClient`) e uma fábrica (`SapClientFactory`) para facilitar o consumo, escondendo os detalhes da DI e da infraestrutura.

## Testes

A solução inclui um projeto de testes unitários (`tests/SAPInteraction.Tests`) que utiliza MSTest e Moq para testar a classe `SAPClient` isoladamente, mockando a camada de serviço (`ISAPService`).

Para executar os testes:

```bash
dotnet test
```

## Contribuição

Contribuições são bem-vindas. Sinta-se à vontade para abrir Issues ou Pull Requests. 