using Moq;
using NUnit.Framework;
using SAPGui.Core.Interfaces;
using SAPGui.Core.Interfaces.Components;
using SAPGui.Core.Models;
using SapGui.Infrastructure.Providers;
using SapGui.Infrastructure.Services;
using System;

namespace SapGui.Infrastructure.Tests.Services;

[TestFixture]
public class SAPServiceTests
{
    // Mocks das dependências injetadas
    private Mock<ISapComProvider> _mockComProvider = null!;
    private Mock<ISapComponentWrapperFactory> _mockWrapperFactory = null!;

    // Mocks dos wrappers retornados pela factory
    private Mock<ISAPWindow> _mockLoginWindow = null!;
    private Mock<ISAPWindow> _mockMainWindow = null!;
    private Mock<ISAPTextField> _mockUserField = null!;
    private Mock<ISAPTextField> _mockPassField = null!;
    private Mock<ISAPTextField> _mockClientField = null!;
    private Mock<ISAPTextField> _mockLangField = null!;
    private Mock<ISAPTextField> _mockOkcdField = null!;
    private Mock<ISAPStatusBar> _mockStatusBar = null!;

    // Mocks dos objetos COM "brutos" retornados pelo provider
    private Mock<object> _mockSessionComObject = null!;
    private Mock<object> _mockLoginWindowComObject = null!;
    private Mock<object> _mockMainWindowComObject = null!;
    private Mock<object> _mockUserFieldComObject = null!;
    private Mock<object> _mockPassFieldComObject = null!;
    private Mock<object> _mockClientFieldComObject = null!;
    private Mock<object> _mockLangFieldComObject = null!;
    private Mock<object> _mockOkcdFieldComObject = null!;
    private Mock<object> _mockStatusBarComObject = null!;

    // A instância do serviço sob teste
    private SAPService _service = null!;

    [SetUp]
    public void Setup()
    {
        // Inicializa mocks das dependências
        _mockComProvider = new Mock<ISapComProvider>();
        _mockWrapperFactory = new Mock<ISapComponentWrapperFactory>();

        // Inicializa mocks dos wrappers
        _mockLoginWindow = new Mock<ISAPWindow>();
        _mockMainWindow = new Mock<ISAPWindow>();
        _mockUserField = new Mock<ISAPTextField>();
        _mockPassField = new Mock<ISAPTextField>();
        _mockClientField = new Mock<ISAPTextField>();
        _mockLangField = new Mock<ISAPTextField>();
        _mockOkcdField = new Mock<ISAPTextField>();
        _mockStatusBar = new Mock<ISAPStatusBar>();

        // Inicializa mocks dos objetos COM
        _mockSessionComObject = new Mock<object>();
        _mockLoginWindowComObject = new Mock<object>();
        _mockMainWindowComObject = new Mock<object>();
        _mockUserFieldComObject = new Mock<object>();
        _mockPassFieldComObject = new Mock<object>();
        _mockClientFieldComObject = new Mock<object>();
        _mockLangFieldComObject = new Mock<object>();
        _mockOkcdFieldComObject = new Mock<object>();
        _mockStatusBarComObject = new Mock<object>();

        // Configuração padrão dos mocks dos wrappers
        _mockLoginWindow.Setup(w => w.Id).Returns("wnd[0]_login"); // IDs diferentes para clareza
        _mockLoginWindow.Setup(w => w.Text).Returns("SAP Login Screen");
        _mockMainWindow.Setup(w => w.Id).Returns("wnd[0]_main");
        _mockMainWindow.Setup(w => w.Text).Returns("SAP Easy Access - Logged In");
        _mockStatusBar.Setup(s => s.MessageType).Returns("S"); // Sucesso por padrão

        // Cria a instância do serviço com os mocks das dependências
        _service = new SAPService(_mockComProvider.Object, _mockWrapperFactory.Object);

        // Configuração comum para que FindById/FindInWindow retorne os mocks corretos
        // via _mockWrapperFactory.CreateWrapper
        SetupWrapperFactoryMock(_mockLoginWindowComObject, _mockLoginWindow);
        SetupWrapperFactoryMock(_mockMainWindowComObject, _mockMainWindow);
        SetupWrapperFactoryMock(_mockUserFieldComObject, _mockUserField);
        SetupWrapperFactoryMock(_mockPassFieldComObject, _mockPassField);
        SetupWrapperFactoryMock(_mockClientFieldComObject, _mockClientField);
        SetupWrapperFactoryMock(_mockLangFieldComObject, _mockLangField);
        SetupWrapperFactoryMock(_mockOkcdFieldComObject, _mockOkcdField);
        SetupWrapperFactoryMock(_mockStatusBarComObject, _mockStatusBar);

        // Configuração comum para que _comProvider retorne os mocks COM corretos
        // Simula o fluxo de OpenSAP encontrando uma sessão
        var mockApp = new Mock<object>();
        var mockEngine = new Mock<object>();
        var mockConnection = new Mock<object>();
        _mockComProvider.Setup(p => p.GetApplication()).Returns(mockApp.Object);
        _mockComProvider.Setup(p => p.GetScriptingEngine(mockApp.Object)).Returns(mockEngine.Object);
        _mockComProvider.Setup(p => p.GetConnectionCount(mockEngine.Object)).Returns(1);
        _mockComProvider.Setup(p => p.GetConnection(mockEngine.Object, 0)).Returns(mockConnection.Object);
        _mockComProvider.Setup(p => p.GetSessionCount(mockConnection.Object)).Returns(1);
        _mockComProvider.Setup(p => p.GetSession(mockConnection.Object, 0)).Returns(_mockSessionComObject.Object);

        // Simula FindById retornando os objetos COM mockados (a factory os transformará em wrappers mockados)
        SetupComProviderFindById("wnd[0]", _mockMainWindowComObject.Object);
        SetupComProviderFindById("wnd[0]_login/usr/txtRSYST-BNAME", _mockUserFieldComObject.Object);
        SetupComProviderFindById("wnd[0]_login/usr/pwdRSYST-BCODE", _mockPassFieldComObject.Object);
        SetupComProviderFindById("wnd[0]_login/usr/txtRSYST-MANDT", _mockClientFieldComObject.Object);
        SetupComProviderFindById("wnd[0]_login/usr/txtRSYST-LANGU", _mockLangFieldComObject.Object);
        SetupComProviderFindById("wnd[0]_main/okcd", _mockOkcdFieldComObject.Object);
        SetupComProviderFindById("wnd[0]_main/sbar", _mockStatusBarComObject.Object);

        // Simula GetActiveWindow retornando a janela de login COM
        _mockComProvider.Setup(p => p.GetActiveWindow(_mockSessionComObject.Object)).Returns(_mockLoginWindowComObject.Object);
    }

    // Método auxiliar para configurar o mock da WrapperFactory
    private void SetupWrapperFactoryMock<TWrapper>(Mock<object> comObjectMock, Mock<TWrapper> wrapperMock)
        where TWrapper : class, ISAPComponent
    {
        _mockWrapperFactory.Setup(f => f.CreateWrapper<TWrapper>(comObjectMock.Object)).Returns(wrapperMock.Object);
        _mockWrapperFactory.Setup(f => f.CreateWrapper(comObjectMock.Object)).Returns(wrapperMock.Object);
    }

    // Método auxiliar para configurar o mock do ComProvider.FindComComponentById
    private void SetupComProviderFindById(string id, object? comObjectToReturn)
    {
        _mockComProvider.Setup(p => p.FindComComponentById(_mockSessionComObject.Object, id)).Returns(comObjectToReturn);
    }

    [TearDown]
    public void TearDown()
    {
        _service?.Dispose();
    }

    [Test]
    public void LoginSAP_GivenValidCredentials_SetsFieldsAndSendsEnterAndReturnsMainWindow()
    {
        // Arrange
        var credentials = new SAPCredentials { User = "u", Password = "p", Client = "c", Language = "l" };

        // Act: Chama o método refatorado
        ISAPWindow? resultWindow = _service.LoginSAP(credentials);

        // Assert
        // Verifica se os campos foram preenchidos através dos mocks dos wrappers
        _mockUserField.VerifySet(f => f.Text = credentials.User, Times.Once);
        _mockPassField.VerifySet(f => f.Text = credentials.Password, Times.Once);
        _mockClientField.VerifySet(f => f.Text = credentials.Client, Times.Once);
        _mockLangField.VerifySet(f => f.Text = credentials.Language, Times.Once);

        // Verifica se Enter foi enviado na janela de login mockada
        _mockLoginWindow.Verify(w => w.SendVKey(0), Times.Once);

        // Verifica se a janela principal foi buscada e retornada
        _mockComProvider.Verify(p => p.FindComComponentById(_mockSessionComObject.Object, "wnd[0]"), Times.AtLeastOnce); // Chamado em OpenSAP e após login
        Assert.That(resultWindow, Is.SameAs(_mockMainWindow.Object));

        // Verifica se a status bar da janela principal foi verificada
        _mockComProvider.Verify(p => p.FindComComponentById(_mockSessionComObject.Object, "wnd[0]_main/sbar"), Times.Once);
    }

    [Test]
    public void LoginSAP_WhenStatusBarReturnsError_ThrowsException()
    {
        // Arrange
        var credentials = new SAPCredentials { User = "u", Password = "p", Client = "c" };
        _mockStatusBar.Setup(s => s.MessageType).Returns("E"); // Simula erro
        _mockStatusBar.Setup(s => s.Text).Returns("Login failed");

        // Act & Assert
        var ex = Assert.Throws<Exception>(() => _service.LoginSAP(credentials));
        Assert.That(ex?.Message, Does.Contain("Erro SAP detectado na Status Bar após login: Login failed"));
    }

    [Test]
    public void AccessTransaction_GivenTransactionCode_SetsOkcdFieldAndSendsEnter()
    {
        // Arrange
        const string transactionCode = "VA01";
        const string expectedOkcdText = "/n" + transactionCode;

        // Garante que temos uma _mainWindow (simulando que OpenSAP ou LoginSAP foi chamado antes)
        _service.GetType().GetField("_mainWindow", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)?.SetValue(_service, _mockMainWindow.Object);

        // Act
        _service.AccessTransaction(transactionCode);

        // Assert
        _mockOkcdField.VerifySet(f => f.Text = expectedOkcdText, Times.Once);
        _mockMainWindow.Verify(w => w.SendVKey(0), Times.Once);
    }

    [Test]
    public void AccessTransaction_WhenMainWindowIsNull_ThrowsInvalidOperationException()
    {
        // Arrange (Garante que _mainWindow é nulo)
        _service.GetType().GetField("_mainWindow", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)?.SetValue(_service, null);

        // Act & Assert
        Assert.Throws<InvalidOperationException>(() => _service.AccessTransaction("VA01"));
    }

    [Test]
    public void Dispose_ReleasesSessionComObject()
    {
        // Arrange: Simula que OpenSAP foi chamado e obteve um objeto COM de sessão
        _service.GetType().GetField("_sapSessionComObject", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)?.SetValue(_service, _mockSessionComObject.Object);

        // Act
        _service.Dispose();

        // Assert
        _mockComProvider.Verify(p => p.ReleaseComObject(_mockSessionComObject.Object), Times.Once);
        // Verifica se o campo interno foi limpo (usando Reflection para teste)
        var sessionField = _service.GetType().GetField("_sapSessionComObject", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
        Assert.That(sessionField?.GetValue(_service), Is.Null);
    }
} 