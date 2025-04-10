using NUnit.Framework;
using SAPGui.Core.Models;

namespace SAPGui.Core.Tests.Models;

[TestFixture]
public class SAPCredentialsTests
{
    [Test]
    public void Properties_ShouldBeSettableAndGettable()
    {
        // Arrange
        var credentials = new SAPCredentials();
        const string testUser = "testUser";
        const string testPassword = "testPassword";
        const string testClient = "100";
        const string testLanguage = "EN";

        // Act
        credentials.User = testUser;
        credentials.Password = testPassword;
        credentials.Client = testClient;
        credentials.Language = testLanguage;

        // Assert
        Assert.That(credentials.User, Is.EqualTo(testUser));
        Assert.That(credentials.Password, Is.EqualTo(testPassword));
        Assert.That(credentials.Client, Is.EqualTo(testClient));
        Assert.That(credentials.Language, Is.EqualTo(testLanguage));
    }

    [Test]
    public void Properties_ShouldInitializeAsEmptyStrings()
    {
        // Arrange & Act
        var credentials = new SAPCredentials();

        // Assert
        Assert.That(credentials.User, Is.EqualTo(string.Empty));
        Assert.That(credentials.Password, Is.EqualTo(string.Empty));
        Assert.That(credentials.Client, Is.EqualTo(string.Empty));
        Assert.That(credentials.Language, Is.EqualTo(string.Empty));
    }
} 