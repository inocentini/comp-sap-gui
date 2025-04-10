using NUnit.Framework;
using SAPGui.Core.Models;

namespace SAPGui.Core.Tests.Models;

[TestFixture]
public class ElementLocatorTests
{
    [Test]
    public void ById_ShouldCreateLocatorWithCorrectId()
    {
        // Arrange
        const string expectedId = "wnd[0]/usr/txtField";

        // Act
        var locator = ElementLocator.ById(expectedId);

        // Assert
        Assert.That(locator, Is.Not.Null);
        Assert.That(locator.Id, Is.EqualTo(expectedId));
        Assert.That(locator.Type, Is.Null);
        Assert.That(locator.Text, Is.Null);
    }

    [Test]
    public void ByTypeAndText_ShouldCreateLocatorWithCorrectTypeAndText()
    {
        // Arrange
        const string expectedType = "GuiButton";
        const string expectedText = "Continue";

        // Act
        var locator = ElementLocator.ByTypeAndText(expectedType, expectedText);

        // Assert
        Assert.That(locator, Is.Not.Null);
        Assert.That(locator.Id, Is.Null);
        Assert.That(locator.Type, Is.EqualTo(expectedType));
        Assert.That(locator.Text, Is.EqualTo(expectedText));
    }

    [Test]
    public void Properties_ShouldBeSettableAndGettable()
    {
        // Arrange
        var locator = new ElementLocator();
        const string testId = "testId";
        const string testType = "testType";
        const string testText = "testText";

        // Act
        locator.Id = testId;
        locator.Type = testType;
        locator.Text = testText;

        // Assert
        Assert.That(locator.Id, Is.EqualTo(testId));
        Assert.That(locator.Type, Is.EqualTo(testType));
        Assert.That(locator.Text, Is.EqualTo(testText));
    }
} 