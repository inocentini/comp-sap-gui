using SAPGui.Core.Interfaces.Components;
using System;
using System.Runtime.InteropServices;

namespace SapGui.Infrastructure.Wrappers;

/// <summary>
/// Classe base abstrata para wrappers de componentes SAP COM.
/// Fornece implementação comum para ISAPComponent.
/// </summary>
public abstract class SAPComponentBase : ISAPComponent
{
    protected readonly dynamic SapComObject; // O objeto COM real

    protected SAPComponentBase(object sapComObject) // Recebe object para segurança inicial
    {
        if (sapComObject == null) throw new ArgumentNullException(nameof(sapComObject));
        // Valida se é um objeto COM
        if (!Marshal.IsComObject(sapComObject))
            throw new ArgumentException("O objeto fornecido não é um objeto COM.", nameof(sapComObject));

        SapComObject = sapComObject; // Atribui ao campo dynamic
    }

    // Implementação das propriedades de ISAPComponent
    public virtual string Id => SapComObject.Id;
    public virtual string Type => SapComObject.Type;
    public virtual string Name => SapComObject.Name;

    // A propriedade Text pode precisar ser sobrescrita em classes derivadas
    // pois alguns componentes podem não ter Text ou ter lógica diferente.
    public virtual string Text
    {
        get => HasProperty(SapComObject, "Text") ? SapComObject.Text : string.Empty;
        set
        {
            if (HasProperty(SapComObject, "Text"))
            {
                SapComObject.Text = value;
            }
            else
            {
                // Lançar exceção ou logar aviso se a propriedade não existir?
                Console.WriteLine($"Aviso: Tentativa de definir 'Text' em um componente tipo '{Type}' que não suporta a propriedade.");
            }
        }
    }

    public virtual void SetFocus()
    {
        if (HasMethod(SapComObject, "SetFocus"))
        {
            SapComObject.SetFocus();
        } else {
            Console.WriteLine($"Aviso: Tentativa de chamar 'SetFocus' em um componente tipo '{Type}' que não suporta o método.");
        }
    }

    // Métodos auxiliares para verificar existência de propriedades/métodos no objeto COM
    // Isso ajuda a evitar exceções em tempo de execução com 'dynamic'
    protected static bool HasProperty(dynamic obj, string propertyName)
    {
        try
        {
            var value = obj.GetType().GetProperty(propertyName)?.GetValue(obj, null);
            // Mesmo que GetProperty retorne não nulo, o acesso real pode falhar para COM
            // Uma tentativa de acesso pode ser mais segura, mas potencialmente mais lenta.
             var temp = obj.GetType().InvokeMember(propertyName, System.Reflection.BindingFlags.GetProperty, null, obj, null);
            return true;
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) // Erro comum ao acessar membro inexistente
        {
            return false;
        }
        catch (System.Reflection.TargetInvocationException tie) when (tie.InnerException is COMException)
        {   // Pode acontecer se a propriedade COM não existir
            return false;
        }
         catch (COMException) // Outros erros COM
         {
             return false;
         }
        catch
        {
            return false; // Outras exceções
        }
    }

     protected static bool HasMethod(dynamic obj, string methodName)
    {
        try
        {
            var method = obj.GetType().GetMethod(methodName);
            // Verificar se o método existe é um começo, mas a chamada ainda pode falhar
            // A chamada via InvokeMember pode ser mais robusta para COM
             obj.GetType().InvokeMember(methodName, System.Reflection.BindingFlags.InvokeMethod, null, obj, Array.Empty<object>());
            // Se chegou aqui sem exceção, o método existe e pode ser chamado (com 0 args)
             // CUIDADO: Isso executa o método! Use apenas se a execução for segura/desejada para verificação.
             // Uma abordagem mais segura é verificar GetMethod != null e assumir que funciona, tratando exceção na chamada real.
            // return obj.GetType().GetMethod(methodName) != null; // Abordagem mais segura, mas menos garantida para COM.
            return true; // Se InvokeMember funcionou (com 0 args)
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
        {
            return false;
        }
        catch (System.Reflection.TargetInvocationException tie) when (tie.InnerException is COMException)
        {   // Pode acontecer se o método COM não existir
            return false;
        }
         catch (COMException) // Outros erros COM
         {
             return false;
         }
        catch
        { 
             // Método não encontrado ou erro ao invocar para teste
            return false; 
        }
    }
} 