using SAPGui.Core.Interfaces.Components;
using SapGui.Infrastructure.Wrappers;
using System;
using System.Runtime.InteropServices;

namespace SapGui.Infrastructure.Providers
{
    /// <summary>
    /// Implementação de ISapComponentWrapperFactory que cria instâncias dos wrappers concretos.
    /// </summary>
    internal class SapComponentWrapperFactory : ISapComponentWrapperFactory
    {
        public ISAPComponent? CreateWrapper(object? comObject)
        {
            if (comObject == null || !Marshal.IsComObject(comObject))
                return null;

            dynamic dynamicObject = comObject;
            try
            {
                string type = dynamicObject.Type; // Obtém o tipo do componente COM

                // Mapeia o tipo COM para a classe Wrapper correspondente
                switch (type)
                {
                    case "GuiMainWindow":
                    case "GuiModalWindow":
                    case "GuiDialogShell":
                        return new SAPWindowWrapper(comObject);
                    case "GuiStatusbar":
                        return new SAPStatusBarWrapper(comObject);
                    case "GuiTextField":
                    case "GuiCTextField":
                    case "GuiPasswordField":
                        return new SAPTextFieldWrapper(comObject);
                    case "GuiButton":
                        return new SAPButtonWrapper(comObject);
                    case "GuiGridView":
                        return new SAPGridViewWrapper(comObject);
                    case "GuiCheckBox":
                        return new SAPCheckboxWrapper(comObject);
                    // Adicione outros tipos aqui conforme necessário

                    default:
                        Console.WriteLine($"Aviso: Tipo de componente SAP não mapeado na fábrica: {type}. Não foi possível criar wrapper específico.");
                        // Poderíamos retornar um wrapper genérico se tivéssemos um
                        return null;
                }
            }
            catch (COMException ex)
            { Console.WriteLine($"Erro COM ao obter tipo ou criar wrapper: {ex.Message}"); return null; }
            catch (Exception ex)
            { Console.WriteLine($"Erro ao criar wrapper: {ex.Message}"); return null; }
        }

        public T? CreateWrapper<T>(object? comObject) where T : class, ISAPComponent
        {
            var wrapper = CreateWrapper(comObject); // Cria o wrapper genérico primeiro
            if (wrapper is T typedWrapper)
            {
                return typedWrapper;
            }

            if (wrapper != null)
            {
                // Se criamos um wrapper, mas não é do tipo T esperado, logamos um aviso.
                Console.WriteLine($"Aviso: O wrapper criado '{wrapper.GetType().Name}' não corresponde ao tipo esperado '{typeof(T).Name}'.");
                // Considerar liberar o 'wrapper' aqui se não for ser usado?
                // Marshal.ReleaseComObject(comObject); // CUIDADO: pode ser prematuro.
            }
            else if (comObject != null)
            {
                // Se não conseguimos nem criar um wrapper genérico (tipo desconhecido?)
                Console.WriteLine($"Aviso: Não foi possível criar um wrapper para o objeto COM fornecido (Tipo COM: {TryGetComType(comObject)}), logo não pode ser convertido para '{typeof(T).Name}'.");
            }

            return null; // Retorna null se a conversão de tipo falhar
        }

        // Método auxiliar para tentar obter o tipo COM sem lançar exceção dura
        private string TryGetComType(object comObject)
        {
            try
            {
                return ((dynamic)comObject).Type ?? "[Tipo COM não disponível]";
            }
            catch { return "[Erro ao obter tipo COM]"; }
        }
    }
} 