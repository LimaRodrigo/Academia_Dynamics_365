using Academia.Dynamics365.Backend.Facade;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Academia.Dynamics365.Plugins.Contato
{
    public class ValidaCpf : IPlugin
    {
        public IOrganizationService service { get; private set; }
        public IOrganizationService serviceGlobal { get; private set; }
        public ITracingService tracing { get; private set; }

        public void Execute(IServiceProvider serviceProvider)
        {
            //https://docs.microsoft.com/en-us/dotnet/api/microsoft.xrm.sdk.ipluginexecutioncontext?view=dynamics-general-ce-9
            //Contexto de execução
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            //Fabrica de conexões
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            //Service no contexto do usuário
            service = serviceFactory.CreateOrganizationService(context.UserId);
            //Service no contexto Global (usuário System)
            serviceGlobal = serviceFactory.CreateOrganizationService(null);
            //Trancing utilizado para reastreamento de mensagem durante o processo
            tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));


            //Validações 
            if (context.MessageName.ToLower().Trim() == "update")
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity entityContext = context.InputParameters["Target"] as Entity;

                    //Definir Evento e Modo
                    //https://docs.microsoft.com/en-us/dotnet/api/microsoft.xrm.sdk.ipluginexecutioncontext.stage?view=dynamics-general-ce-9#microsoft-xrm-sdk-ipluginexecutioncontext-stage
                    //https://docs.microsoft.com/en-us/dotnet/api/microsoft.xrm.sdk.iexecutioncontext.mode?view=dynamics-general-ce-9#microsoft-xrm-sdk-iexecutioncontext-mode
                    if (context.Stage == 10 && context.Mode == 0)
                    {
                        Validar(entityContext);
                    }

                }
            }
        }

        /// <summary>
        /// Função de validação de CPF
        /// </summary>
        /// <param name="entityContext"></param>
        /// <exception cref="InvalidPluginExecutionException"></exception>
        private void Validar(Entity entityContext)
        {
            tracing.Trace("Entrei na função Validar");

            string cpf = string.Empty;

            //Recupero do Banco
            Entity contato = serviceGlobal.Retrieve(entityContext.LogicalName, entityContext.Id, new ColumnSet("ptr_cpf"));

            //Caso não venha no contexto, atribuo o valor do banco
            if (entityContext.Contains("ptr_cpf"))
                cpf = entityContext["ptr_cpf"].ToString();
            else if (contato.Contains("ptr_cpf"))
                cpf = contato["ptr_cpf"].ToString();


            if (string.IsNullOrEmpty(cpf))
                throw new InvalidPluginExecutionException("CPF Inválido!");

            if (!Helper.ValidarCpf(cpf))
                throw new InvalidPluginExecutionException("CPF Inválido!");

        }
    }
}
