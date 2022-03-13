using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Academia.Dynamics365.Plugins.Contato
{
    public class GeraTarefa : IPlugin
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
            if (context.MessageName.ToLower().Trim() == "create" && context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity entityContext = context.InputParameters["Target"] as Entity;
                if (entityContext.LogicalName != "contact")
                    throw new InvalidPluginExecutionException("Plugin registrado na entidade errada, processo projetado para entidade de contato.");

                //validação de profundidade
                if (context.Depth > 1) return;

                //Definir Evento e Modo
                //https://docs.microsoft.com/en-us/dotnet/api/microsoft.xrm.sdk.ipluginexecutioncontext.stage?view=dynamics-general-ce-9#microsoft-xrm-sdk-ipluginexecutioncontext-stage
                //https://docs.microsoft.com/en-us/dotnet/api/microsoft.xrm.sdk.iexecutioncontext.mode?view=dynamics-general-ce-9#microsoft-xrm-sdk-iexecutioncontext-mode
                if (context.Stage == 40 && context.Mode == 0)
                {
                    CriarTarefa(entityContext);
                }


            }
        }

        private void CriarTarefa(Entity entityContext)
        {
            Entity entity = new Entity("task");
            entity["subject"] = "Análise de Perfil";
            entity["description"] = "Efetuar analise de perfil do novo contato";
            entity["regardingobjectid"] = entityContext.ToEntityReference();
            entity["actualdurationminutes"] = 60;
            serviceGlobal.Create(entity);

        }
    }
}
