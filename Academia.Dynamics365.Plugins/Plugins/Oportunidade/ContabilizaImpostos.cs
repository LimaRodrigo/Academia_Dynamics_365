using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Academia.Dynamics365.Plugins.Oportunidade
{
    public class ContabilizaImpostos : IPlugin
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
            if (context.MessageName.ToLower().Trim() == "create")
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity entityContext = context.InputParameters["Target"] as Entity;

                    //Definir Evento e Modo
                    //https://docs.microsoft.com/en-us/dotnet/api/microsoft.xrm.sdk.ipluginexecutioncontext.stage?view=dynamics-general-ce-9#microsoft-xrm-sdk-ipluginexecutioncontext-stage
                    //https://docs.microsoft.com/en-us/dotnet/api/microsoft.xrm.sdk.iexecutioncontext.mode?view=dynamics-general-ce-9#microsoft-xrm-sdk-iexecutioncontext-mode
                    if (context.Stage == 20 && context.Mode == 0)
                    {
                        Contabilizar(entityContext);
                    }

                }
            }
            if (context.MessageName.ToLower().Trim() == "update")
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity entityContext = context.InputParameters["Target"] as Entity;
                    Entity PreImage = context.PreEntityImages["preImage"] as Entity;

                    //Definir Evento e Modo
                    //https://docs.microsoft.com/en-us/dotnet/api/microsoft.xrm.sdk.ipluginexecutioncontext.stage?view=dynamics-general-ce-9#microsoft-xrm-sdk-ipluginexecutioncontext-stage
                    //https://docs.microsoft.com/en-us/dotnet/api/microsoft.xrm.sdk.iexecutioncontext.mode?view=dynamics-general-ce-9#microsoft-xrm-sdk-iexecutioncontext-mode
                    if (context.Stage == 20 && context.Mode == 0)
                    {
                        Contabilizar(entityContext, PreImage);
                    }

                }
            }
        }

        private void Contabilizar(Entity entityContext, Entity preImage = null)
        {
            decimal TotalImpostos = decimal.Zero;
            decimal PrecoUnitario = decimal.Zero;
            decimal Quantidade = decimal.Zero;
            string DetalhesImpostos = string.Empty;
            Guid ProductId = Guid.Empty;

            if (entityContext.Contains("productid"))
                ProductId = entityContext.GetAttributeValue<EntityReference>("productid").Id;
            else if (preImage.Contains("productid"))
                ProductId = preImage.GetAttributeValue<EntityReference>("productid").Id;

            if (entityContext.Contains("priceperunit"))
                PrecoUnitario = entityContext.GetAttributeValue<Money>("priceperunit").Value;
            else if (preImage.Contains("priceperunit"))
                PrecoUnitario = preImage.GetAttributeValue<Money>("priceperunit").Value;

            if (entityContext.Contains("quantity"))
                Quantidade = entityContext.GetAttributeValue<decimal>("quantity");
            else if (preImage.Contains("quantity"))
                Quantidade = preImage.GetAttributeValue<decimal>("quantity");

            tracing.Trace("Atribuições de Quantidade e Preço OK");


            //https://docs.microsoft.com/en-us/dotnet/api/microsoft.xrm.sdk.query.queryexpression?view=dynamics-general-ce-9
            QueryExpression query = new QueryExpression("ptr_impostos");
            query.ColumnSet.AddColumns("ptr_percentual", "ptr_produto", "ptr_nome");
            query.AddOrder("ptr_nome", OrderType.Ascending);
            query.Criteria.AddCondition("ptr_produto", ConditionOperator.Equal, ProductId);

            //Recuperação de registros da tabela Impostos
            var Impostos = serviceGlobal.RetrieveMultiple(query);

            if (Impostos != null && Impostos.Entities.Count > 0)
            {
                tracing.Trace("Localizou {0} Impostos ", Impostos.Entities.Count);
                foreach (var item in Impostos.Entities)
                {
                    if (!item.Contains("ptr_percentual")) continue;
                    decimal ValorCalculado = (PrecoUnitario * Quantidade) * (item.GetAttributeValue<decimal>("ptr_percentual") / 100);

                    DetalhesImpostos += string.Format("{0} - {1}\n", item["ptr_nome"].ToString(), ValorCalculado.ToString("C", CultureInfo.CreateSpecificCulture("en-US")));
                    TotalImpostos += ValorCalculado;

                }

                //Atribuição de valores no registro
                entityContext["tax"] = new Money(TotalImpostos);
                entityContext["ptr_detalhesimpostos"] = DetalhesImpostos;
            }

        }
    }
}
