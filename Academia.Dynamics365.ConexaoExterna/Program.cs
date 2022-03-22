using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Academia.Dynamics365.ConexaoExterna
{
    internal class Program
    {
        private static IOrganizationService Service;

        static void Main(string[] args)
        {
            Console.WriteLine("Inicio");
            Service = GerarService(ConfigurationManager.AppSettings["ConexaoAppSecrect"]);
            Console.WriteLine("Service Conectado");
            //ContarCLientes();
            //ConsultarContas();
            //UpdateContato();
            UpsertContato();
            //UpdateMultipleRequest();
            //UpdateTransactionMultipleRequest();
            Console.WriteLine("pressione qualquer tecla para terminar.");
            Console.ReadKey();

        }


        private static void ContarClientes()
        {
            int Total = 0;
            QueryExpression query = new QueryExpression("account");
            query.ColumnSet.AddColumns("name");
            query.Orders.Add(new OrderExpression("name", OrderType.Ascending));
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;

            while (true)
            {
                var results = Service.RetrieveMultiple(query);
                foreach (var item in results.Entities)
                     Total++;
                if (results.MoreRecords)
                {
                    Console.WriteLine("Pagina: {0}", query.PageInfo.PageNumber);
                    query.PageInfo.PageNumber++;
                    query.PageInfo.PagingCookie = results.PagingCookie;
                }
                else
                    break;
            }
            Console.WriteLine("\nTotal de clientes: {0}", Total);

        }
        private static void UpdateMultipleRequest()
        {
            ExecuteMultipleRequest requestWithResults = new ExecuteMultipleRequest()
            {
                Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError = false,
                    ReturnResponses = true
                },
                Requests = new OrganizationRequestCollection()
            };

            QueryExpression query = new QueryExpression("account");
            query.ColumnSet.AddColumns("name");

            var results = Service.RetrieveMultiple(query);
            int counter = 1;
            foreach (var item in results.Entities)
            {
                item["name"] = item["name"].ToString() + " " + counter;
                UpdateRequest request = new UpdateRequest { Target = item };
                requestWithResults.Requests.Add(request);
                counter++;
            }

            ExecuteMultipleResponse response = (ExecuteMultipleResponse)Service.Execute(requestWithResults);
            Console.WriteLine();

        }
        private static void UpdateTransactionMultipleRequest()
        {
            var requestToTransaction = new ExecuteTransactionRequest()
            {
                Requests = new OrganizationRequestCollection(),
                ReturnResponses = true
            };

            QueryExpression query = new QueryExpression("contact");
            query.ColumnSet.AddColumns("firstname");

            var results = Service.RetrieveMultiple(query);
            int counter = 1;
            foreach (var item in results.Entities)
            {
                item["firstname"] = item["firstname"].ToString() + " " + counter;
                item["ptr_cpf"] = null;
                UpdateRequest request = new UpdateRequest { Target = item };
                requestToTransaction.Requests.Add(request);
                counter++;
            }

            var response = (ExecuteTransactionResponse)Service.Execute(requestToTransaction);
            Console.WriteLine();

        }
        private static void UpdateContato()
        {
            Entity entity = new Entity("contact", Guid.Parse("678c7b32-3f72-ea11-a811-000d3a1b1f2c"));
            entity["firstname"] = "João";
            entity["lastname"] = "da Silva";

            Service.Update(entity);

        }
        private static void UpsertContato()
        {
            string Chave = "12696084017";
            KeyAttributeCollection Keys = new KeyAttributeCollection();
            Keys.Add("ptr_cpf", Chave);

            Entity entity = new Entity("contact", Keys);
            entity["firstname"] = "Rodrigo teste ";
            entity["lastname"] = "da Silva";

            UpsertRequest upsert = new UpsertRequest { Target = entity };

            var result = (UpsertResponse)Service.Execute(upsert);
        }
        private static void ConsultarContas()
        {
            LinkEntity link = new LinkEntity("contact", "account", "parentcustomerid", "accountid", JoinOperator.Inner);
            link.EntityAlias = "Conta";
            link.Columns.AddColumns("name");
            link.LinkCriteria.AddCondition("name", ConditionOperator.Like, "%Trey%");

            QueryExpression query = new QueryExpression("contact");
            query.ColumnSet.AddColumns("fullname");
            query.Orders.Add(new OrderExpression("fullname", OrderType.Ascending));
            query.LinkEntities.Add(link);

            var results = Service.RetrieveMultiple(query);

            foreach (var item in results.Entities)
            {
                Console.WriteLine();
                Console.WriteLine("Contato: {0}", item["fullname"].ToString());
                Console.WriteLine("Conta: {0}", item.GetAttributeValue<AliasedValue>("Conta.name").Value.ToString());
                Console.WriteLine("--------------------------------------------------------");

            }
        }

        /// <summary>
        /// Gerar conexão por meio de uma string de conexão 
        /// </summary>
        /// <param name="connectionString">string de conexão do ambiente 
        /// referência: https://docs.microsoft.com/en-us/powerapps/developer/data-platform/xrm-tooling/use-connection-strings-xrm-tooling-connect </param>
        /// <returns></returns>
        public static IOrganizationService GerarService(string connectionString)
        {
            if (string.IsNullOrEmpty(connectionString))
                throw new Exception(string.Format("Parametro {0} não pode ser nulo ou vazio.", nameof(connectionString)));
            try
            {
                CrmServiceClient crmService = new CrmServiceClient(connectionString);
                return crmService;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }
    }
}
