using Microsoft.Xrm.Sdk;
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
            ContarClientes();
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
