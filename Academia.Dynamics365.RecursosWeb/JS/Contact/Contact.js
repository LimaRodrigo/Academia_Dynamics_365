// JavaScript source code
if (typeof (Demo) === "undefined") { Demo = {}; }

// "use strict"
Demo.Contact = {

    //Definições de contexto
    formContext: {},
    saveContext: {},
    globalContext: Xrm.Utility.getGlobalContext(),

    //Definição dos Campos
    campoLimiteCredito: "creditlimit",
    campoCotacaoDolar: "ptr_cotacaodolar",
    campoLimiteCreditoBr: "ptr_limitedecreditobr",
    campoConta: "parentcustomerid",
    campoPreferenciaContato: "preferredcontactmethodcode",
    campoNome: "firstname",
    campoAniversario: "birthdate",
    campoSuspensaoCredito: "creditonhold",
    campoEmail: "emailaddress1",
    campoTelefoneComercial: "telephone1",
    campoFax: "fax",


    //Valores Globais
    valLimiteCredito: null,
    valConta: null,
    valPreferenciaContato: null,
    valNome: null,
    valAniversario: null,
    valSuspensaoCredito: null,
    valIdOportunidade: null,

    /**
     * Chama todos as funções no carregamento da página
     * @param {object} executionContext 
     */
    OnLoad: function (executionContext) {
        //Chamada de atribuição do formContext 
        //https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/executioncontext/getformcontext
        Demo.Contact.formContext = executionContext.getFormContext();
        Demo.Contact.OnChange();
        Demo.Contact.RecuperarValoresTipagem();
    },

    /**
     * Função responsável pela configuração dos eventos de modificações dos campos, ela deve sempre ser chamada no carregamento da página
     */
    OnChange: function () {
        //Adiciona evento onchange em um campo
        //https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/addonchange
        Demo.Contact.formContext.getAttribute(Demo.Contact.campoLimiteCredito).addOnChange(Demo.Contact.ValidarCriacao);
        Demo.Contact.formContext.getAttribute(Demo.Contact.campoTelefoneComercial).addOnChange(Demo.Contact.ValidarTelefone);
        Demo.Contact.formContext.getAttribute(Demo.Contact.campoFax).addOnChange(Demo.Contact.ValidarFax);
        Demo.Contact.formContext.getAttribute(Demo.Contact.campoSuspensaoCredito).addOnChange(Demo.Contact.Update);

    },

    /**
     * Função reponsável por chamadas no salvamento da página
     * @param {object} executionContext 
     */
    OnSave: function (executionContext) {
        Demo.Contact.saveContext = executionContext.getFormContext();
        Demo.Contact.GetDadosApi();
    },


    /**
     * DEMONSTRAÇÃO DE TIPAGEM DO DYNAMICS
     */
    RecuperarValoresTipagem: function () {

        /*Get campos dos tipos */

        //Data
        console.log("**** Data valores ****");
        Demo.Contact.valAniversario = Demo.Contact.formContext.getAttribute(Demo.Contact.campoAniversario).getValue();
        console.log(Demo.Contact.valAniversario);

        //Booleanos 
        console.log("**** Booleanos valores ****");
        Demo.Contact.valSuspensaoCredito = Demo.Contact.formContext.getAttribute(Demo.Contact.campoSuspensaoCredito).getValue();
        console.log(Demo.Contact.valSuspensaoCredito);

        //Numericos (Moeda, Decimais e inteiros)
        console.log("**** Númericos valores ****");
        Demo.Contact.valLimiteCredito = Demo.Contact.formContext.getAttribute(Demo.Contact.campoLimiteCredito).getValue();
        console.log(Demo.Contact.valLimiteCredito);

        //Texto
        console.log("**** Texto valores ****");
        Demo.Contact.valNome = Demo.Contact.formContext.getAttribute(Demo.Contact.campoNome).getValue();
        console.log(Demo.Contact.valNome);

        //Optionset
        console.log("**** OptionSet valores ****");
        Demo.Contact.valPreferenciaContato = Demo.Contact.formContext.getAttribute(Demo.Contact.campoPreferenciaContato).getValue();
        console.log(Demo.Contact.valPreferenciaContato);
        console.log(Demo.Contact.formContext.getAttribute(Demo.Contact.campoPreferenciaContato).getSelectedOption());

        //Lookup
        console.log("**** LookUp Valores *****");
        Demo.Contact.valConta = Demo.Contact.formContext.getAttribute(Demo.Contact.campoConta).getValue();
        console.log(Demo.Contact.valConta);

    },

    /**
     * DEMONSTRAÇÃO DO CRUD 
     */
    //https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/xrm-webapi/createrecord
    Create: async function (valor) {
        try {
            let idContato = Demo.Contact.formContext.data.entity.getId().replace("{", "").replace("}", "");
            let idConta = Demo.Contact.valConta[0].id.replace("{", "").replace("}", "");
            //Objeto a ser criado
            var Oportunidade = {};
            Oportunidade["parentcontactid@odata.bind"] = "/contacts(" + idContato + ")";
            Oportunidade["parentaccountid@odata.bind"] = "/accounts(" + idConta + ")";
            Oportunidade.name = "Limite Credito - " + Demo.Contact.valConta[0].name + " - Potencial grande";
            Oportunidade.purchasetimeframe = 4; //OptionSet - Desconhecido  
            Oportunidade.estimatedvalue = valor;
            //chamada assincrona para criação
            let result = await Demo.Helper.CriarRegistroAsync("opportunity", Oportunidade);
            Demo.Contact.valIdOportunidade = result.id;
            let msg = "Registro " + result.id + " Criado!";
            Demo.Helper.ExibirAlertaSemConfirmacao(msg, "Sucesso");
            console.log(result);

        } catch (error) {
            let retorno = JSON.parse(error.raw);
            Demo.Helper.ExibirAlertaSemConfirmacao(retorno.message, "Erro");
        }
    },
    //https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/xrm-webapi/retrieverecord
    Read: async function () {
        try {
            let Oportunidade = await Demo.Helper.RetornarRegistroAsync("opportunity", Demo.Contact.valIdOportunidade, "?$select=name,purchasetimeframe");
            let Texto = "Titulo " + Oportunidade.name;
            Texto += "\nPeríodo de Compra " + Oportunidade["purchasetimeframe@OData.Community.Display.V1.FormattedValue"];
            //@OData.Community.Display.V1.FormattedValue - Em caso de campo option set para recupera o label utilizamos essa conotação, 
            //caso passe só nome da propridade ira retornar o valor numerico
            Demo.Helper.ExibirAlertaSemConfirmacao(Texto, "Oportunidade");

        } catch (error) {
            Demo.Contact.TratarErroWebApi(error);
        }
    },
    //https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/xrm-webapi/retrievemultiplerecords
    ReadMultiplos: async function () {
        //add filter
        //https://docs.microsoft.com/pt-br/azure/search/search-query-odata-logical-operators
        url = "?$filter=estimatedvalue gt 20000";
        //add select
        url = url.concat("&$select=name, estimatedvalue");
        //add order
        url = url.concat("&$orderby=estimatedvalue desc");

        try {
            let result = await Demo.Helper.RetornarRegistrosAsync("opportunity", url);
            let msg = "";
            for (let i = 0; i < result.length; i++) {
                const item = result[i];
                msg += item["estimatedvalue@OData.Community.Display.V1.FormattedValue"] + " - Titulo - " + item.name + "\n\n";
            }

            Demo.Helper.ExibirAlertaSemConfirmacao(msg, "RESULTADOS OPORTUNIDADES - Xrm.WebApi", 800, 500);

        } catch (error) {
            Demo.Contact.TratarErroWebApi(error);
        }
    },
    //https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/xrm-webapi/updaterecord
    Update: async function () {
        try {
            let Oportunidade = {
                name: "[SUSPENÇÃO DE CREDITO] - Atualizado " + Date.toLocaleString(),
                statuscode: 2, //OptionSet - SUSPENSO
            };
            let result = await Demo.Helper.AtualizarRegistroAsync("opportunity", Demo.Contact.valIdOportunidade, Oportunidade);
            let msg = "Registro " + result.id + " Atualizado!";
            Demo.Helper.ExibirAlertaSemConfirmacao(msg, "Oportunidade Atualizada");
        } catch (error) {
            Demo.Contact.TratarErroWebApi(error);
        }
    },
    //https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/xrm-webapi/deleterecord
    Delete: async function () {
        try {
            await Demo.Helper.DeletarRegistroAsync("opportunity", Demo.Contact.valIdOportunidade);
            Demo.Helper.ExibirAlertaSemConfirmacao("Registro Deletado", "Sucesso");
        } catch (error) {
            Demo.Contact.TratarErroWebApi(error);
        }
    },



    /**
     * DEMOSTRAÇÕES VALIDAÇÕES *
     * Existem dois tipos de de notificação utilizadas para validações no form e no campo
     * com 3 classificações: Aviso (WARNING), Informação (INFO), Erro(ERROR)
     * Observações
     * Enquanto as notificação de CAMPO de erro estiverem ativas no formulário o Dynamics CE não permite o salvamento 
     * notificações de FORM de erro permitem salvamento 
     * 
     * Notificação no campo
     * https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/formcontext-ui/setformnotification
     */
    ValidarTelefone: function () {
        let telefone = Demo.Contact.formContext.getAttribute(Demo.Contact.campoTelefoneComercial).getValue();
        let pattern = /^[\+]?[(]?[0-9]{3}[)]?[-\s\.]?[0-9]{3}[-\s\.]?[0-9]{4,6}$/im;

        if (pattern.test(telefone)) {
            Demo.Contact.formContext.ui.clearFormNotification("ValidaTel");
        }
        else {
            Demo.Contact.formContext.ui.setFormNotification("Telefone inválido", "ERROR", "ValidaTel"); //ERROR / WARNING / INFO

        }
    },
    /**
     * Notificações do Tipo Form
     * https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/controls/setnotification
     */
    ValidarFax: function () {
        let email = Demo.Contact.formContext.getAttribute(Demo.Contact.campoEmail).getValue();
        let pattern = /^[\w-]+(\.[\w-]+)*@(([A-Za-z\d][A-Za-z\d-]{0,61}[A-Za-z\d]\.)+[A-Za-z]{2,6}|\[\d{1,3}(\.\d{1,3}){3}\])$/;

        if (pattern.test(email)) {
            Demo.Contact.formContext.getControl(Demo.Contact.campoEmail).clearNotification("ValidaEmail");
        }
        else {
            Demo.Contact.formContext.getControl(Demo.Contact.campoEmail).setNotification("Email inválido", "ValidaEmail");
        }
    },



    /**
     * DEMONSTRAÇÃO DE UTILIZAR API REST 
     * https://docs.microsoft.com/en-us/previous-versions/dynamicscrm-2016/developers-guide/mt593051(v=crm.8)
     * https://docs.microsoft.com/en-us/previous-versions/dynamicscrm-2016/developers-guide/mt770366(v=crm.8)
     * GET DADOS 
     */
    GetDadosApi: async function () {
        let url = "/api/data/v9.1/";
        //Add entidade
        url = url.concat("opportunities");// o nome da entidade é sempre pluralizado ex. account vira accounts, opportunity vira opportunities
        //add filter
        //https://docs.microsoft.com/pt-br/azure/search/search-query-odata-logical-operators
        url = url.concat("?$filter=estimatedvalue gt 20000");
        //add select
        url = url.concat("&$select=name, estimatedvalue");
        //add order
        url = url.concat("&$orderby=estimatedvalue desc");
        try {

            let result = await Demo.Helper.GetApi(url);

            let msg = "";
            for (let i = 0; i < result.value.length; i++) {
                const item = result.value[i];
                msg += item["estimatedvalue@OData.Community.Display.V1.FormattedValue"] + " - Titulo - " + item.name + "\n\n";
            }

            Demo.Helper.ExibirAlertaSemConfirmacao(msg, "RESULTADOS OPORTUNIDADES", 800, 500);


        } catch (error) {
            Demo.Helper.ExibirAlertaSemConfirmacao(error, "Erro");
        }

    },



    ValidarCriacao: function () {
        let limite = Demo.Contact.formContext.getAttribute(Demo.Contact.campoLimiteCredito).getValue();
        if (limite !== null && limite > 5000)
            Demo.Contact.Create(limite);
    },
    TratarErroWebApi: function (e) {
        let retorno = JSON.parse(e.raw);
        Demo.Helper.ExibirAlertaSemConfirmacao(retorno.message, "Erro");
    }
};