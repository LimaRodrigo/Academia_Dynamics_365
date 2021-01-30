// JavaScript source code
if (typeof (Demo) === "undefined") { Demo = {}; }
// "use strict"
Demo.Contato = {

    //Definições de contexto
    formContext: {},
    saveContext: {},
    globalContext: Xrm.Utility.getGlobalContext(),

    //Definição dos Campos
    campoLimiteCredito: "creditlimit",
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
     * @param {object} formContext 
     */
    OnLoad: function (formContext) {
        /**
         * Chamada de atibição do formContext 
         * https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/executioncontext/getformcontext
         */
        Demo.Contato.formContext = formContext.getFormContext();
        Demo.Contato.OnChange();
        Demo.Contato.RecuperarValoresTipagem();
    },

    /**
     * Função responsável peka configuração dos eventos onchange dos campos, ela deve sempre ser chamada no carregamento da página
     */
    OnChange: function () {
        //Adiciona evento onchange em um campo
        //https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/addonchange
        Demo.Contato.formContext.getAttribute(Demo.Contato.campoLimiteCredito).addOnChange(Demo.Contato.ValidarCriacao);
        Demo.Contato.formContext.getAttribute(Demo.Contato.campoTelefoneComercial).addOnChange(Demo.Contato.ValidarTelefone);
        Demo.Contato.formContext.getAttribute(Demo.Contato.campoFax).addOnChange(Demo.Contato.ValidarFax);
        Demo.Contato.formContext.getAttribute(Demo.Contato.campoSuspensaoCredito).addOnChange(Demo.Contato.Update);

    },

    /**
     * Função reponsável por chamadas no evento OnSave
     * @param {object} saveContext 
     */
    OnSave: function (saveContext) {
        Demo.Contato.saveContext = saveContext.getFormContext();
        Demo.Contato.GetDadosApi();
    },


    /**
     * DEMONSTRAÇÃO DE TIPAGEM DO DYNAMICS
     */
    RecuperarValoresTipagem: function () {

        /*Get campos dos tipos */

        //Data
        console.log("**** Data valores ****");
        Demo.Contato.valAniversario = Demo.Contato.formContext.getAttribute(Demo.Contato.campoAniversario).getValue();
        console.log(Demo.Contato.valAniversario);

        //Booleanos 
        console.log("**** Booleanos valores ****");
        Demo.Contato.valSuspensaoCredito = Demo.Contato.formContext.getAttribute(Demo.Contato.campoSuspensaoCredito).getValue();
        console.log(Demo.Contato.valSuspensaoCredito);

        //Numericos (Moeda, Decimais e inteiros)
        console.log("**** Númericos valores ****");
        Demo.Contato.valLimiteCredito = Demo.Contato.formContext.getAttribute(Demo.Contato.campoLimiteCredito).getValue();
        console.log(Demo.Contato.valLimiteCredito);

        //Texto
        console.log("**** Texto valores ****");
        Demo.Contato.valNome = Demo.Contato.formContext.getAttribute(Demo.Contato.campoNome).getValue();
        console.log(Demo.Contato.valNome);

        //Optionset
        console.log("**** OptionSet valores ****");
        Demo.Contato.valPreferenciaContato = Demo.Contato.formContext.getAttribute(Demo.Contato.campoPreferenciaContato).getValue();
        console.log(Demo.Contato.valPreferenciaContato);
        console.log(Demo.Contato.formContext.getAttribute(Demo.Contato.campoPreferenciaContato).getSelectedOption());

        //Lookup
        console.log("**** LookUp Valores *****");
        Demo.Contato.valConta = Demo.Contato.formContext.getAttribute(Demo.Contato.campoConta).getValue();
        console.log(Demo.Contato.valConta);

    },

    /**
     * DEMONSTRAÇÃO DO CRUD 
     */
    //https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/xrm-webapi/createrecord
    Create: async function (valor) {
        try {
            let idContato = Demo.Contato.formContext.data.entity.getId().replace("{", "").replace("}", "");
            let idConta = Demo.Contato.valConta[0].id.replace("{", "").replace("}", "");
            //Objeto a ser criado
            var Oportunidade = {};
            Oportunidade["parentcontactid@odata.bind"] = "/contacts(" + idContato + ")";
            Oportunidade["parentaccountid@odata.bind"] = "/accounts(" + idConta + ")";
            Oportunidade.name = "Limite Credito - " + Demo.Contato.valConta[0].name + " - Potencial grande";
            Oportunidade.purchasetimeframe = 4; //OptionSet - Desconhecido  
            Oportunidade.estimatedvalue = valor;
            //chamada assincrona para criação
            let result = await Demo.Helper.CriarRegistroAsync("opportunity", Oportunidade);
            Demo.Contato.valIdOportunidade = result.id;
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
            let Oportunidade = await Demo.Helper.RetornarRegistroAsync("opportunity", Demo.Contato.valIdOportunidade, "?$select=name,purchasetimeframe");
            let Texto = "Titulo " + Oportunidade.name;
            Texto += "\nPeríodo de Compra " + Oportunidade["purchasetimeframe@OData.Community.Display.V1.FormattedValue"];
            //@OData.Community.Display.V1.FormattedValue - Em caso de campo option set para recupera o label utilizamos essa conotação, 
            //caso passe só nome da propridade ira retornar o valor numerico
            Demo.Helper.ExibirAlertaSemConfirmacao(Texto, "Oportunidade");

        } catch (error) {
            Demo.Contato.TratarErroWebApi(error);
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
            Demo.Contato.TratarErroWebApi(error);
        }
    },
    //https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/xrm-webapi/updaterecord
    Update: async function () {
        try {
            let Oportunidade = {
                name: "[SUSPENÇÃO DE CREDITO] - Atualizado " + Date.toLocaleString(),
                statuscode: 2, //OptionSet - SUSPENSO
            };
            let result = await Demo.Helper.AtualizarRegistroAsync("opportunity", Demo.Contato.valIdOportunidade, Oportunidade);
            let msg = "Registro " + result.id + " Atualizado!";
            Demo.Helper.ExibirAlertaSemConfirmacao(msg, "Oportunidade Atualizada");
        } catch (error) {
            Demo.Contato.TratarErroWebApi(error);
        }
    },
    //https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/xrm-webapi/deleterecord
    Delete: async function () {
        try {
            await Demo.Helper.DeletarRegistroAsync("opportunity", Demo.Contato.valIdOportunidade);
            Demo.Helper.ExibirAlertaSemConfirmacao("Registro Deletado", "Sucesso");
        } catch (error) {
            Demo.Contato.TratarErroWebApi(error);
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
        let telefone = Demo.Contato.formContext.getAttribute(Demo.Contato.campoTelefoneComercial).getValue();
        let pattern = /^[\+]?[(]?[0-9]{3}[)]?[-\s\.]?[0-9]{3}[-\s\.]?[0-9]{4,6}$/im;

        if (pattern.test(telefone)) {
            Demo.Contato.formContext.ui.clearFormNotification("ValidaTel");
        }
        else {
            Demo.Contato.formContext.ui.setFormNotification("Telefone inválido", "ERROR", "ValidaTel"); //ERROR / WARNING / INFO

        }
    },
    /**
     * Notificações do Tipo Form
     * https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/controls/setnotification
     */
    ValidarFax: function () {
        let email = Demo.Contato.formContext.getAttribute(Demo.Contato.campoEmail).getValue();
        let pattern = /^[\w-]+(\.[\w-]+)*@(([A-Za-z\d][A-Za-z\d-]{0,61}[A-Za-z\d]\.)+[A-Za-z]{2,6}|\[\d{1,3}(\.\d{1,3}){3}\])$/;

        if (pattern.test(email)) {
            Demo.Contato.formContext.getControl(Demo.Contato.campoEmail).clearNotification("ValidaEmail");
        }
        else {
            Demo.Contato.formContext.getControl(Demo.Contato.campoEmail).setNotification("Email inválido", "ValidaEmail");
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
        let limite = Demo.Contato.formContext.getAttribute(Demo.Contato.campoLimiteCredito).getValue();
        if (limite !== null && limite > 5000)
            Demo.Contato.Create(limite);
    },
    TratarErroWebApi: function (e) {
        let retorno = JSON.parse(e.raw);
        Demo.Helper.ExibirAlertaSemConfirmacao(retorno.message, "Erro");
    }
};