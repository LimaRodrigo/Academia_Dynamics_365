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

    //Valores Globais
    valLimiteCredito: null,
    valConta: null,
    valPreferenciaContato: null,
    valNome: null,
    valAniversario: null,
    valSuspensaoCredito: null,
    valOportunidade: {},

    OnLoad: function (formContext) {
        /**
         * Chamada de atibição do formContext 
         * https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/executioncontext/getformcontext
         */
        Demo.Contato.formContext = formContext.getFormContext();
        Demo.Contato.OnChange();
        Demo.Contato.RecuperarValoresTipagem();
    },

    OnChange: function () {
        //Adiciona evento onchange em um campo
        //https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/attributes/addonchange
        Demo.Contato.formContext.getAttribute(Demo.Contato.campoLimiteCredito).addOnChange(Demo.Contato.ValidarCriacao);
    },

    OnSave: function (saveContext) {
        Demo.Contato.saveContext = saveContext.getFormContext();
    },



    /**
     * Funções de ações em Tela
     */

    /**
     * Domenstração de Tipagem do Dynamics
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

    ValidarCriacao: function () {
        let limite = Demo.Contato.formContext.getAttribute(Demo.Contato.campoLimiteCredito).getValue();
        if (limite !== null && limite > 5000)
            Demo.Contato.Create();
    },

    /**
     * Demostração do CRUD 
     */

    Create: async function () {
        try {
            let idContato = Demo.Contato.formContext.data.entity.getId().replace("{", "").replace("}", "");
            let idConta = Demo.Contato.valConta[0].id.replace("{", "").replace("}", "");
            //Objeto a ser criado
            var Oportunidade = {};
            Oportunidade["parentcontactid@odata.bind"] = "/contacts(" + idContato + ")";
            Oportunidade["parentaccountid@odata.bind"] = "/accounts(" + idConta + ")";
            Oportunidade.name = "Limite Credito -" + Demo.Contato.valConta[0].name + " - Potencial grande";
            Oportunidade.purchasetimeframe = 3;
            Demo.Contato.valOportunidade = Oportunidade;
            console.log(Oportunidade);
            //chamada assincrona para criação
            let result = await Demo.Helper.CriarRegistroAsync("opportunity", Oportunidade);
            let msg = "Registro " + result.id + " Criado!";
            Demo.Helper.ExibirAlertaSemConfirmacao(msg, "Sucesso");
            console.log(result);

        } catch (error) {
            let retorno = JSON.parse(error.raw);
            Demo.Helper.ExibirAlertaSemConfirmacao(retorno.message, "Erro");
        }
    },
    Read: async function () {

    },
    Update: async function () {

    },
    Delete: async function () {

    },
};