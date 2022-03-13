// JavaScript source code
if (typeof (Demo) === "undefined") { Demo = {}; }

// "use strict"
Demo.Account = {
    //Definições de contexto
    formContext: {},
    saveContext: {},
    globalContext: Xrm.Utility.getGlobalContext(),

    //definição de campos
    campoCnpj: "ptr_cnpj",
    campoTel: "telephone1",
    campoRua: "address1_line1",
    campoCapital: "revenue",
    campoCidade: "address1_city",
    campoEstado: "address1_stateorprovince",
    campoCep: "address1_postalcode",
    campoNome: "name",

    /**
     * Chama todos as funções no carregamento da página
     * @param {object} executionContext
     */
    OnLoad: function (executionContext) {
        //Chamada de atribuição do formContext 
        //https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/executioncontext/getformcontext
        Demo.Account.formContext = executionContext.getFormContext();
        Demo.Account.OnChange();
    },
    /**
    * Função responsável pela configuração dos eventos onchange dos campos, ela deve sempre ser chamada no carregamento da página
    */
    OnChange: function () {
        Demo.Account.formContext.getAttribute(Demo.Account.campoCnpj).addOnChange(Demo.Account.ConsultarCnpj);
    },
    /**
    * Função reponsável por chamadas no salvamento da página 
    * @param {object} executionContext 
    */
    OnSave: function (executionContext) {
        Demo.Account.saveContext = executionContext.getFormContext();
    },

    ConsultarCnpj: async function () {

        if (Demo.Account.formContext.ui.getFormType() !== 1)
            return;

        let urlFlow = "https://prod-22.brazilsouth.logic.azure.com:443/workflows/e8ca336bd8704ab3afc908d7705b1e78/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=J1TnPCRPw1kwHmTEcZnjTkyYjMtOEQ7mpVo5zqOARw0";
        try {
            let cnpj = Demo.Account.formContext.getAttribute(Demo.Account.campoCnpj).getValue();
            cnpj = cnpj.replace(/\D/g, "");

            if (Demo.Account.ValidarCnpj(cnpj)) {
                Demo.Account.formContext.getControl(Demo.Account.campoCnpj).clearNotification("cnpjErro");
            }
            else {
                Demo.Account.formContext.getControl(Demo.Account.campoCnpj).setNotification("cnpj inválido!", "cnpjErro");
                return;
            }

            let data = { "CNPJ": cnpj };
            Xrm.Utility.showProgressIndicator("Buscando ...");
            let result = await Demo.Helper.POSTAjax(urlFlow, data);
            console.log(result);
            Demo.Account.formContext.getAttribute(Demo.Account.campoNome).setValue(result.nome);
            Demo.Account.formContext.getAttribute(Demo.Account.campoTel).setValue(result.telefone);
            Demo.Account.formContext.getAttribute(Demo.Account.campoCapital).setValue(parseFloat(result.capital_social));
            Demo.Account.formContext.getAttribute(Demo.Account.campoRua).setValue(result.logradouro);
            Demo.Account.formContext.getAttribute(Demo.Account.campoCidade).setValue(result.municipio);
            Demo.Account.formContext.getAttribute(Demo.Account.campoEstado).setValue(result.uf);
            Demo.Account.formContext.getAttribute(Demo.Account.campoCep).setValue(result.cep);
            Xrm.Utility.closeProgressIndicator();
        } catch (error) {
            Demo.Helper.ExibirAlertaSemConfirmacao(error, "Erro");
        }
    },

    ValidarCnpj: function (cnpj) {
        cnpj = cnpj.replace(/\D/g, "");
        let numeros;
        let digitos;
        let soma;
        let i;
        let resultado;
        let pos;
        let tamanho;
        let digitos_iguais;
        digitos_iguais = 1;

        if (cnpj.length < 14 && cnpj.length < 15)
            return false;

        for (i = 0; i < cnpj.length - 1; i++) {
            if (cnpj.charAt(i) != cnpj.charAt(i + 1)) {
                digitos_iguais = 0;
                break;
            }
        }

        if (!digitos_iguais) {
            tamanho = cnpj.length - 2;
            numeros = cnpj.substring(0, tamanho);
            digitos = cnpj.substring(tamanho);
            soma = 0;
            pos = tamanho - 7;

            for (i = tamanho; i >= 1; i--) {
                soma += numeros.charAt(tamanho - i) * pos--;
                if (pos < 2)
                    pos = 9;
            }

            resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
            if (resultado != digitos.charAt(0))
                return false;

            tamanho = tamanho + 1;
            numeros = cnpj.substring(0, tamanho);
            soma = 0;
            pos = tamanho - 7;
            for (i = tamanho; i >= 1; i--) {
                soma += numeros.charAt(tamanho - i) * pos--;
                if (pos < 2)
                    pos = 9;
            }

            resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
            if (resultado != digitos.charAt(1))
                return false;

            return true;
        }
        else {
            return false;
        }
    }
}