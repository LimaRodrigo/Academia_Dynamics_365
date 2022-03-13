if (typeof (Demo) === "undefined") { Demo = {}; }

Demo.Helper = {


    /**
     * Cria um registro com possiblidade de relacionar registro e/ou criar registros relcionados, na chamada é necessário chamar dentro de um try-catch e utilizar await na frente para aguardar o resultado. 
     * Mais detalhes no Link: 
     * https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/xrm-webapi/createrecord
     * @param {string} nomeEntidade - Nome lógico da entidade
     * @param {object} objeto - Objeto a ser criado
     * @returns {object}
     */
    CriarRegistroAsync: function (nomeEntidade, objeto) {
        return new Promise((resolve, reject) => {
            Xrm.WebApi.createRecord(nomeEntidade, objeto).then(
                (result) => { 
                    resolve(result); 
                },
                (e) => { reject(e); }
            );
        });
    },
     /**
     * Atualiza um Resgitro, na chamada é necessário chamar dentro de um try-catch e utilizar await na frente para aguardar o resultado. 
     * Mais detalhes no Link: 
     * https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/xrm-webapi/updaterecord
     * @param {string} nomeEntidade - Nome lógico da entidade
     * @param {string} id Id do registro a ser atualizado ex.: CFF5AA7F-B426-EB11-BBF3-000D3A887C31
     * @param {object} objeto - Objeto a ser atualizado
     * @returns {object}
     */
    AtualizarRegistroAsync: function (nomeEntidade, id, objeto) {
        return new Promise((resolve, reject) => {
            Xrm.WebApi.updateRecord(nomeEntidade, id, objeto).then(
                (result) => { resolve(result); },
                (e) => { reject(e); }
            );
        });
    },
     /**
     * Deleta um registro
     * @param {string} nomeEntidade - Nome lógico da entidade
     * @param {string} id Id do registro a ser atualizado ex.: CFF5AA7F-B426-EB11-BBF3-000D3A887C31
     */
    DeletarRegistroAsync: function (nomeEntidade, id) {
        return new Promise((resolve, reject) => {
            Xrm.WebApi.deleteRecord(nomeEntidade, id).then(
                (result) => { resolve(result); },
                (e) => { reject(e); }
            );
        });
    },
     /**
     * Retorna apenas o registros de acordo com o Id, na chamada é necessário chamar dentro de um try-catch e utilizar await na frente para aguardar o resultado. 
     * Mais detalhes no Link: 
     * https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/xrm-webapi/retrieverecord
     * @param {string} nomeEntidade - Nome lógico da entidade
     * @param {string} id - Id do registro a ser recuperado ex.: CFF5AA7F-B426-EB11-BBF3-000D3A887C31
     * @param {string} optionOdata -  Passa estruta odata select e expand(opcional) ex.:  "?$select=name&$expand=primarycontactid($select=contactid,fullname)"
     * @returns {object}
     */
    RetornarRegistroAsync: function (nomeEntidade, id, optionOdata) {
        return new Promise((resolve, reject) => {
            Xrm.WebApi.retrieveRecord(nomeEntidade, id, optionOdata).then(
                (result) => { resolve(result); },
                (e) => { reject(e); }
            );
        });
    },
    /**
     * Retorna um array de objetos de acordo com a busca, na chamada é necessário chamar dentro de um try-catch e utilizar await na frente para aguardar o resultado. 
     * Mais detalhes no Link: 
     * https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/xrm-webapi/retrievemultiplerecords
     * @param {string} nomeEntidade - Nome Lógico da Entidade 
     * @param {string} optionOdata - Passa estruta odata completa com select, filter, top, etc ex.:?$select=name&$filter=campo eq 'teste' &$top=2
     * @returns {object}
     */
    RetornarRegistrosAsync: function (nomeEntidade, optionOdata) {
        return new Promise((resolve, reject) => {
            Xrm.WebApi.retrieveMultipleRecords(nomeEntidade, optionOdata).then(
                (result) => { resolve(result.entities); },
                (e) => { reject(e); }
            );
        });
    },
    /**
     * Exibe um alerta modal na tela
     * @param {string} mensagem - Mensagem a ser exibida
     * @param {string} titulo - titulo do modal
     * @param {number} largura - tamanho lagura medido em pixels, caso nulo padrão 350
     * @param {number} altura - tamanho altura medido em pixels,  caso nulo padrão 550
     */
    ExibirAlertaSemConfirmacao: function (mensagem, titulo = null, largura = null, altura = null) {
        let alertStrings = {
            title: titulo == null ? "" : titulo,
            confirmButtonLabel: "OK",
            text: mensagem
        };
        let alertOptions = {
            height: altura == null ? 350 : altura,
            width: largura == null ? 550 : largura
        };
        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
            function success(result) {
                console.log("Sucesso! Alerta exibido!");
            },
            function (error) {
                concole.log(error.message);
            }
        );
    },


    GetApi: async function (url) {
        return new Promise((resolve, reject) => {
            $.ajax({
                type: "GET",
                async: true,
                url: url,
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                beforeSend: function (XMLHttpRequest) {
                    XMLHttpRequest.setRequestHeader("Accept", "application/json");
                    XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=*");
                },                
                success: function (data) {
                    resolve(data);
                },
                error: function (result) {
                    if (result.responseJSON != undefined && result.responseJSON != null) {
                        console.log(result.responseJSON);
                        reject(result.responseJSON.message);
                    }
                    else if (result.responseText != null) {
                        reject(result.responseText)
                        console.log(result.responseText)
                    };
                }
            });
        });
    },
    POSTAjax: function (url, data) {
        return new Promise((resolve, reject) => {
            $.ajax({
                type: "POST",
                async: true,
                url: url,
                data: JSON.stringify(data),
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                beforeSend: function (XMLHttpRequest) {
                    XMLHttpRequest.setRequestHeader("Accept", "application/json");
                },
                success: function (data) {
                    resolve(data);
                },
                error: function (result) {
                    if (result.responseJSON != undefined && result.responseJSON != null) {
                        console.log(result.responseJSON);
                        reject(result.responseJSON.message);
                    }
                    else if (result.responseText != null) {
                        reject(result.responseText)
                        console.log(result.responseText)
                    };
                }
            });
        });
    }
};