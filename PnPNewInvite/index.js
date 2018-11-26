const sp = require("@pnp/sp").sp;
const SPFetchClient = require("@pnp/nodejs").SPFetchClient;

const KeyVault = require('azure-keyvault');
const msRestAzure = require('ms-rest-azure');

module.exports = async function (context, req) {
    context.log('JavaScript HTTP trigger function processed a request.');
    if ((req.body && req.body.site)) {
    // if ((req.body && req.body.site && req.body.date && req.body.email)) {
        try {
            const siteName = req.body.site;

            const vaultUri = "https://<vaultname>.vault.azure.net/";
            
            // Should always be https://vault.azure.net
            const credentials = await msRestAzure.loginWithAppServiceMSI({resource: 'https://vault.azure.net'});

            const keyVaultClient = new KeyVault.KeyVaultClient(credentials);

            // Get SharePoint key value
            const spVaultSecret = await keyVaultClient.getSecret(vaultUri, "spSecret", "");
            const spSecret = spVaultSecret.value;

            // Setup PnPJs via sp, with spSecret from Key Vault
            sp.setup({
                sp: {
                    fetchClientFactory: () => {
                        return new SPFetchClient(
                            `${process.env.spTenantUrl}/sites/${siteName}/`,
                            process.env.spId,
                            spSecret
                        );
                    },
                },
            });

            // Get all the list in a site
            const lists = await sp.web.lists.get();
            
            context.res = {
                // status: 200, /* Defaults to 200 */
                body: lists
            };
        } catch (error) {
            context.res = {
                status: error.status,
                body: error
            }
        }
    }
    else {
        context.res = {
            status: 400,
            body: "Please pass site in the request body"
        };
    }

};