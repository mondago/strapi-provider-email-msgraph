"use strict";

/**
 * Module dependencies
 */

/* eslint-disable import/no-unresolved */
/* eslint-disable prefer-template */
// Public node modules.
require("isomorphic-fetch");
const { Client } = require("@microsoft/microsoft-graph-client");
const { ClientSecretCredential } = require("@azure/identity");
const {
  TokenCredentialAuthenticationProvider,
} = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");

module.exports = {
  provider: "msgraph",
  name: "Microsoft Graph Email Plugin",

  init: (providerOptions = {}, settings = {}) => {
    const authProvider = new TokenCredentialAuthenticationProvider(
      new ClientSecretCredential(
        providerOptions.tenantId,
        providerOptions.clientId,
        providerOptions.clientSecret
      ),
      { scopes: ["https://graph.microsoft.com/.default"] }
    );

    return {
      send: (options) => {
        return new Promise((resolve, reject) => {
          const client = Client.initWithMiddleware({
            debugLogging: false,
            authProvider: authProvider,
          });

          const mail = {
            subject: options.subject,
            from: {
              emailAddress: { address: options.from || settings.defaultFrom },
            },
            toRecipients: [
              {
                emailAddress: {
                  address: options.to,
                },
              },
            ],
            body: options.text
              ? {
                  content: options.text,
                  contentType: "text",
                }
              : {
                  content: options.html,
                  contentType: "html",
                },
          };

          client
            .api(`/users/${options.from || settings.defaultFrom}/sendMail`)
            .post({ message: mail })
            .then(resolve)
            .catch(reject);
        });
      },
    };
  },
};
