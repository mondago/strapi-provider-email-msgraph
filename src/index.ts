"use strict";

/**
 * Module dependencies
 */

/* eslint-disable import/no-unresolved */
/* eslint-disable prefer-template */
// Public node modules.
import "isomorphic-fetch"
import { Client } from "@microsoft/microsoft-graph-client"
import { ClientSecretCredential } from "@azure/identity"
import {
	TokenCredentialAuthenticationProvider,
} from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials"

interface Settings {
	defaultFrom: string;
	defaultReplyTo: string;
}

interface SendOptions {
	from?: string;
	to: string;
	cc: string;
	bcc: string;
	replyTo?: string;
	subject: string;
	text: string;
	html: string;
	[key: string]: unknown;
}

interface ProviderOptions {
	tenantId: string;
	clientId: string;
	clientSecret: string;
}

export = {
	provider: "msgraph",
	name: "Microsoft Graph Email Plugin",

	init(providerOptions: ProviderOptions, settings: Settings) {
		const authProvider = new TokenCredentialAuthenticationProvider(
			new ClientSecretCredential(
				providerOptions.tenantId,
				providerOptions.clientId,
				providerOptions.clientSecret
			),
			{ scopes: ["https://graph.microsoft.com/.default"] }
		);

		return {
			send: async (options: SendOptions) => {
				const getEmailFromAddress = () => {
					if (!options.from) {
						return settings.defaultFrom;
					}

					const regex = /[^< ]+(?=>)/g;
					const matches = options.from.match(regex);
					return matches?.length ? matches[0] : settings.defaultFrom;
				};

				const client = Client.initWithMiddleware({
					debugLogging: false,
					authProvider: authProvider,
				});

				const from = getEmailFromAddress();
				const mail = {
					subject: options.subject,
					from: {
						emailAddress: { address: from },
					},
					toRecipients: [
						{
							emailAddress: {
								address: options.to,
							},
						},
					],
					body: options.html
						? {
							content: options.html,
							contentType: "html",
						}
						: {
							content: options.text,
							contentType: "text",
						},
				};

				await client
					.api(`/users/${from}/sendMail`)
					.post({ message: mail })
			},
		};
	},
};
