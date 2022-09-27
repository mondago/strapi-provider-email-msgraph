# strapi-provider-email-msgraph

Microsoft Graph email provider plugin for Strapi 3.x.

## Prerequisites

An app registration for the tenant with Mail.Send permission is required. You'll need:

- Tenant ID
- Client App ID
- Client App Secret

## Installation

Install the package with either npm or yarn.

`npm install mondago/strapi-provider-email-msgraph`

`yarn add mondago/strapi-provider-email-msgraph`

## Configuration

To use this provider setup your config/plugins.js file:

```javascript
module.exports = ({ env }) => ({
  email: {
    provider: "msgraph",
    providerOptions: {
      clientId: env("GRAPH_MAIL_CLIENT_ID"),
      clientSecret: env("GRAPH_MAIL_CLIENT_SECRET"),
      tenantId: env("GRAPH_MAIL_TENANT_ID"),
    },
    settings: {
      defaultFrom: "hello@example.com",
    },
  },
});
```
