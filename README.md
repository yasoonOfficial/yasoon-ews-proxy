# Restful API proxy for Exchange Web Services (EWS)

This repository contains the sources of our EWS proxy. It's necessary to proxy the requests, as Exchange does not support CORS at all, so accessing EWS from the browser is not possible. The proxy is currently used by our following products:

- Office Calendars for Confluence

This proxy also aims to stick as closely as possible to the data format that the Microsoft Graph API uses, to make it easy to support both Graph & EWS at the same time.

## Installation

Todo

## Security

As the proxy call EWS on behalf of each individul user, it obviously needs to handle user credentials. Unfortunately, on-premise Exchange installations only support NTLM & basic auth, both based on plaintext username and password.

We understand security is a critical concern for our customers, so we provide the source of our proxy for everyone to see. We are not storing credentials in any form, they are just used for the sole purpose of accessing EWS. 

As this may be a dealbreaker for certain industries, it it possible to self-host the proxy. We'll be providing a possibility to change the proxy URL in our products that use this proxy. Please [contact](mailto:contact@yasoon.com) us to discuss specifics, we can help you setup the proxy in your environment and enable features like SSL.

## Coming up next

- We'll be enabling our producs (currently only Office Calendars for Confluence) to support a self-hosted proxy. This will be as simple as switching the proxy URL. 
- We will be providing a docker image / VM for quick installation
- We will also provide a detailed guide for manual installation