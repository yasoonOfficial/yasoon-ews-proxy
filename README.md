# Restful API proxy for Exchange Web Services (EWS)

This repository contains the sources of our EWS proxy. The proxy is currently used by our following products:

- Office Calendars for Confluence

This proxy also aims to stick as closely as possible to the data format that the Microsoft Graph API uses, to make it easy to support both Graph & EWS at the same time.

## Background

![image](https://user-images.githubusercontent.com/2111803/33890249-f4298e4c-df52-11e7-804e-67c68fbcc762.png)
It's necessary to proxy the SOAP XML requests to the Exchange server, as it does not support [CORS](https://developer.mozilla.org/en-US/docs/Web/HTTP/CORS). Due to this restriction, connecting to EWS from the browser is not possible, so proxying the requests is (currently) the only option.

## Installation

To use the proxy locally, you have two options:
- Run the proxy from source
- Use the provided binaries (soon)

### Running from source

Just clone the repository and execute the following commands (you'll need NodeJS to run the proxy)

`
npm install
`

`
npm run build
`

`
node ./dist/app.js
`

### Running from binaries (soon)

Just download the binary file from the releases page for your OS. It comes bundled with all dependencies, so you can just run it without installing anything else first.

## Commandline Options

You can use the following commandline options to customize some behaviour:

### Use a different port than default (3000)
`
--port=1234
`

### Use a different secret than default (recommended)
`
--secret=somelongsecret
`

### Enable debug logging
`
--verbose
`

## Security

As the proxy calls EWS on behalf of each individul user, it obviously needs to handle user credentials. Unfortunately, on-premise Exchange installations only support NTLM & basic auth.

We understand security is a critical concern for our customers, so we provide the source of our proxy for everyone to see. We are not storing credentials in any form, they are just used for the sole purpose of accessing EWS. 

As this may be a dealbreaker for certain industries, it it possible to self-host the proxy. We provide a possibility to change the proxy URL in our products that use this proxy. Please [contact](mailto:contact@yasoon.com) us to discuss specifics, we can help you setup the proxy in your environment and enable features like SSL.

You can also decide to test everything using our proxy and decide to run the proxy on your own for production usage.

## Coming up next

- We will be providing a docker image / VM for quick installation
- We will also provide a detailed guide for manual installation
