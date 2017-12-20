import * as app from './express-app';
import { proxy, createServer } from "aws-serverless-express";
import { APIGatewayEvent, ProxyHandler, Context, ProxyCallback } from 'aws-lambda';

const server = createServer(app, null, ['image/jpeg', 'image/jpg']);

let handlerFunc: ProxyHandler = (event: APIGatewayEvent, context: Context, callback?: ProxyCallback) => {
    proxy(server, event, context);
};

export let handler = handlerFunc;