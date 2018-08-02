import cloudform, { ApiGateway, StringParameter } from "cloudform";
import { Integration } from "cloudform/types/apiGateway/method";

export default cloudform({
    Description: 'My template',
    Parameters: {
        DeployEnv: new StringParameter({
            Description: 'Deploy environment name',
            AllowedValues: ['stage', 'production']
        })
    },
    Mappings: {
        DeploymentConfig: {
            stage: {
                InstanceType: 't2.small'
            },
            production: {
                InstanceType: 't2.large'
            }
        }
    },
    Resources: {
        ApiGateway: new ApiGateway.RestApi({
            Name: "ews"
        }),
        RootResource: new ApiGateway.Resource({
            RestApiId: "ApiGateway",
            ParentId: "ApiGateway",
            PathPart: "/"
        }),
        V2VersionResource: new ApiGateway.Resource({
            RestApiId: "ApiGateway",
            ParentId: "RootResource",
            PathPart: "/v2"
        }),
        ProxyResource: new ApiGateway.Resource({
            RestApiId: "ApiGateway",
            ParentId: "V2VersionResource",
            PathPart: "/{proxy+}"
        }),
        ProxyMethodAny: new ApiGateway.Method({
            RestApiId: "ApiGateway",
            ResourceId: "ProxyResource",
            HttpMethod: "ANY",
            Integration: new Integration({
                Type: "AWS_PROXY",
                IntegrationHttpMethod: "POST",
                Uri: ""
            })
        })
    }
});