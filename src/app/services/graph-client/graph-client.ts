import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import { authProvider, appAuthProvider } from './';
import { PERMISSION_MODE_TYPE } from '../graph-constants';

export class GraphClient {
  private static client: Client;

  private static createClient(permissionModeType: PERMISSION_MODE_TYPE): Client {

    const clientOptions = permissionModeType === PERMISSION_MODE_TYPE.User ? {
      authProvider,
    } : { appAuthProvider };

    return Client.initWithMiddleware(clientOptions);
  }

  public static getInstance(permissionModeType: PERMISSION_MODE_TYPE): Client {
    if (!GraphClient.client) {
      GraphClient.client = this.createClient(permissionModeType);
    }
    return GraphClient.client;
  }
}

