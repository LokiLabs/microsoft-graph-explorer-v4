import { AuthenticationProvider } from '@microsoft/microsoft-graph-client';

import { authenticationWrapper } from '../../../modules/authentication';
import { PERMISSION_MODE_TYPE } from '../graph-constants';

export class GraphApplicationAuthenticationProvider implements AuthenticationProvider {

    /**
     * getAccessToken
     */
    public async getAccessToken(): Promise<string> {
        try {

            const authResult = await authenticationWrapper.getToken(PERMISSION_MODE_TYPE.TeamsApp);
            return authResult.accessToken;
        } catch (error) {
            throw error;
        }
    }
}
