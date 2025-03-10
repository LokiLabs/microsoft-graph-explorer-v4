import { MessageBarType } from 'office-ui-fabric-react';
import { geLocale } from '../../../appLocale';
import { authenticationWrapper } from '../../../modules/authentication';
import { IAction } from '../../../types/action';
import { IRequestOptions } from '../../../types/request';
import { IRootState } from '../../../types/root';
import { sanitizeQueryUrl } from '../../utils/query-url-sanitization';
import { parseSampleUrl } from '../../utils/sample-url-generation';
import { translateMessage } from '../../utils/translate-messages';
import {
  FETCH_SCOPES_ERROR,
  FETCH_SCOPES_PENDING,
  FETCH_SCOPES_SUCCESS,
} from '../redux-constants';
import {
  PERMS_SCOPE,
  PERMISSION_MODE_TYPE,
  ACCOUNT_TYPE
} from '../graph-constants';
import {
  getAuthTokenSuccess,
  getConsentedScopesSuccess,
} from './auth-action-creators';
import { setQueryResponseStatus } from './query-status-action-creator';

export function fetchScopesSuccess(response: object): IAction {
  return {
    type: FETCH_SCOPES_SUCCESS,
    response,
  };
}

export function fetchScopesPending(): any {
  return {
    type: FETCH_SCOPES_PENDING,
  };
}

export function fetchScopesError(response: object): IAction {
  return {
    type: FETCH_SCOPES_ERROR,
    response,
  };
}

export function fetchScopes(): Function {
  return async (dispatch: Function, getState: Function) => {
    let hasUrl = false; // whether permissions are for a specific url
    try {
      const { devxApi, permissionsPanelOpen, profileType, permissionModeType, sampleQuery: query }: IRootState = getState();
      let permissionsUrl = `${devxApi.baseUrl}/permissions`;
      const permsScopeLookup = {
        [PERMISSION_MODE_TYPE.User]: PERMS_SCOPE.WORK,
        [PERMISSION_MODE_TYPE.TeamsApp]: PERMS_SCOPE.APPLICATION,
      }
      let scope = permsScopeLookup[permissionModeType];

      if (scope !== PERMS_SCOPE.APPLICATION) {
        if (profileType === ACCOUNT_TYPE.AAD) {
          scope = PERMS_SCOPE.WORK;
        } else if (profileType === ACCOUNT_TYPE.MSA) {
          scope = PERMS_SCOPE.PERSONAL;
        }
      }

      if (!permissionsPanelOpen) {
        const signature = sanitizeQueryUrl(query.sampleUrl);
        const { requestUrl, sampleUrl } = parseSampleUrl(signature);

        if (!sampleUrl) {
          throw new Error('url is invalid');
        }
        permissionsUrl = `${permissionsUrl}?requesturl=/${requestUrl}&method=${query.selectedVerb}&scopeType=${scope}`;
        hasUrl = true;
      }

      if (devxApi.parameters) {
        permissionsUrl = `${permissionsUrl}${query ? '&' : '?'}${devxApi.parameters}`;
      }

      if (permissionsPanelOpen) {
        permissionsUrl = `${devxApi.baseUrl}/permissions?scopeType=${scope}`;
      }

      const headers = {
        'Content-Type': 'application/json',
        'Accept-Language': geLocale,
      };

      const options: IRequestOptions = { headers };

      dispatch(fetchScopesPending());

      const response = await fetch(permissionsUrl, options);
      if (response.ok) {
        const scopes = await response.json();
        return dispatch(
          fetchScopesSuccess({
            hasUrl,
            scopes,
          })
        );
      }
      throw response;
    } catch (error) {
      return dispatch(
        fetchScopesError({
          hasUrl,
          error,
        })
      );
    }
  };
}

export function consentToScopes(scopes: string[]): Function {
  return async (dispatch: Function) => {
    try {
      const authResponse = await authenticationWrapper.consentToScopes(scopes);
      if (authResponse && authResponse.accessToken) {
        dispatch(getAuthTokenSuccess(true));
        dispatch(getConsentedScopesSuccess(authResponse.scopes));
      }
    } catch (error) {
      const { errorCode } = error;
      dispatch(
        setQueryResponseStatus({
          statusText: translateMessage('Scope consent failed'),
          status: errorCode,
          ok: false,
          messageType: MessageBarType.error,
        })
      );
    }
  };
}
