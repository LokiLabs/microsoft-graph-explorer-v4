import { IQuery } from '../../../types/query-runner';
import { PERMISSION_MODE_TYPE } from '../graph-constants';
import { PROFILE_IMAGE_REQUEST_SUCCESS, PROFILE_REQUEST_ERROR, PROFILE_REQUEST_SUCCESS } from '../redux-constants';
import { authenticatedRequest, isImageResponse, parseResponse } from './query-action-creator-util';

export function profileRequestSuccess(response: object): any {
  return {
    type: PROFILE_REQUEST_SUCCESS,
    response,
  };
}

export function profileImageRequestSuccess(response: object): any {
  return {
    type: PROFILE_IMAGE_REQUEST_SUCCESS,
    response,
  };
}

export function profileRequestError(response: object): any {
  return {
    type: PROFILE_REQUEST_ERROR,
    response,
  };
}

export function getProfileInfo(query: IQuery): Function {
  return (dispatch: Function) => {
    const respHeaders: any = {};

    if (!query.sampleHeaders) {
      query.sampleHeaders = [];
    }

    query.sampleHeaders.push({
      name: 'Cache-Control',
      value: 'no-cache'
    });

    // TODO: The permission type of the call needs to be passed down from the function
    // Can temporarily be set as User since its only being used in the context of a user
    return authenticatedRequest(dispatch, query, PERMISSION_MODE_TYPE.User).then(async (response: Response) => {

      if (response && response.ok) {
        const json = await parseResponse(response, respHeaders);
        const contentType = respHeaders['content-type'];
        const isImageResult = isImageResponse(contentType);

        if (isImageResult) {
          return dispatch(
            profileImageRequestSuccess(json),
          );
        } else {
          return dispatch(
            profileRequestSuccess(json),
          );
        }
      }
      return dispatch(profileRequestError({ response }));
    });

  };
}