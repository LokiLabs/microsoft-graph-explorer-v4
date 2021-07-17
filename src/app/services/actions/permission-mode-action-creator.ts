import { IAction } from '../../../types/action';
import { PERMISSION_MODE_TYPE } from '../graph-constants';
import {
    APP_POP_UP_SUCCESS,
    CHANGE_PERMISSIONS_MODE_SUCCESS
} from '../redux-constants';

export function changeMode(newPerm: PERMISSION_MODE_TYPE): IAction {
    return {
        type: CHANGE_PERMISSIONS_MODE_SUCCESS,
        response: newPerm,
    };
}

export function setPopUp(response: boolean): IAction {
    return {
        type: APP_POP_UP_SUCCESS,
        response,
    };
}