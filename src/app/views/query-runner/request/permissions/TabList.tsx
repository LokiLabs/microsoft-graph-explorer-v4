import { DetailsList, DetailsListLayoutMode, IColumn, Label, Link, SelectionMode } from 'office-ui-fabric-react';
import React from 'react';
import { FormattedMessage } from 'react-intl';
import { useDispatch, useSelector } from 'react-redux';
import { IPermission } from '../../../../../types/permissions';
import { IRootState } from '../../../../../types/root';
import { togglePermissionsPanel } from '../../../../services/actions/permissions-panel-action-creator';
import { DISPLAY_APPLICATION_PERMISSIONS, RSC_PERMISSIONS_ENDINGS } from '../../../../services/graph-constants';
import { setConsentedStatus } from './util';

interface ITabList {
  columns: any[];
  classes: any;
  renderItemColumn: Function;
  renderDetailsHeader: Function;
  maxHeight: string;
}

const TabList = ({ columns, classes, renderItemColumn, renderDetailsHeader, maxHeight }: ITabList) => {
  const dispatch = useDispatch();
  const { consentedScopes, scopes, authToken, permissionModeType } = useSelector((state: IRootState) => state);
  const permissions: IPermission[] = scopes.hasUrl ? scopes.data : [];
  const tokenPresent = !!authToken.token;

  const isRSC = (permission: IPermission) => {
    return RSC_PERMISSIONS_ENDINGS.some((ending) => permission.value.indexOf(ending) !== -1)
  }

  const filterPermissionsForRSC = () => {
    return permissionModeType === DISPLAY_APPLICATION_PERMISSIONS ? permissions.filter(isRSC) : permissions;
  }

  const filteredPermissions = filterPermissionsForRSC();

  setConsentedStatus(tokenPresent, filteredPermissions, consentedScopes);

  const openPermissionsPanel = () => {
    dispatch(togglePermissionsPanel(true));
  }

  const displayNoPermissionsFoundMessage = () => {
    return (<Label className={classes.permissionLabel}>
      <FormattedMessage id='permissions not found in permissions tab' />
      <Link onClick={openPermissionsPanel}>
        <FormattedMessage id='open permissions panel' />
      </Link>
      <FormattedMessage id='permissions list' />
    </Label>);
  }

  const displayNotSignedInMessage = () => {
    return (<Label className={classes.permissionLabel}>
      <FormattedMessage id='sign in to view a list of all permissions' />
    </Label>)
  }

  if (tokenPresent && !scopes.hasUrl) {
    return displayNoPermissionsFoundMessage();
  }

  if (!tokenPresent && !scopes.hasUrl) {
    return displayNotSignedInMessage();
  }

  if (filteredPermissions.length === 0) {
    return displayNoPermissionsFoundMessage();
  }

  if (filteredPermissions.length === 1 && filteredPermissions[0].value === "Not supported.") {
    return displayNoPermissionsFoundMessage();
  }

  return (
    <>
      <Label className={classes.permissionLength}>
        <FormattedMessage id='Permissions' />&nbsp;({filteredPermissions.length})
      </Label>
      <Label className={classes.permissionText}>
        {!tokenPresent && <FormattedMessage id={permissionModeType ? 'sign in to consent to permissions' : 'sign in to consent to application permissions'} />}
        {tokenPresent && <FormattedMessage id={permissionModeType ? 'permissions required to run the query' : 'application permissions required to run the query'} />}
      </Label>
      <DetailsList styles={{ root: { maxHeight } }}
        items={filteredPermissions}
        columns={columns}
        onRenderItemColumn={(item?: any, index?: number, column?: IColumn) => renderItemColumn(item, index, column)}
        selectionMode={SelectionMode.none}
        layoutMode={DetailsListLayoutMode.justified}
        onRenderDetailsHeader={(props?: any, defaultRender?: any) => renderDetailsHeader(props, defaultRender)} />
      {filteredPermissions && filteredPermissions.length === 0 &&
        displayNoPermissionsFoundMessage()
      }
    </>
  );
};

export default TabList;




