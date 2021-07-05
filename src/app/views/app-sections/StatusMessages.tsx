import { Link, MessageBar } from 'office-ui-fabric-react';
import React, { Fragment } from 'react';
import { FormattedMessage } from 'react-intl';
import { IPermission } from '../../../types/permissions';

import { IQuery } from '../../../types/query-runner';
import { PERMISSION_MODE_TYPE, GRAPH_URL, RSC_PERMISSIONS_ENDINGS, SWITCH_MODE_NOTICE } from '../../services/graph-constants';
import {
  convertArrayToObject, extractUrl, getMatchesAndParts,
  matchIncludesLink, replaceLinks
} from '../../utils/status-message';

export function statusMessages(queryState: any, sampleQuery: IQuery, actions: any, scopes: any, permissionModeType: PERMISSION_MODE_TYPE) {
  function displayStatusMessage(message: string, urls: any) {
    const { matches, parts } = getMatchesAndParts(message);

    if (!parts || !matches || !urls || Object.keys(urls).length === 0) {
      return message;
    }

    return parts.map((part: string, index: number) => {
      const includesLink = matchIncludesLink(matches, part);
      const displayLink = (): React.ReactNode => {
        const link = urls[part];
        if (link) {
          if (link.includes(GRAPH_URL)) {
            return <Link onClick={() => setQuery(link)}>{link}</Link>;
          }
          return <Link target="_blank" href={link}>{link}</Link>;
        }
      };
      return (
        <Fragment key={part + index}>{includesLink ?
          displayLink() : part}
        </Fragment>
      );
    })
  }

  function setQuery(link: string) {
    const query: IQuery = { ...sampleQuery };
    query.sampleUrl = link;
    actions.setSampleQuery(query);
  };

  if (queryState) {
    const { messageType, statusText, status, duration } = queryState;
    let urls: any = {};
    let message = status;
    const extractedUrls = extractUrl(status);
    if (extractedUrls) {
      message = replaceLinks(status);
      urls = convertArrayToObject(extractedUrls);
    }
    const isRSC = (permission: IPermission) => {
      return RSC_PERMISSIONS_ENDINGS.some((ending) => permission.value.indexOf(ending) !== -1)
    }

    const filterPermissionsForRSC = () => {
      return permissionModeType === PERMISSION_MODE_TYPE.TeamsApp ? permissions.filter(isRSC) : permissions;
    }

    const permissions: IPermission[] = scopes.hasUrl ? scopes.data : [];
    const filteredScopes = filterPermissionsForRSC();
    const noPerms = filteredScopes.length === 0;

    return (
      <MessageBar messageBarType={noPerms ? SWITCH_MODE_NOTICE : messageType}
        isMultiline={true}
        onDismiss={actions.clearQueryStatus}
        dismissButtonAriaLabel='Close'
        aria-live={'assertive'}>

        {noPerms && permissionModeType === PERMISSION_MODE_TYPE.TeamsApp && <FormattedMessage id='delegated user mode' />}
        {noPerms && permissionModeType === PERMISSION_MODE_TYPE.User && <FormattedMessage id='application mode' />}

        {!noPerms && `${statusText} - `}{!noPerms && displayStatusMessage(message, urls)}

        {!noPerms && duration && <>
          {` - ${duration}`}<FormattedMessage id='milliseconds' />
        </>}

        {status === 403 && <>.
          <FormattedMessage id='consent to scopes' />
          <span style={{ fontWeight: 600 }}>
            <FormattedMessage id='modify permissions' />
          </span>
          <FormattedMessage id='tab' />
        </>}

      </MessageBar>);
  }
}
