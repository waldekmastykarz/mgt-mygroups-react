import { Persona, PersonaSize, Spinner } from '@fluentui/react';
import { MsalProvider, Providers, ProviderState, TemplateHelper } from '@microsoft/mgt';
import { Get, Login, MgtTemplateProps } from '@microsoft/mgt-react';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import { Group } from '@microsoft/microsoft-graph-types';
import React, { useState } from 'react';
import './App.css';

const GroupInfo = (props: MgtTemplateProps) => {
  const [imageUrl, setGroupImage] = useState('');
  const url = window.URL || window.webkitURL;

  const group: Group = props.dataContext;
  const openGroup = (event: any) => window.open(`https://outlook.office365.com/mail/group/${group.mail?.substr(group.mail.indexOf('@') + 1)}/${group.mailNickname}/email`, '_blank');

  if (!imageUrl) {
    Providers.globalProvider.graph.client
      .api(`/groups/${group.id}/photo/$value`)
      .responseType(ResponseType.RAW)
      .get()
      .then((res: Response): Promise<Blob> => {
        if (!res.ok) {
          return Promise.reject(res.statusText);
        }

        return res.blob();
      })
      .then((blob: Blob): void => {
        const blobUrl = url.createObjectURL(blob);
        setGroupImage(blobUrl);
      }, _ => { });
  }

  if (group.groupTypes && group.groupTypes.indexOf('Unified') > -1) {
    return <Persona text={group.displayName as string}
        secondaryText={group.description as string}
        imageUrl={imageUrl}
        size={PersonaSize.size56}
        onClick={openGroup}
        className="group" />;
  }
  else {
    return <div />;
  }
}

const GroupsLoading = (props: MgtTemplateProps) => {
  return <Spinner label='Loading groups...' labelPosition='right' />
}

function App() {
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  Providers.globalProvider = new MsalProvider({
    clientId: '22ea8a07-7069-4d02-a7a0-c47224a1c401'
  });
  Providers.globalProvider.onStateChanged(e => {
    if (Providers.globalProvider.state !== ProviderState.Loading) {
      setIsLoggedIn(Providers.globalProvider.state === ProviderState.SignedIn);
    }
  });
  TemplateHelper.setBindingSyntax('[[', ']]');

  return (
    <div className="App">
      <header>
        <Login />
      </header>
      { isLoggedIn &&
        <Get resource="me/memberOf?$top=500" scopes={["group.read.all"]}>
          <GroupInfo template="value" />
          <GroupsLoading template="loading" />
        </Get>
      }
    </div>
  );
}

export default App;
