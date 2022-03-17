import * as React from "react";
import { useState, useEffect } from 'react';
import { Pivot, PivotItem } from 'office-ui-fabric-react';
import { Providers, ProviderState } from '@microsoft/mgt-element';
import { PersonCardInteraction, PersonViewType, ViewType } from '@microsoft/mgt-spfx';
import { Login, PeoplePicker, FileList, Get, MgtTemplateProps } from '@microsoft/mgt-react/dist/es6/spfx';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';


function useIsSignedIn(): [boolean] {
  const [isSignedIn, setIsSignedIn] = useState(true);
  const provider = Providers.globalProvider;
  
  useEffect(() => {
    const updateState = () => {
      const provider = Providers.globalProvider;
      setIsSignedIn(provider && provider.state === ProviderState.SignedIn);
    };

    Providers.onProviderUpdated(updateState);
    updateState();

    return () => {
      Providers.removeProviderUpdatedListener(updateState);
    }
  }, []);

  return [isSignedIn];
}

function MGTReusable() {
  const [isSignedIn] = useIsSignedIn();

  return (
    <Pivot aria-label="Basic Pivot Example">
        <PivotItem headerText="Login">
            <Login />
            </PivotItem>
          <PivotItem headerText="Files">
            <FileList></FileList> 
          </PivotItem>
          <PivotItem headerText="People">
            <br/>
            <PeoplePicker></PeoplePicker>
            </PivotItem>
            <PivotItem headerText="File Upload">
            <FileList driveId="b!mKw3q1anF0C5DyDiqHKMr8iJr_oIRjlGl4854HhHtho07AdbOeaLT5rMH83yt89B" 
          itemPath="/" enableFileUpload></FileList>
              </PivotItem>
          
          <PivotItem headerText="Sites Search Using MSGraph">
              <Get resource="/sites?search=contoso" scopes={['Sites.Read.All']} maxPages={2}>
                      <SiteResult template="value" />
              </Get>
          </PivotItem>
    </Pivot>
      
  );
}

const SiteResult = (props: MgtTemplateProps) => {
    const site = props.dataContext as MicrosoftGraph.Site;

    return (
        <div>
            <h1>{site.name}</h1>
            {site.webUrl}
      </div>
      );
    };

export default MGTReusable;
