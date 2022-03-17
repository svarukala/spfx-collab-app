import * as React from 'react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { useState, useEffect } from 'react';
import { Pivot, PivotItem } from 'office-ui-fabric-react';
import { Configuration, LogLevel, PublicClientApplication, AccountInfo, SilentRequest, 
    InteractionRequiredAuthError, AuthorizationUrlRequest } from "@azure/msal-browser";
import { Providers, SharePointProvider, SimpleProvider, ProviderState } from '@microsoft/mgt-spfx';    
import SPOReusable from './SPOReusable';
import MSGReusable from './MSGReusable';
import MGTReusable from './MGTReusable';
import {  FileList, PeoplePicker, Get, MgtTemplateProps } from '@microsoft/mgt-react/dist/es6/spfx';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

const msalConfig: Configuration = {
    auth: {
      clientId: process.env.SPFX_CLIENTID, //"0d3aa5dd-93b9-40e3-aaf4-73e209f153d3",
      authority: "https://login.microsoftonline.com/"+ process.env.SPFX_TENANTID, //"https://login.microsoftonline.com/044f7a81-1422-4b3d-8f68-3001456e6406",
      redirectUri: process.env.SPFX_REDIRECTURI //"https://m365x229910.sharepoint.com/sites/DevDemo/_layouts/15/workbench.aspx",
    },
    cache: {
      cacheLocation: "localStorage", // This configures where your cache will be stored
      storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
  },
    system: {
      iframeHashTimeout: 10000,
      loggerOptions: {
        loggerCallback: (level, message, containsPii) => {
          if (containsPii) {
            return;
          }
          switch (level) {
            case LogLevel.Error:
              console.error(message);
              return;
            case LogLevel.Info:
              console.info(message);
              return;
            case LogLevel.Verbose:
              console.debug(message);
              return;
            case LogLevel.Warning:
              console.warn(message);
              return;
          }
        },
      },
    },
  };
  
const msalInstance: PublicClientApplication = new PublicClientApplication(
    msalConfig
);

let currentAccount: AccountInfo = null;

const mgtTokenrequest: SilentRequest = {
    scopes: JSON.parse(process.env.SPFX_MGT_SCOPES),
    //process.env.SPFX_MGT_SCOPES.split(","), 
    //['Mail.Read','calendars.read', 'user.read', 'openid', 'profile', 'people.read', 'user.readbasic.all', 'files.read', 'files.read.all'],
    account: currentAccount,
};
const oboTokenrequest = {
    scopes: JSON.parse(process.env.SPFX_OBOBROKER_SCOPES), //.split(","), 
    //["api://3271e1a1-0da7-476b-b573-e360600674a9/access_as_user"],
    account: currentAccount,
};

function AppInit(props) {
    const [ssoToken, setSsoToken] = useState<string>();
    const [loginName, setLoginName] = useState<string>();
    const [error, setError] = useState<string>();
    const [mgtAccessToken, setMgtAccessToken] = useState<string>();
    const [oboAccessToken, setOboAccessToken] = useState<string>();
  
    useEffect(() => {    
        if(!loginName){
          setLoginName(props.spoContext.pageContext.user.loginName);
        }
        if(!ssoToken) { 
            getAccessTokenForOBOBroker();
        }
        if (!Providers.globalProvider) {
            console.log('Initializing global provider');
            Providers.globalProvider = new SimpleProvider(async ()=>{return getAccessTokenForMGT()});  
            //new SharePointProvider(props.spoContext);
            Providers.globalProvider.setState(ProviderState.SignedIn);
        } 
    }, []);    

    const setCurrentAccount = (request): void => {
        const currentAccounts: AccountInfo[] = msalInstance.getAllAccounts();
        if (currentAccounts === null || currentAccounts.length == 0) {
          currentAccount = msalInstance.getAccountByUsername(
            //this.context.pageContext.user.loginName
            loginName
          );
        } else if (currentAccounts.length > 1) {
          console.warn("Multiple accounts detected.");
          currentAccount = msalInstance.getAccountByUsername(
            //this.context.pageContext.user.loginName
            loginName
          );
        } else if (currentAccounts.length === 1) {
          currentAccount = currentAccounts[0];
        }
        request.account = currentAccount;
      }; 
    
      const getAccessTokenForOBOBroker = async (): Promise<string> => {
        console.log("Getting access token async for OBO broker");
        if(oboAccessToken) return oboAccessToken;
        setCurrentAccount(oboTokenrequest);
        console.log(currentAccount);
        return msalInstance
            .acquireTokenSilent(oboTokenrequest)
            .then((tokenResponse) => {
            console.log("Inside Silent");
            console.log("Access token: "+ tokenResponse.accessToken);
            console.log("ID token: "+ tokenResponse.idToken);
            setOboAccessToken(tokenResponse.accessToken);
            return tokenResponse.accessToken;
            })
            .catch((err) => {
            console.log(err);
            console.log("Silent Failed");
            if (err instanceof InteractionRequiredAuthError) {
                return interactionRequired(oboTokenrequest);
            } else {
                console.log("Some other error. Inside SSO.");
                const loginPopupRequest: AuthorizationUrlRequest = oboTokenrequest as AuthorizationUrlRequest;
                loginPopupRequest.loginHint = loginName;
                return msalInstance
                .ssoSilent(loginPopupRequest)
                .then((tokenResponse) => {
                    setOboAccessToken(tokenResponse.accessToken);
                    return tokenResponse.accessToken;
                })
                .catch((ssoerror) => {
                    console.error(ssoerror);
                    console.error("SSO Failed");
                    if (ssoerror) {
                    return interactionRequired(oboTokenrequest);
                    }
                    return null;
                });
            }
            });
    };

    const getAccessTokenForMGT = async (): Promise<string> => {
        console.log("Getting access token async");
        if(mgtAccessToken) return mgtAccessToken;
        setCurrentAccount(mgtTokenrequest);
        console.log(currentAccount);
        return msalInstance
            .acquireTokenSilent(mgtTokenrequest)
            .then((tokenResponse) => {
            console.log("Inside Silent");
            console.log("Access token: "+ tokenResponse.accessToken);
            console.log("ID token: "+ tokenResponse.idToken);
            setMgtAccessToken(tokenResponse.accessToken);
            return tokenResponse.accessToken;
            })
            .catch((err) => {
            console.log(err);
            console.log("Silent Failed");
            if (err instanceof InteractionRequiredAuthError) {
                return interactionRequired(mgtTokenrequest);
            } else {
                console.log("Some other error. Inside SSO.");
                const loginPopupRequest: AuthorizationUrlRequest = mgtTokenrequest as AuthorizationUrlRequest;
                loginPopupRequest.loginHint = loginName;
                return msalInstance
                .ssoSilent(loginPopupRequest)
                .then((tokenResponse) => {
                    setMgtAccessToken(tokenResponse.accessToken);
                    return tokenResponse.accessToken;
                })
                .catch((ssoerror) => {
                    console.error(ssoerror);
                    console.error("SSO Failed");
                    if (ssoerror) {
                    return interactionRequired(mgtTokenrequest);
                    }
                    return null;
                });
            }
            });
    };

      const interactionRequired = (tokenrequest): Promise<string> => {
        console.log("Inside Interaction");
        const loginPopupRequest: AuthorizationUrlRequest = tokenrequest as AuthorizationUrlRequest;
        loginPopupRequest.loginHint = loginName; //??"meganb@m365x229910.onmicrosoft.com"; //this.context.pageContext.user.loginName;
        return msalInstance
          .acquireTokenPopup(loginPopupRequest)
          .then((tokenResponse) => {
            setSsoToken(tokenResponse.idToken);
            if(tokenrequest.scopes.indexOf("access_as_user") > -1) {
              setOboAccessToken(tokenResponse.accessToken);
            }
            else {
                setMgtAccessToken(tokenResponse.accessToken);
            }
            return tokenResponse.accessToken;
          })
          .catch((error) => {
            console.error(error);
            // I haven't implemented redirect but it is fairly easy
            console.error("Maybe it is a popup blocked error. Implement Redirect");
            return null;
          });
      }; 

    return (
        <div>
            {error && "Error: " + error}
            {              
                oboAccessToken &&
          
                <Pivot aria-label="Basic Pivot Example">
                    <PivotItem headerText="SPO REST API">
                        <SPOReusable idToken={oboAccessToken} />
                    </PivotItem>
                    <PivotItem headerText="MS Graph REST API">
                        <MSGReusable idToken={oboAccessToken} />
                    </PivotItem>
                    <PivotItem headerText="MS Graph Toolkit">
                        <Pivot>
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
                    </PivotItem>
                    {/*}
                    <PivotItem headerText="MGT - Fail">
                        <Pivot><MGTReusable /></Pivot>
                    </PivotItem>           
                    */}
                </Pivot>
            }
        </div>
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

export default AppInit;