import * as React from 'react';
import { useState, useEffect } from 'react';
import { List } from "@fluentui/react-northstar";

function MSGReusable(props) {
  // Declare a new state variable, which we'll call "count"
  const [ssoToken, setSsoToken] = useState<string>();
  const [error, setError] = useState<string>();
  const [resourceOboToken, setResourceOboToken] = useState<string>();
  const [searchResults, setSearchResults] = useState<any[]>();

  useEffect(() => {
    if (props.idToken) {
      setSsoToken(props.idToken);
    }
  }, []);
  
  const getResourceAccessTokenOBO = async () => {
    const requestOptions = {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + ssoToken }
    };
    //fetch(`https://azfun.ngrok.io/api/TeamsOBOHelper?tokenFor=msg`, requestOptions).then(async response =>{
    fetch(process.env.SPFX_OBOBROKER_URL +`?tokenFor=msg`, requestOptions).then(async response =>{
        const responsePayload = await response.json();
        if (response.ok) {
            setResourceOboToken(responsePayload.access_token);
        } else {
          if (responsePayload!.error === "consent_required") {
            setError("consent_required");
          } else {
            setError("unknown SSO error");
          }
        }
    })
    .catch((err) => {
        console.log(err);
        console.log("OBO Failed");
        setError("OBO Failed: " + err);
    });
  };

  useEffect(() => {
    if (ssoToken && ssoToken.length > 0) {
      getResourceAccessTokenOBO();
    }
  }, [ssoToken]);      

  useEffect(() => {
    getSPOSearchResutls();
  }, [resourceOboToken]);

    const getSPOSearchResutls = async () => {
        if (!resourceOboToken) { return; }

        const endpoint = process.env.SPFX_MSG_SEARCHQUERY;
        //`https://graph.microsoft.com/v1.0/sites?search=Contoso`;
        console.log("MSG endpoint: " + endpoint);
        const requestObject = {
        method: 'GET',
        headers: {
            "Authorization": "Bearer " + resourceOboToken,
            "Content-Type": "application/json"
            }
        };
        const response = await fetch(endpoint, requestObject);
        const responsePayload = await response.json();
    
        console.log(responsePayload.value);
        if (response.ok) {
            const resultSet = responsePayload.value.map((result: any) => ({
            key:result.id,
            //header:result.Cells[0].Value,
            headerMedia:result.displayName,
            content:result.webUrl,
            }));
        console.log(JSON.stringify(resultSet));
        setSearchResults(resultSet);
    }
    }

  return (
    <div>
        {error && "Error: " + error}
        {searchResults && <List items={searchResults} />}
        <br/><br/>
        {ssoToken && "ID Token: " + ssoToken}
        <br/><br/>
        {resourceOboToken && "Access Token: " + resourceOboToken}
    </div>
  );
}

export default MSGReusable;


/*
  const getResourceAccessTokenOBO = async () => {
    const response = await fetch(`https://azfun.ngrok.io/api/TeamsOBOHelper?ssoToken=${ssoToken}&tokenFor=msg`);
    const responsePayload = await response.json();
    if (response.ok) {
        setResourceOboToken(responsePayload.access_token);
    } else {
      if (responsePayload!.error === "consent_required") {
        setError("consent_required");
      } else {
        setError("unknown SSO error");
      }
    }
  };*/