// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import './App.css';
import './Tab.css'
import {
  TeamsUserCredential,
  loadConfiguration,
} from "@microsoft/teamsfx";
import { Button } from "@fluentui/react-northstar"

import { Providers, ProviderState, SimpleProvider } from '@microsoft/mgt-element';
import { PeoplePicker, PersonCard, Person, PersonViewType } from '@microsoft/mgt-react';

class Tab extends React.Component {

  constructor(props) {
    super(props);
    this.state = {
      showLoginPage: undefined,
      selectedPerson: undefined
    }
  }

  async componentDidMount() {
    await this.initTeamsFx();
    await this.initGraphToolkit(this.credential, this.scope);
    await this.checkIsConsentNeeded();
  }

  async initGraphToolkit(credential, scope) {

    async function getAccessToken(scopes) {
      let tokenObj = await credential.getToken(scopes);
      return tokenObj.token;
    }
  
    async function login() {
      try {
        await credential.login(scopes);
      } catch (err) {
        alert("Login failed: " + err);
        return;
      }
      Providers.globalProvider.setState(ProviderState.SignedIn);
    }
  
    async function logout() {}

    Providers.globalProvider = new SimpleProvider(getAccessToken, login, logout);
    Providers.globalProvider.setState(ProviderState.SignedIn);
  }

  async initTeamsFx() {
    loadConfiguration({
      authentication: {
        initiateLoginEndpoint: process.env.REACT_APP_START_LOGIN_PAGE_URL,
        simpleAuthEndpoint: process.env.REACT_APP_TEAMSFX_ENDPOINT,
        clientId: process.env.REACT_APP_CLIENT_ID
      }
    });
    const credential = new TeamsUserCredential();

    this.credential = credential;
    // Only these two permission can be used without admin approval in microsoft tenant
    this.scope = [
      "User.Read",
      "User.ReadBasic.All",
    ];
  }

  async loginBtnClick() {
    try {
      await this.credential.login(this.scope);
      this.setState({
        showLoginPage: false
      });
    } catch (err) {
      alert("Login failed: " + err);
      return;
    }
  }

  async checkIsConsentNeeded() {
    try {
      await this.credential.getToken(this.scope);
    } catch (error) {
      this.setState({
        showLoginPage: true
      });
      return true;
    }
    this.setState({
      showLoginPage: false
    });
    return false;
  }

  render() {
    const handleInputChange = (e) => {
      this.setState({
          selectedPerson: e.target.selectedPeople[0],
      });
    };
    
    return (
      <div>
        {this.state.showLoginPage === false && <div className="flex-container">
          <div className="features-col">
            <div className="features">

            <div>
                <div className="header">
                  <div className="title">
                    <h2>My Account</h2>
                  </div>
                </div>
              </div>
              <div className="my-account-area">
                <Person personQuery="me" view={PersonViewType.threelines}></Person>
              </div>

              <div>
                <div className="header">
                  <div className="title">
                    <h2>Person Card</h2>
                  </div>
                </div>
              </div>
              <div className="person-card-area">
                <PeoplePicker selectionChanged={handleInputChange} selectionMode="single"></PeoplePicker>
                {this.state.selectedPerson &&
                  <div>
                    <br/>
                    <div>Selected Person: <br/>{this.state.selectedPerson.userPrincipalName}</div>
                    <br/>
                  <PersonCard  userId={this.state.selectedPerson.id}></PersonCard>
                  <br/>
                    <br/> 
                  </div>
                }
              </div>
            </div>
          </div>
        </div>}

        {this.state.showLoginPage === true && <div className="auth">
          <h2>Welcome to TeamsFx Integration with Graph Toolkit App!</h2>
          <Button primary onClick={() => this.loginBtnClick()}>Start</Button>
        </div>}
      </div>
    );
  }
}
export default Tab;
