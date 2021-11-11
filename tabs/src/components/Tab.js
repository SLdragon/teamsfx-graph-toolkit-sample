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

import { Providers, ProvidersChangedState, ProviderState, SimpleProvider } from '@microsoft/mgt-element';
import { PeoplePicker, PersonCard, Person, PersonViewType, FileList } from '@microsoft/mgt-react';
import { TeamsFxProvider } from './TeamsFxProvider';

class Tab extends React.Component {

  constructor(props) {
    super(props);
    this.state = {
      showLoginPage: true,
      selectedPerson: undefined
    }
  }

  async componentDidMount() {
    await this.initGraphToolkit();
  }

  async initGraphToolkit() {
    Providers.globalProvider = new TeamsFxProvider({
      clientId: process.env.REACT_APP_CLIENT_ID,
      initiateLoginEndpoint: process.env.REACT_APP_START_LOGIN_PAGE_URL,
      simpleAuthEndpoint: process.env.REACT_APP_TEAMSFX_ENDPOINT,
      scopes: [
        "User.Read",
        "User.ReadBasic.All",
        "Files.Read"
      ]
    });

    Providers.onProviderUpdated((stateEvent) => {
      if(stateEvent == ProvidersChangedState.ProviderStateChanged) {
        const provider = Providers.globalProvider;
        this.setState({
          showLoginPage: provider && provider.state === ProviderState.SignedOut
        });
      }      
    });
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
                    <h2>My Files</h2>
                  </div>
                </div>
              </div>
              <div className="my-account-area">
                <FileList></FileList>
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
        </div>}
      </div>
    );
  }
}
export default Tab;
