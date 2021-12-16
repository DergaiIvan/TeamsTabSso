import React, { Component } from 'react';

import './custom.css'

import * as microsoftTeams from "@microsoft/teams-js";
import { HashRouter as Router, Route } from "react-router-dom";

import AuthService from '../src/services/AuthService'

import Tab from '../src/components/Tab';
import TeamsAuthPopup from '../src/components/TeamsAuthPopup';
import Web from '../src/components/Web';

export default class App extends Component {

  constructor() {
    super();
    this.state = {
      authInitialized: false
    }
  }

  componentDidMount() {
    // React routing and OAuth don't play nice together
    // Take care of the OAuth fun before routing
    AuthService.init().then(() => {
      console.log("auth success");
      this.setState({
        authInitialized: true
      });
    })
  }

  render() {
    if (microsoftTeams) {
      if (!this.state.authInitialized) {
        console.log("need auth");
        // Wait for Auth Service to initialize
        return (<div className="App"><p>Authorizing...</p></div>);
      } else {
        // Set app routings that don't require microsoft Teams
        // SDK functionality.  Show an error if trying to access the
        // Home page.
        if (window.parent === window.self) {
          console.log("web");
          return (
            <div className="App">
              <Router>
                <Route exact path="/" component={Web} />
                <Route exact path="/web" component={Web} />
                <Route exact path="/teamsauthpopup" component={TeamsAuthPopup} />
              </Router>
            </div>
          );
        }

        // Initialize the Microsoft Teams SDK
        microsoftTeams.initialize(window);

        // Display the app home page hosted in Teams
        return (
          <Router>
            <Route exact path="/" component={Tab} />
          </Router>
        );
      }
    }
  }
}
