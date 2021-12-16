import React, { Component } from 'react';
import { Collapse, Container, Navbar, NavbarBrand, NavbarToggler, NavItem, NavLink } from 'reactstrap';
import { Link } from 'react-router-dom';
import './NavMenu.css';

import * as microsoftTeams from "@microsoft/teams-js";
import { HashRouter as Router, Route } from "react-router-dom";

import AuthService from '../services/AuthService'

import Tab from "./Tab";
import TeamsAuthPopup from './TeamsAuthPopup';
import Web from './Web';

export class NavMenu extends Component {
  static displayName = NavMenu.name;

  constructor (props) {
    super(props);

    this.toggleNavbar = this.toggleNavbar.bind(this);
    this.state = {
      collapsed: true,
      authInitialized: false
    };
  }

  componentDidMount() {
    // React routing and OAuth don't play nice together
    // Take care of the OAuth fun before routing
    AuthService.init().then(() => {
      this.setState({
        authInitialized: true
      });
      console.log("success!")
    })
  }

  toggleNavbar () {
    this.setState({
      collapsed: !this.state.collapsed
    });
  }

  render () {
    // return (
    //   <header>
    //     <Navbar className="navbar-expand-sm navbar-toggleable-sm ng-white border-bottom box-shadow mb-3" light>
    //       <Container>
    //         <NavbarBrand tag={Link} to="/">TeamsTabSso</NavbarBrand>
    //         <NavbarToggler onClick={this.toggleNavbar} className="mr-2" />
    //         <Collapse className="d-sm-inline-flex flex-sm-row-reverse" isOpen={!this.state.collapsed} navbar>
    //           <ul className="navbar-nav flex-grow">
    //             <NavItem>
    //               <NavLink tag={Link} className="text-dark" to="/">Home</NavLink>
    //             </NavItem>
    //             <NavItem>
    //               <NavLink tag={Link} className="text-dark" to="/counter">Counter</NavLink>
    //             </NavItem>
    //             <NavItem>
    //               <NavLink tag={Link} className="text-dark" to="/fetch-data">Fetch data</NavLink>
    //             </NavItem>
    //           </ul>
    //         </Collapse>
    //       </Container>
    //     </Navbar>
    //   </header>
    // );

    if (microsoftTeams) {
      if (!this.state.authInitialized) {
        // Wait for Auth Service to initialize
        return (<div className="App"><p>Authorizing...</p></div>);
      } else {
        // Set app routings that don't require microsoft Teams
        // SDK functionality.  Show an error if trying to access the
        // Home page.
        if (window.parent === window.self) {

          return (
            <div className="App">
              <Router>
                <Route exact path="/tab" component={Web} />
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
            <Route exact path="/tab" component={Tab} />
          </Router>
        );
      }
    }
  }
}
