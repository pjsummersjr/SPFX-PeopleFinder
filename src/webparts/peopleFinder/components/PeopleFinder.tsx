import * as React from 'react';
import styles from './PeopleFinder.module.scss';
import { IPeopleFinderProps } from './IPeopleFinderProps';
import { escape } from '@microsoft/sp-lodash-subset';

import Profile from "../models/Profile";
import PeopleCard from "./PeopleCard";

export interface IPeopleFinderState {
  profiles: Profile[];
}

export default class PeopleFinder extends React.Component<IPeopleFinderProps, IPeopleFinderState> {

  constructor(props) {
    super(props);
    this.state = { 
      profiles: []
    };
  }

  componentDidUpdate(prevProps: any, newProps: any, snapShot: any): void {

  }

  componentWillMount(): void {
    this.props.profileService.getProfiles().then((newProfiles: Profile[]) => {
      this.setState({profiles: newProfiles});
    });
  }

  public render(): React.ReactElement<IPeopleFinderProps> {
    return (
      //This is the root container - no styling at this point
      <div> 
        {this.state.profiles.map((profile, index) => {
          return (              
              <PeopleCard profileData={profile} />
          );
        })}
      </div>
    );
  }
}
