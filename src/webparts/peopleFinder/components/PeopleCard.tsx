import * as React from 'react';

import Profile from '../models/Profile';

export interface IPeopleCardProps {
    profileData: Profile;
}

export interface IPeopleCardState {
    profileData: Profile;
}

export default class PeopleCard extends React.Component<IPeopleCardProps, IPeopleCardState> {

    componentWillMount() {
        this.setState({profileData: this.props.profileData});
    }

    componentDidUpdate(prevProps: any, newProps: any, snapShot: any): void {
        if(prevProps !== newProps) {
            this.setState({profileData: this.props.profileData});
        }
    }

    public render(): React.ReactElement<IPeopleCardProps> {
        return (
            <div>{this.state.profileData.displayName} {this.state.profileData.emailAddress}</div>
        );
    }
}