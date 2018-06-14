import * as React from 'react';
import styles from './PeopleFinder.module.scss';
import {
    DocumentCard,
    DocumentCardActivity,
    DocumentCardPreview,
    DocumentCardTitle,
    IDocumentCardPreviewProps
  } from 'office-ui-fabric-react/lib/DocumentCard';

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
            <DocumentCard>
                <div className="ms-Grid">
                    <div className={styles.row}>
                        <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg2">
                            <img src={this.state.profileData.pictureUrl} width="50" />
                        </div>
                        <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg10">
                            <div className={styles.title}>{this.state.profileData.displayName}</div>
                            <div>{this.state.profileData.title}</div>
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                            {this.state.profileData.emailAddress}
                        </div>
                    </div>
                </div>                
            </DocumentCard>
        );
    }
}