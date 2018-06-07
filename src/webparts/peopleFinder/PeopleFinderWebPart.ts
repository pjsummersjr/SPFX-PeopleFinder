import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { Environment, EnvironmentType } from "@microsoft/sp-core-library"

import * as strings from 'PeopleFinderWebPartStrings';
import PeopleFinder from './components/PeopleFinder';
import { IPeopleFinderProps } from './components/IPeopleFinderProps';
import { IProfileService, ProfileService,  MockProfileService } from './services/ProfileService';

export interface IPeopleFinderWebPartProps {
  description: string;
}

export default class PeopleFinderWebPart extends BaseClientSideWebPart<IPeopleFinderWebPartProps> {

  public render(): void {

    let profileService: IProfileService  = new MockProfileService();
    if(Environment.type === EnvironmentType.SharePoint){
      profileService = new ProfileService(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl);
    }

    const element: React.ReactElement<IPeopleFinderProps > = React.createElement(
      PeopleFinder,
      {
        description: this.properties.description,
        profileService: profileService
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
