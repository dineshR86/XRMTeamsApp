import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from '@microsoft/sp-loader';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import * as microsoftTeams from '@microsoft/teams-js';

import * as strings from 'XrmTeamsAppWebPartStrings';
import {XrmTeamsApp,IXrmTeamsAppProps} from './components/XrmTeamsApp';
import { graphservice } from './service/graphservice';


export interface IXrmTeamsAppWebPartProps {
  description: string;
}

export default class XrmTeamsAppWebPart extends BaseClientSideWebPart<IXrmTeamsAppWebPartProps> {

  private _teamsContext:microsoftTeams.Context;
  private _graphservice:graphservice;

  protected onInit(): Promise<any> {
    SPComponentLoader.loadCss("https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css");
    console.log("webpart init");
    let retVal: Promise<any> = Promise.resolve();
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }

    this._graphservice=new graphservice(this.context.msGraphClientFactory);
    return retVal;
  }

  public render(): void {
    const element: React.ReactElement<IXrmTeamsAppProps > = React.createElement(
      XrmTeamsApp,
      {
        description: this.properties.description,
        teamsContext:this._teamsContext,
        graphservice:this._graphservice
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
