import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import styles from './NewsPublishedByWebPart.module.scss';
import * as strings from 'NewsPublishedByWebPartStrings';

export interface INewsPublishedByWebPartProps {
  description: string;
}

export default class NewsPublishedByWebPart extends BaseClientSideWebPart<INewsPublishedByWebPartProps> {

  public async render(): Promise<void> {
    this._loadFonts();

    const userEmail = this.context.pageContext.user.email;

    if (!userEmail) {
      this.domElement.innerHTML = `<div>Could not retrieve user email address.</div>`;
      return;
    }

    try {
      const graphClient = await this.context.msGraphClientFactory.getClient("3");

      const user = await graphClient
        .api(`/users/${userEmail}`)
        .select("displayName,jobTitle,userPrincipalName")
        .get();

      const name = user.displayName || 'Unknown';
      const job = user.jobTitle || '';
      const email = user.userPrincipalName || '';

      const photoUrl = `${this.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(email)}`;

      this.domElement.innerHTML = `
        <div class="${styles.publishedByLabel}">Publicerad av:</div>
        <div class="${styles.authorBlock}">
          <img class="${styles.avatar}" src="${photoUrl}" alt="${name}" 
              onerror="this.src='https://static2.sharepointonline.com/files/fabric/assets/persona/persona-placeholder.svg'" />
          <div class="${styles.authorInfo}">
            <div class="${styles.name}">${name}</div>
            <div class="${styles.job}">${job}</div>
          </div>
        </div>
      `;

    } catch (error) {
      console.error('❌ Error while fetching data from Microsoft Graph:', error);
      this.domElement.innerHTML = `<div>Error loading author information.</div>`;
    }
  }

  protected onInit(): Promise<void> {
    return Promise.resolve();
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
    private _loadFonts(): void {
    const fontId = 'google-font-lato';
    if (!document.getElementById(fontId)) {
      const link1 = document.createElement('link');
      link1.rel = 'preconnect';
      link1.href = 'https://fonts.googleapis.com';
      document.head.appendChild(link1);

      const link2 = document.createElement('link');
      link2.rel = 'preconnect';
      link2.href = 'https://fonts.gstatic.com';
      link2.crossOrigin = 'anonymous';
      document.head.appendChild(link2);

      const link3 = document.createElement('link');
      link3.id = fontId;
      link3.rel = 'stylesheet';
      link3.href = 'https://fonts.googleapis.com/css2?family=Lato:ital,wght@0,300;0,400;0,700&display=swap';
      document.head.appendChild(link3);
    }
  }
}
