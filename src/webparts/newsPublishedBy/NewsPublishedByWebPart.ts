import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient } from '@microsoft/sp-http';


import styles from './NewsPublishedByWebPart.module.scss';
import * as strings from 'NewsPublishedByWebPartStrings';

export interface INewsPublishedByWebPartProps {
  description: string;
}

export default class NewsPublishedByWebPart extends BaseClientSideWebPart<INewsPublishedByWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

public async render(): Promise<void> {
  const pageUrl = this.context.pageContext.site.serverRequestPath;

  const response = await this.context.spHttpClient.get(
    `${this.context.pageContext.web.absoluteUrl}/_api/web/getfilebyserverrelativeurl('${pageUrl}')/ListItemAllFields?$expand=Author&$select=Author/Title,Author/JobTitle,Author/EMail`,
    SPHttpClient.configurations.v1
  );


  const item = await response.json();
  const author = item.Author;

  const name = author?.Title || 'Okänd';
  const job = author?.JobTitle || '';
  const email = author?.EMail || '';
  const photoUrl = `${this.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(email)}`;

  this.domElement.innerHTML = `
    <div class="${styles.authorBlock}">
      <img class="${styles.avatar}" src="${photoUrl}" alt="${name}" />
      <div class="${styles.authorInfo}">
        <div class="${styles.name}">${name}</div>
        <div class="${styles.job}">${job}</div>
      </div>
    </div>
  `;
}


  protected onInit(): Promise<void> {
    this._loadFonts();
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
    console.log(this._environmentMessage,this._isDarkTheme)
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


  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
