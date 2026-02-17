import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
    PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SpfxCmDetailsWebPartStrings';
import SpfxCmDetails from './components/SpfxCmDetails';
import { ISpfxCmDetailsProps } from './components/ISpfxCmDetailsProps';
import { getSP } from '../../pnpConfig';

export interface ISpfxCmDetailsWebPartProps {
    description: string;
    context: WebPartContext;
    prefLang: string;
}

export default class SpfxCmDetailsWebPart extends BaseClientSideWebPart<ISpfxCmDetailsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ISpfxCmDetailsProps> = React.createElement(
      SpfxCmDetails,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          context: this.context,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
          userDisplayName: this.context.pageContext.user.displayName,
        prefLang: this.properties.prefLang,
      }
    );

    ReactDom.render(element, this.domElement);
  }

    protected async onInit(): Promise<void> {

      const linkId = "fontawesome-cdn";
      if (!document.getElementById(linkId)) {
        const link = document.createElement("link");
        link.id = linkId;
        link.rel = "stylesheet";
        link.href = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css";
        document.head.appendChild(link);
      }

      const fontId = "inter-font";
      if (!document.getElementById(fontId)) {
        const link = document.createElement("link");
        link.id = fontId;
        link.rel = "stylesheet";
        link.href =
          "https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap";
        document.head.appendChild(link);
      }

      await super.onInit();

      getSP(this.context);
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

  protected onDispose(): void {
    const link = document.getElementById("fontawesome-cdn");
    if (link && link.parentNode) {
      link.parentNode.removeChild(link);
    }

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
                  PropertyPaneDropdown('prefLang', {
                      label: 'Preferred Language',
                      options: [
                          { key: 'account', text: 'Account' },
                          { key: 'en-us', text: 'English' },
                          { key: 'fr-fr', text: 'Français' }
                      ]
                  }),
              ]
            }
          ]
        }
      ]
    };
  }
}
