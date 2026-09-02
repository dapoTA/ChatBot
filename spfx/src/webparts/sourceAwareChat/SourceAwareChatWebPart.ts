import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import SourceAwareChat, { ISourceAwareChatProps } from './components/SourceAwareChat';

export interface ISourceAwareChatWebPartProps {
  chatbotUrl: string;
}

export default class SourceAwareChatWebPart
  extends BaseClientSideWebPart<ISourceAwareChatWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISourceAwareChatProps> = React.createElement(
      SourceAwareChat,
      {
        chatbotUrl: this.properties.chatbotUrl || 'https://chatbot.technicalassurance.com',
        loginName: this.context.pageContext.user.loginName || '',
        displayName: this.context.pageContext.user.displayName || ''
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Connect this Web Part to the hosted inSite Assistant.'
          },
          groups: [
            {
              groupName: 'Connection',
              groupFields: [
                PropertyPaneTextField('chatbotUrl', {
                  label: 'Chatbot server URL',
                  description: 'For example: https://chatbot.technicalassurance.com'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}