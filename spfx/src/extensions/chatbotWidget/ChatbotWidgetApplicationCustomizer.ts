import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';

const LOG_SOURCE: string = 'ChatbotWidgetApplicationCustomizer';
const SCRIPT_ID: string = 'on-pnt-chatbot-widget';
const DEFAULT_CHATBOT_URL: string = 'https://chatbot.technicalassurance.com';

export interface IChatbotWidgetApplicationCustomizerProperties {
  /** Override the chatbot server URL. Defaults to https://chatbot.technicalassurance.com */
  chatbotUrl: string;
}

/**
 * Injects the ON-PNT floating chat widget on every SharePoint Online modern page.
 * The script tag is injected once per page load; duplicate injection is prevented by ID check.
 */
export default class ChatbotWidgetApplicationCustomizer
  extends BaseApplicationCustomizer<IChatbotWidgetApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ChatbotWidgetApplicationCustomizer');

    // Avoid injecting the script more than once per page
    if (document.getElementById(SCRIPT_ID)) {
      return Promise.resolve();
    }

    const chatbotUrl = (this.properties.chatbotUrl || DEFAULT_CHATBOT_URL).replace(/\/$/, '');

    const script: HTMLScriptElement = document.createElement('script');
    script.id = SCRIPT_ID;
    script.src = `${chatbotUrl}/embed.js`;
    script.async = true;

    document.head.appendChild(script);

    Log.info(LOG_SOURCE, `Injected embed script from ${script.src}`);

    return Promise.resolve();
  }
}
