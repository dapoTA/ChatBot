import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';

const LOG_SOURCE: string = 'ChatbotWidgetApplicationCustomizer';
const FRAME_ID: string = 'on-pnt-chatbot-frame';
const DEFAULT_CHATBOT_URL: string = 'https://chatbot.technicalassurance.com';

// iframe sizes
const CLOSED_W: string = '100px';
const CLOSED_H: string = '100px';
const OPEN_W: string = '480px';
const OPEN_H: string = '660px';
const MINIMIZED_W: string = '320px';
const MINIMIZED_H: string = '80px';

export interface IChatbotWidgetApplicationCustomizerProperties {
  /** Override the chatbot server URL. Defaults to https://chatbot.technicalassurance.com */
  chatbotUrl: string;
}

/**
 * Injects the ON-PNT floating chat widget on every SharePoint Online modern page.
 * All DOM work is done inside the trusted SPFx bundle to avoid CSP restrictions on
 * external script-src. The iframe src (the chatbot server) is not subject to script-src CSP.
 */
export default class ChatbotWidgetApplicationCustomizer
  extends BaseApplicationCustomizer<IChatbotWidgetApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ChatbotWidgetApplicationCustomizer');

    if (!document.getElementById(FRAME_ID)) {
      this._injectWidget();
    }

    return Promise.resolve();
  }

  private _injectWidget(): void {
    const chatbotUrl: string = (this.properties.chatbotUrl || DEFAULT_CHATBOT_URL).replace(/\/$/, '');

    // Append the current user's login name so the chatbot can log it.
    // pageContext.user.loginName is e.g. "dapo@technicalassurance.com"
    const loginName: string = this.context.pageContext.user.loginName || '';
    const iframeSrc: string = loginName
      ? `${chatbotUrl}?spuser=${encodeURIComponent(loginName)}`
      : chatbotUrl;

    const frame: HTMLIFrameElement = document.createElement('iframe');
    frame.id = FRAME_ID;
    frame.src = iframeSrc;
    frame.title = 'ON-PNT ChatBot';
    frame.setAttribute('allowtransparency', 'true');
    frame.setAttribute('frameborder', '0');
    frame.setAttribute('scrolling', 'no');

    Object.assign(frame.style, {
      position:     'fixed',
      bottom:       '0',
      right:        '0',
      width:        CLOSED_W,
      height:       CLOSED_H,
      border:       'none',
      background:   'transparent',
      zIndex:       '2147483647',
      pointerEvents:'all',
      overflow:     'hidden',
    });

    document.body.appendChild(frame);

    // Resize iframe when the widget posts open/close/minimize messages
    const origin: string = new URL(chatbotUrl).origin;
    window.addEventListener('message', (event: MessageEvent) => {
      if (event.origin !== origin) { return; }

      if (event.data === 'chatbot:open') {
        frame.style.width  = OPEN_W;
        frame.style.height = OPEN_H;
      } else if (event.data === 'chatbot:minimize') {
        frame.style.width  = MINIMIZED_W;
        frame.style.height = MINIMIZED_H;
      } else if (event.data === 'chatbot:close') {
        frame.style.width  = CLOSED_W;
        frame.style.height = CLOSED_H;
      }
    });

    Log.info(LOG_SOURCE, `Chatbot iframe injected from ${chatbotUrl}`);
  }
}
