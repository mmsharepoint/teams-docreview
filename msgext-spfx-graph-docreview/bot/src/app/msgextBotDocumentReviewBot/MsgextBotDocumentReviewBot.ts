import { BotDeclaration, MessageExtensionDeclaration, PreventIframe } from "express-msteams-host";
import * as debug from "debug";
import { DialogSet, DialogState } from "botbuilder-dialogs";
import { StatePropertyAccessor, CardFactory, TurnContext, MemoryStorage, ConversationState, ActivityTypes, TeamsActivityHandler, MessagingExtensionAction, CardAction, CardImage, MessagingExtensionActionResponse } from "botbuilder";
import HelpDialog from "./dialogs/HelpDialog";
import WelcomeCard from "./dialogs/WelcomeDialog";

// Initialize debug logging module
const log = debug("msteams");

/**
 * Implementation for Msgext Bot Document Review Bot
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID,
    process.env.MICROSOFT_APP_PASSWORD)

export class MsgextBotDocumentReviewBot extends TeamsActivityHandler {
  private readonly conversationState: ConversationState;
  private readonly dialogs: DialogSet;
  private dialogState: StatePropertyAccessor<DialogState>;

  /**
   * The constructor
   * @param conversationState
   */
  public constructor(conversationState: ConversationState) {
    super();        
  }

  /**
   * Handles the task/fetch operation
   * That is when the Bot hero card calls the 'Review' button
   * @param context 
   * @param value 
   */
  protected handleTeamsTaskModuleFetch(context: TurnContext, value: any): Promise<any> {
    console.log(context);
    console.log(value);
    const componentID = '75f1c63b-e3d1-46b2-957f-3d19a622c463';
    const itemID = value.data.item.key;
    return Promise.resolve({
      task: {
        type: "continue",
        value: {
          title: "Mark document as reviewed",
          height: 300,
          width: 400,
          url: `https://{teamSiteDomain}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/_layouts/15/teamstaskhostedapp.aspx%3Fteams%26personal%26componentId=${componentID}%26forceLocale={locale}%26itemID=${itemID}`,
          fallbackUrl: `https://{teamSiteDomain}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/_layouts/15/teamstaskhostedapp.aspx%3Fteams%26personal%26componentId=${componentID}%26forceLocale={locale}`
        }
      }
    });
  }

  /**
   * Handles the selection of a document from the task module
   * @param context 
   * @param action 
   */
  protected async handleTeamsMessagingExtensionSubmitAction(context: TurnContext, action: MessagingExtensionAction): Promise<any> {
    const viewAction: CardAction = { title: 'View', type: 'openUrl', value: action.data.url };
    const revAction: CardAction =  { title: 'Reviewed', type: 'invoke', value: { type: 'task/fetch', item: action.data } };
    const revImage: CardImage = { url: `https://${process.env.HOSTNAME}/assets/icon.png` };   
    
    const thumbCard = CardFactory.thumbnailCard(action.data.name, action.data.description, [revImage], [viewAction, revAction]);
    thumbCard.content.subtitle = action.data.author;
    const response: MessagingExtensionActionResponse = {
      composeExtension: {
        type: 'result',
        attachmentLayout: 'grid',
        attachments:  [ thumbCard ]
      }
    }
    return Promise.resolve(response);
  }
}
