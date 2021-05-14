// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as React from 'react';
import styles from './MyDayEmail.module.scss';
import * as strings from 'MyDayEmailWebPartStrings';
import { HeaderDisplay, EMailDisplay, ClickAction } from '../enums';
import { IMyDayEmailProps, IMyDayEmailState, IMessage, IMessages, IMessageDetails, IEmailAddress } from '.';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/components/Spinner';
import { CommandBar, ICommandBarItemProps } from '@fluentui/react/lib/CommandBar';
import { Label } from '@fluentui/react/lib/Label';
import { Stack, IStackStyles, IStackTokens, IStackItemStyles } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { List } from '@fluentui/react/lib/components/List';
import { Link } from '@fluentui/react/lib/components/Link';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { PrimaryButton, IIconProps, Separator } from 'office-ui-fabric-react';
import { FontIcon } from '@fluentui/react/lib/Icon';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { initializeIcons } from '@uifabric/icons';
import { mergeStyles, mergeStyleSets } from '@fluentui/react/lib/Styling';
import { GraphRequest } from '@microsoft/sp-http';
import { DisplayMode } from '@microsoft/sp-core-library';

export class MyDayEmail extends React.Component<IMyDayEmailProps, IMyDayEmailState> {

  private _interval: number;

  private _messageIconClass: string = mergeStyles({
    fontSize: 16,
    height: 16,
    width: 16,
    margin: '0',
  });

  private _classNames = mergeStyleSets({
    attachment: [this._messageIconClass],
    highImportance: [{ color: 'red' }, this._messageIconClass],
    lowImportance: [this._messageIconClass],
    flagged: [{ color: 'red' }, this._messageIconClass],
    completed: [{ color: 'green' }, this._messageIconClass]
  });

  

  constructor(props: IMyDayEmailProps) {
    super(props);  

    initializeIcons();

    this.state = {
      messages: [],
      loading: false,
      error: undefined,
      renderedDateTime: new Date(),
      filter: (props.emailDisplay == EMailDisplay.Default) ? strings.AllPivot : props.emailDisplay,
      isOpen: false,
      activeMessage: {} as IMessageDetails
    };
  }

  /**
   * Load recent messages for the current user
   */
  private _loadMessages(): void {
    if (!this.props.graphClient) {
      return;
    }

    // update state to indicate loading and remove any previously loaded
    // messages
    this.setState({
      error: null,
      loading: true,
      messages: []
    });   

    const request: GraphRequest = this.props.graphClient
      .api("me/mailFolders/Inbox/messages")
      .version("v1.0")
      .select("id,bodyPreview,receivedDateTime,from,subject,webLink,isRead,importance,flag,hasAttachments") //,meetingMessageType             
      .top(this.props.numberOfMessages || 5);

    // Graph API does not like ordering when we are viewing only flagged items.
    if ((this.state.filter == strings.AllPivot) || (this.state.filter == strings.UnreadPivot)) {
        request.orderby("receivedDateTime desc");
    }

    if (this.state.filter == strings.UnreadPivot) {
      request.filter("isRead eq false");
    }
    else if (this.state.filter == EMailDisplay.Important) {
      request.filter("importance eq 'high'");
    }
    else if (this.state.filter == EMailDisplay.Flagged) {
      request.filter("flag/flagStatus eq 'flagged'");
    }
      
    console.log(`email request: filter = ${this.state.filter} - url: ${request.buildFullUrl()}`);

    request.get((err: any, res: IMessages): void => {
      if (err) {
        // Something failed calling the MS Graph
        this.setState({
          error: err.message ? err.message : strings.Error,
          loading: false
        });
        return;
      }

      // Check if a response was retrieved
      if (res && res.value && res.value.length > 0) {
        this.setState({
          messages: res.value,
          loading: false
        });
      }
      else {
        // No messages found
        this.setState({
          loading: false
        });
      }
    });
    
  }  
  private _handlePivotChange = (item: PivotItem): void => {
    console.log(`Pivot Change: ${item.props.itemKey}`);

    this.setState({ 
      filter: item.props.itemKey 
    });

    this._reRender();
  }

  private _onNewEmail = (): void => {
    window.open("https://outlook.office.com/?path=/mail/action/compose", "_blank");
  }

  private _showMessageDetails = (messageId: string, webLink: string): void => {    
    if (this.props.clickAction === ClickAction.OpenInOutlook) {
      window.open(webLink, "_blank");
      return;
    }

    const request: GraphRequest = this.props.graphClient
      .api(`me/mailFolders/Inbox/messages/${messageId}`)
      .version("v1.0")
      .select("id,bodyPreview,receivedDateTime,from,subject,webLink,isRead,importance,flag,hasAttachments,body,toRecipients,ccRecipients"); //,meetingMessageType            

    request
      .get()
      .then((result: IMessageDetails) => {          
        this.setState({
          isOpen: true,
          activeMessage: result,
          messages: ((this.props.clickAction === ClickAction.PreviewRead) && (!result.isRead)) ? 
            this._changeMessageReadStatus(messageId, true) : 
            this.state.messages
        });                      
      })
      .catch((err) => {
        this.setState({
          error: err.message ? err.message : strings.Error,            
        });
      });
  }

  private _changeMessageReadStatus = (messageId: string, isRead: boolean):IMessage[] => {
    this.props.graphClient
      .api(`me/messages/${messageId}`)
      .version("v1.0")
      .patch( { isRead: isRead });

    this.state.messages.forEach((message: IMessage) => {
      if (message.id === messageId) {
        message.isRead = isRead;
      }        
    });
    
    return this.state.messages;
  }

  private _deleteMessage = (messageId: string, isRead: boolean):IMessage[] => {
    this.props.graphClient
      .api(`me/messages/${messageId}`)
      .version("v1.0")
      .delete();      

    const messages: IMessage[] = [];    
    this.state.messages.forEach((message: IMessage) => {
      if (message.id !== messageId) {
        messages.push(message);
      }
    });

    return messages;  
  }

  private _reRender = (): void => {
    //this._loadMessages();

    // update the render date to force reloading data and re-rendering
    // the component    
    this.setState({ renderedDateTime: new Date() });
  }

  private _setInterval = (): void => {
    let { refreshInterval } = this.props;
    // set up safe default if the specified interval is not a number
    // or beyond the valid range
    if (isNaN(refreshInterval) || refreshInterval < 0 || refreshInterval > 60) {
      refreshInterval = 5;
    }
    // refresh the component every x minutes
    this._interval = setInterval(this._reRender, refreshInterval * 1000 * 60);
    this._reRender();
  }

 

  public componentDidMount(): void {
    // load data initially after the component has been instantiated
    this._loadMessages();
    this._setInterval();
  }

  public componentWillUnmount(): void {
    // remove the interval so that the data won't be reloaded
    clearInterval(this._interval);
  }

  public componentDidUpdate(prevProps: IMyDayEmailProps, prevState: IMyDayEmailState): void {

    if (prevProps.refreshInterval !== this.props.refreshInterval) {
      clearInterval(this._interval);
      this._setInterval();
      return;
    }

    // verify if the component should update. Helps avoid unnecessary re-renders
    // when the parent has changed but this component hasn't
    if ((prevProps.numberOfMessages !== this.props.numberOfMessages) || 
        (prevState.renderedDateTime !== this.state.renderedDateTime) || 
        (prevState.filter !== this.state.filter)) {
      this._loadMessages();
    }
  }

  public render(): React.ReactElement<IMyDayEmailProps> {  
    if (this._editModeRefresh()) {
      return null;
    }

    const { semanticColors, fonts }: IReadonlyTheme = this.props.themeVariant;

    console.log(semanticColors);
    console.log(fonts);

    const recipientStackTokens: IStackTokens = {
      childrenGap: 7    
    };
  
    const messageDetailsStackTokens: IStackTokens = {
      childrenGap: 3    
    };

    const messageDetailsCommandBarFarItems: ICommandBarItemProps[] = [      
      {
        key: 'viewInOutlook',
        text: 'View in Outlook',
        iconProps: { iconName: 'OpenInNewWindow' },
        href: this.state.activeMessage.webLink,
        target: '_blank'
      }
    ];    

    return (
      <div className={styles.myDayEmail} style={{ backgroundColor: semanticColors.bodyBackground }}>
        {
          this.props.headerDisplay != HeaderDisplay.None &&           
          <WebPartTitle 
            displayMode={this.props.displayMode}                  
            title={this.props.title}
            className={ (this.props.headerDisplay == HeaderDisplay.Standard) ? `${styles.webPartTitle} ${styles.webPartTitleStandard}` : styles.webPartTitle}
            updateProperty={this.props.updateProperty} //className={styles.title}
            themeVariant={this.props.themeVariant}
            moreLink={ (this.props.showViewAll) ? <Link href="https://outlook.office.com/owa/" target="_blank" style={{ color: semanticColors.link }}>See all</Link> : null } 
          />
        }
        {
          this.state.loading &&
          <Spinner label={strings.Loading} size={SpinnerSize.large} />
        }
        { 
          (!this.state.loading && this.props.showNew) ? 
          //don't render the button till we figure out how to go straight to new mail
          <PrimaryButton iconProps={{iconName: 'NewMail'}} onClick={this._onNewEmail} disabled={false} >
            New message
          </PrimaryButton> : null
        }
        {
          (!this.state.loading && this.props.emailDisplay == EMailDisplay.Default) ? 
          <Pivot                  
            onLinkClick={this._handlePivotChange}
            selectedKey={this.state.filter}
            headersOnly={true}>
            <PivotItem 
              headerText={strings.AllPivot} 
              itemKey={strings.AllPivot} 
              style={{ color: semanticColors.bodyText}} />
            <PivotItem 
              headerText={strings.UnreadPivot} 
              itemKey={strings.UnreadPivot} 
              style={{ color: semanticColors.bodyText}} />                
          </Pivot> : null
        }
        {
          this.state.messages &&
            this.state.messages.length > 0 ? (
              <div>                                          
              
              <List items={this.state.messages}
                onRenderCell={this._onRenderEmailCell} className={styles.list} />
              {
                this.props.showViewAll && this.props.headerDisplay == HeaderDisplay.None && 
                <Link href='https://outlook.office.com/owa/' target='_blank' className={styles.viewAll}>{strings.ViewAll}</Link>
              }
            </div>
            ) : (
              !this.state.loading && (
                this.state.error ?
                  <div className={styles.error}>{this.state.error}</div> :                  
                  <div className={styles.noMessages}>{strings.NoMessages}</div>
                  
              )
            )
        }
        {
          (true && this.state.isOpen) ? 
        <Panel
          className={styles.messageDetails}
          isLightDismiss          
          isOpen={this.state.isOpen}
          onDismiss={() => { this.setState( { isOpen: false }); } }
          type={PanelType.largeFixed}          
          closeButtonAriaLabel="Close"
          headerText={this.state.activeMessage.subject}
        >   
          <Stack tokens={messageDetailsStackTokens}>            
            <CommandBar
              items={this._getMessageDetailsCommandBarItems()}
              farItems={messageDetailsCommandBarFarItems}
              className={styles.commandBar}
              
              ariaLabel="Use left and right arrow keys to navigate between commands"
            />
            <Text>
              {this.state.activeMessage.from.emailAddress.name || this.state.activeMessage.from.emailAddress.address}
            </Text>          
            <Text>
              {new Date(this.state.activeMessage.receivedDateTime).toLocaleDateString()} 
              {new Date(this.state.activeMessage.receivedDateTime).toLocaleTimeString()}
            </Text>          
            <Stack horizontal disableShrink className={styles.recipients} tokens={recipientStackTokens}>
                <Text className={styles.prompt}>To: </Text>
                {this.state.activeMessage.toRecipients.map((recipient, index) =>(
                  <Text className={styles.recipient}>
                    {(recipient.emailAddress.name || recipient.emailAddress.address)}
                    {(index !== this.state.activeMessage.toRecipients.length -1) ? "; " : ""}
                  </Text>      
                ))}             
            </Stack>
            {  (this.state.activeMessage.ccRecipients != null && this.state.activeMessage.ccRecipients.length > 0) ?
            <Stack horizontal disableShrink className={styles.recipients} tokens={recipientStackTokens}>
                <Text className={styles.prompt}>Cc: </Text>
                {this.state.activeMessage.ccRecipients.map((recipient, index) =>(
                  <Text className={styles.recipient}>
                    {(recipient.emailAddress.name || recipient.emailAddress.address)}
                    {(index !== this.state.activeMessage.ccRecipients.length -1) ? "; " : ""}
                  </Text>      
                ))}             
            </Stack> : null
            }                  
            {  (this.state.activeMessage.body.contentType === "html") ?           
            <div dangerouslySetInnerHTML={{ __html: this.state.activeMessage.body.content }}></div> : 
            <div>{this.state.activeMessage.body.content}</div>
            }        
          </Stack>    
        </Panel> : null
      }
      </div>
    );
  }
  
  

   /**
   * Render message item
   */
  private _onRenderEmailCell = (item: IMessage, index: number | undefined): JSX.Element => {
    const { semanticColors, fonts }: IReadonlyTheme = this.props.themeVariant;
    
    var messageStyle:string = styles.message;

    if (item.isRead) {
      messageStyle = styles.message + " " + styles.isRead;
    }

    return (     
      <Link className={messageStyle} onClick={ () => this._showMessageDetails(item.id, item.webLink) }>
         <div className={styles.from} style={{ color: semanticColors.bodyText }}>
          {(item.from.emailAddress.name || item.from.emailAddress.address)}           
        </div>
        <div className={styles.icons}>
          {(item.hasAttachments) ? <FontIcon iconName="Attach" className={this._classNames.attachment} />: null}          
          {(item.importance == "high") ? <FontIcon iconName="Important" className={this._classNames.highImportance} /> : null}
          {(item.importance == "low") ? <FontIcon iconName="SortDown" className={this._classNames.lowImportance} /> : null}
          {(item.flag.flagStatus == "flagged") ? <FontIcon iconName="Flag" className={this._classNames.flagged} /> : null }
          {(item.flag.flagStatus == "complete") ? <FontIcon iconName="CheckMark" className={this._classNames.completed} />: null }          
        </div>
        <div className={styles.subject}>{item.subject}</div>
        <div className={styles.date}>{(new Date(item.receivedDateTime).toLocaleDateString())}</div>
        <div className={styles.preview}>{item.bodyPreview}</div>        
      </Link>      
    );
  } 

  private _getMessageDetailsCommandBarItems = (): ICommandBarItemProps[] => {
    const items: ICommandBarItemProps[] = [];    
    // if (this.props.enableToggleReadStatus) {
    //   items.push(
    //     {
    //       key: 'followUp',
    //       text: 'Follow Up',
    //       iconProps: { iconName: 'Flag' },
    //       href: this.state.activeMessage.webLink,
    //       target: '_blank'
    //     }
    //   )
    // }

    return items;
  }

  
  private _editModeRefresh = (): boolean => {    
    //If editing the web part
    if ((this.props.displayMode === DisplayMode.Edit) && (
      // Property is flagged but filter is not flagged
      ((this.props.emailDisplay  === EMailDisplay.Flagged) && (this.state.filter !== EMailDisplay.Flagged)) ||
      // Property is important but filter is not important
      ((this.props.emailDisplay  === EMailDisplay.Important) && (this.state.filter !== EMailDisplay.Important)) ||
      // Property is default but filter is not all or unread
      ((this.props.emailDisplay === EMailDisplay.Default) && (this.state.filter !== strings.AllPivot) && (this.state.filter !== strings.UnreadPivot)))) {
        this.setState( { filter : (this.props.emailDisplay === EMailDisplay.Default) ? strings.AllPivot : this.props.emailDisplay });
        this._loadMessages();
        return true;
      }

    return false;    
  }
}
