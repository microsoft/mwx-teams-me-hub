// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { isElementFocusSubZone } from "office-ui-fabric-react";

export interface IMessages {
  value: IMessage[];
}
  
export interface IEmailAddress {
  emailAddress: {
    name: string;
    address:string;
  };
}

export interface IMessage {
  bodyPreview: string;
  from: {
    emailAddress: {
      address: string;
      name: string;
    }
  };
  isRead: boolean;
  receivedDateTime: string;
  subject: string;
  webLink: string;
  id: string;
  importance: string;
  flag: {
    flagStatus: string
  };
  hasAttachments: boolean;
  meetingMessageType: string;
}

export interface IMessageDetails extends IMessage {
  body: {
    contentType: string;
    content: string;
  };
  toRecipients: IEmailAddress[];
  ccRecipients: IEmailAddress[];
}
