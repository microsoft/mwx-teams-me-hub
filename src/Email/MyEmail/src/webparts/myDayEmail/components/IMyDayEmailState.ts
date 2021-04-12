import { IMessage, IMessageDetails } from '.';

export interface IMyDayEmailState {
  error: string;
  loading: boolean;
  messages: IMessage[];
  renderedDateTime: Date;
  filter: string;
  isOpen: boolean;
  activeMessage: IMessageDetails;
}