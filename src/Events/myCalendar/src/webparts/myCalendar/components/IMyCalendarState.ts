// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Event } from '@microsoft/microsoft-graph-types';
import { IMeeting } from './IMeeting';

export interface IMyCalendarState {
    error: string;
    loading: boolean;
    renderedDateTime: Date;
    timeZone?: string;
    isOpen:boolean;
    meetings: IMeeting[];
    activeEvent: Event;
  }