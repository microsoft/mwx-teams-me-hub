// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Event } from '@microsoft/microsoft-graph-types';

export interface IMyCalendarState {
    error: string;
    loading: boolean;
    renderedDateTime: Date;
    timeZone?: string;
    isOpen:boolean;
    activeEvent: Event;
  }