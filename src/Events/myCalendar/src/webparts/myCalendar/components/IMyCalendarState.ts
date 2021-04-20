// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
export interface IMyCalendarState {
    error: string;
    loading: boolean;
    renderedDateTime: Date;
    timeZone?: string;
  }