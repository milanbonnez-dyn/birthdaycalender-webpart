import { SPHttpClient } from '@microsoft/sp-http';

/**
 * Represents an item from the SharePoint Verjaardagen list
 */
export interface ISharePointBirthdayItem {
  Id: number;
  Title: string;
  Geboortedatum: string;
  DatumEersteWerkdag?: string;
}

/**
 * @deprecated Use ISharePointBirthdayItem for the Verjaardagen list
 * Represents a birthday event from SharePoint Events list
 */
export interface ISharePointEvent {
  Id: number;
  Title: string;
  EventDate: string;
  EndDate: string;
  Description?: string;
  Category?: string;
  fRecurrence?: boolean;
  RecurrenceData?: string;
}

/**
 * Leap year handling strategies for Feb 29 birthdays
 */
export enum LeapYearStrategy {
  /** Show on Feb 28 in non-leap years */
  February28 = 'feb28',
  /** Show on Mar 1 in non-leap years */
  March1 = 'mar1'
}

/**
 * Processed birthday with calculated next occurrence
 */
export interface IBirthday {
  id: number;
  name: string;
  originalDate: Date;
  nextOccurrence: Date;
  daysUntil: number;
  /** Years turning (age for birthdays, years of service for work anniversaries) */
  yearsCount: number;
  age?: number;
  startDate?: Date;
  isToday: boolean;
  isLeapYearBirthday: boolean;
  /** True for work anniversary entries (based on DatumEersteWerkdag) */
  isWorkAnniversary: boolean;
  /** True for milestone birthdays: 20, 30, 40, 50, 60 */
  isMilestone: boolean;
}

/**
 * Web part configuration properties
 */
export interface IBirthdayCalendarWebPartProps {
  /** SharePoint Events list title (default: 'Events') */
  eventsListTitle: string;
  /** Category filter to identify birthday events (default: 'Birthday') */
  birthdayCategory: string;
  /** Number of days ahead to show (default: 30) */
  daysAhead: number;
  /** Show calculated age (default: false) */
  showAge: boolean;
  /** Leap year handling strategy (default: feb28) */
  leapYearStrategy: LeapYearStrategy;
  /** Maximum items to display (default: 10) */
  maxItems: number;
  /** Web part title (default: 'Upcoming Birthdays') */
  webPartTitle: string;
}

/**
 * Props passed to the main BirthdayCalendar component
 */
export interface IBirthdayCalendarProps {
  eventsListTitle: string;
  birthdayCategory: string;
  daysAhead: number;
  showAge: boolean;
  leapYearStrategy: LeapYearStrategy;
  maxItems: number;
  webPartTitle: string;
  siteUrl: string;
  spHttpClient: SPHttpClient;
  isDarkTheme: boolean;
}
