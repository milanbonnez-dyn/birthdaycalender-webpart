import { LeapYearStrategy, IBirthday, ISharePointBirthdayItem } from '../models/IBirthday';

/**
 * Birthday recurrence calculation utilities
 *
 * Core algorithm:
 * 1. Extract month/day from the original event date
 * 2. Project to current year - if already passed, project to next year
 * 3. Handle Feb 29 specially for non-leap years
 * 4. Calculate days until and age
 */

/**
 * Check if a given year is a leap year
 */
export function isLeapYear(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0);
}

/**
 * Check if a date is February 29
 */
export function isFebruary29(date: Date): boolean {
  return date.getMonth() === 1 && date.getDate() === 29;
}

/**
 * Get the start of day (midnight) for a date, ignoring time component
 */
export function startOfDay(date: Date): Date {
  const result = new Date(date.getTime());
  result.setHours(0, 0, 0, 0);
  return result;
}

/**
 * Calculate the number of days between two dates (ignoring time)
 */
export function daysBetween(from: Date, to: Date): number {
  const fromStart = startOfDay(from);
  const toStart = startOfDay(to);
  const diffMs = toStart.getTime() - fromStart.getTime();
  return Math.round(diffMs / (1000 * 60 * 60 * 24));
}

/**
 * Calculate the next occurrence of a birthday given the original birth date
 *
 * @param birthDate - Original birth date
 * @param referenceDate - Date to calculate from (typically today)
 * @param leapYearStrategy - How to handle Feb 29 birthdays in non-leap years
 * @returns Next occurrence date
 */
export function getNextBirthdayOccurrence(
  birthDate: Date,
  referenceDate: Date,
  leapYearStrategy: LeapYearStrategy
): Date {
  const today = startOfDay(referenceDate);
  const currentYear = today.getFullYear();

  const birthMonth = birthDate.getMonth();
  const birthDay = birthDate.getDate();
  const isFeb29Birthday = birthMonth === 1 && birthDay === 29;

  // Try this year first, then next year
  for (let yearOffset = 0; yearOffset <= 1; yearOffset++) {
    const targetYear = currentYear + yearOffset;
    let occurrenceDate: Date;

    if (isFeb29Birthday && !isLeapYear(targetYear)) {
      // Handle leap year birthday in non-leap year
      if (leapYearStrategy === LeapYearStrategy.February28) {
        occurrenceDate = new Date(targetYear, 1, 28); // Feb 28
      } else {
        occurrenceDate = new Date(targetYear, 2, 1); // Mar 1
      }
    } else {
      occurrenceDate = new Date(targetYear, birthMonth, birthDay);
    }

    occurrenceDate = startOfDay(occurrenceDate);

    // If this occurrence is today or in the future, use it
    if (occurrenceDate >= today) {
      return occurrenceDate;
    }
  }

  // Fallback (should never reach here): next year
  const nextYear = currentYear + 1;
  if (isFeb29Birthday && !isLeapYear(nextYear)) {
    if (leapYearStrategy === LeapYearStrategy.February28) {
      return new Date(nextYear, 1, 28);
    } else {
      return new Date(nextYear, 2, 1);
    }
  }
  return new Date(nextYear, birthMonth, birthDay);
}

/**
 * Calculate age based on birth date and current date
 * Returns the age they will be on their next birthday
 */
export function calculateAge(birthDate: Date, nextBirthday: Date): number {
  return nextBirthday.getFullYear() - birthDate.getFullYear();
}

/** Birthday ages that count as milestones */
const MILESTONE_AGES = [20, 30, 40, 50, 60];

/**
 * Transform a SharePoint Verjaardagen item into a processed Birthday object
 */
export function transformEventToBirthday(
  event: ISharePointBirthdayItem,
  referenceDate: Date,
  leapYearStrategy: LeapYearStrategy,
  showAge: boolean
): IBirthday {
  const originalDate = new Date(event.Geboortedatum);
  const nextOccurrence = getNextBirthdayOccurrence(originalDate, referenceDate, leapYearStrategy);
  const today = startOfDay(referenceDate);
  const daysUntil = daysBetween(today, nextOccurrence);
  const isToday = daysUntil === 0;
  const isLeapYearBirthday = isFebruary29(originalDate);
  const yearsCount = calculateAge(originalDate, nextOccurrence);

  const birthday: IBirthday = {
    id: event.Id,
    name: event.Title,
    originalDate,
    nextOccurrence,
    daysUntil,
    yearsCount,
    startDate: event.DatumEersteWerkdag ? new Date(event.DatumEersteWerkdag) : undefined,
    isToday,
    isLeapYearBirthday,
    isWorkAnniversary: false,
    isMilestone: MILESTONE_AGES.indexOf(yearsCount) !== -1
  };

  if (showAge) {
    birthday.age = yearsCount;
  }

  return birthday;
}

/**
 * Transform a SharePoint Verjaardagen item into a work anniversary object
 * Returns null when DatumEersteWerkdag is missing or it would be a 0-year anniversary
 */
export function transformToWorkAnniversary(
  event: ISharePointBirthdayItem,
  referenceDate: Date,
  leapYearStrategy: LeapYearStrategy
): IBirthday | undefined {
  if (!event.DatumEersteWerkdag) return undefined;

  const originalDate = new Date(event.DatumEersteWerkdag);
  const nextOccurrence = getNextBirthdayOccurrence(originalDate, referenceDate, leapYearStrategy);
  const today = startOfDay(referenceDate);
  const daysUntil = daysBetween(today, nextOccurrence);
  const isToday = daysUntil === 0;
  const yearsCount = calculateAge(originalDate, nextOccurrence);

  // Skip 0-year anniversaries (person started today)
  if (yearsCount === 0) return undefined;

  return {
    id: event.Id,
    name: event.Title,
    originalDate,
    nextOccurrence,
    daysUntil,
    yearsCount,
    startDate: originalDate,
    isToday,
    isLeapYearBirthday: false,
    isWorkAnniversary: true,
    isMilestone: false
  };
}

/**
 * Filter and sort birthdays for display
 *
 * @param events - Raw items from the Verjaardagen list
 * @param referenceDate - Date to calculate from
 * @param daysAhead - Maximum days in the future to include
 * @param maxItems - Maximum number of items to return
 * @param leapYearStrategy - Leap year handling
 * @param showAge - Whether to calculate age
 * @returns Sorted array of upcoming birthdays
 */
export function processAndFilterBirthdays(
  events: ISharePointBirthdayItem[],
  referenceDate: Date,
  daysAhead: number,
  maxItems: number,
  leapYearStrategy: LeapYearStrategy,
  showAge: boolean
): IBirthday[] {
  const allEntries: IBirthday[] = [];

  for (const event of events) {
    const birthday = transformEventToBirthday(event, referenceDate, leapYearStrategy, showAge);
    if (birthday.daysUntil >= 0 && birthday.daysUntil <= daysAhead) {
      allEntries.push(birthday);
    }

    const workAnniversary = transformToWorkAnniversary(event, referenceDate, leapYearStrategy);
    if (workAnniversary && workAnniversary.daysUntil >= 0 && workAnniversary.daysUntil <= daysAhead) {
      allEntries.push(workAnniversary);
    }
  }

  allEntries.sort((a, b) => a.daysUntil - b.daysUntil);
  return allEntries.slice(0, maxItems);
}

/**
 * Format a date for display (e.g., "January 15")
 */
export function formatBirthdayDate(date: Date): string {
  return date.toLocaleDateString(undefined, { month: 'long', day: 'numeric' });
}

/**
 * Get a friendly "days until" message
 */
export function getDaysUntilMessage(daysUntil: number): string {
  if (daysUntil === 0) return 'Today!';
  if (daysUntil === 1) return 'Tomorrow';
  return `In ${daysUntil} days`;
}
