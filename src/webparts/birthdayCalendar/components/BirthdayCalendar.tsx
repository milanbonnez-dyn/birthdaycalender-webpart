import * as React from 'react';
import { IBirthdayCalendarProps, IBirthday } from '../models/IBirthday';
import { BirthdayService } from '../services/BirthdayService';
import { BirthdayCard } from './BirthdayCard';
import styles from './BirthdayCalendar.module.scss';

export interface IBirthdayCalendarState {
  loading: boolean;
  error?: string;
  birthdays: IBirthday[];
}

/**
 * Main Birthday Calendar component
 * Fetches and displays upcoming birthdays from SharePoint Events list
 */
export default class BirthdayCalendar extends React.Component<IBirthdayCalendarProps, IBirthdayCalendarState> {
  private birthdayService: BirthdayService;

  constructor(props: IBirthdayCalendarProps) {
    super(props);

    this.birthdayService = new BirthdayService(props.spHttpClient, props.siteUrl);

    this.state = {
      loading: true,
      birthdays: []
    };
  }

  public componentDidMount(): void {
    void this.loadBirthdays();
  }

  public componentDidUpdate(prevProps: IBirthdayCalendarProps): void {
    // Reload if configuration changed
    if (
      prevProps.eventsListTitle !== this.props.eventsListTitle ||
      prevProps.birthdayCategory !== this.props.birthdayCategory ||
      prevProps.daysAhead !== this.props.daysAhead ||
      prevProps.maxItems !== this.props.maxItems ||
      prevProps.leapYearStrategy !== this.props.leapYearStrategy ||
      prevProps.showAge !== this.props.showAge
    ) {
      void this.loadBirthdays();
    }
  }

  private async loadBirthdays(): Promise<void> {
    this.setState({ loading: true, error: undefined });

    try {
      const birthdays = await this.birthdayService.getUpcomingBirthdays(
        this.props.eventsListTitle,
        this.props.birthdayCategory,
        this.props.daysAhead,
        this.props.maxItems,
        this.props.leapYearStrategy,
        this.props.showAge
      );

      this.setState({
        loading: false,
        birthdays,
        error: undefined
      });
    } catch (error: unknown) {
      console.error('BirthdayCalendar: Error loading birthdays', error);
      const errorMessage = error instanceof Error ? error.message : 'Failed to load birthdays';
      this.setState({
        loading: false,
        error: errorMessage
      });
    }
  }

  public render(): React.ReactElement<IBirthdayCalendarProps> {
    const { webPartTitle, showAge, isDarkTheme } = this.props;
    const { loading, error, birthdays } = this.state;

    return (
      <div className={`${styles.birthdayCalendar} ${isDarkTheme ? styles.dark : ''}`}>
        <h2 className={styles.title}>{webPartTitle}</h2>

        {loading && (
          <div className={styles.loading}>
            <div className={styles.spinner} />
            <span>Loading birthdays...</span>
          </div>
        )}

        {error && (
          <div className={styles.error}>
            <span className={styles.errorIcon}>!</span>
            <span>{error}</span>
            {(error.indexOf('niet gevonden') !== -1 || error.indexOf('not found') !== -1 || error.indexOf('404') !== -1) && (
              <div style={{ marginTop: '8px', fontSize: '12px' }}>
                Ga naar <strong>Site-inhoud</strong> en klik op <strong>App toevoegen</strong> om de app op deze site te installeren.
              </div>
            )}
          </div>
        )}

        {!loading && !error && birthdays.length === 0 && (
          <div className={styles.empty}>
            No upcoming birthdays in the next {this.props.daysAhead} days.
          </div>
        )}

        {!loading && !error && birthdays.length > 0 && (
          <div className={styles.birthdayList}>
            {birthdays.map(birthday => (
              <BirthdayCard
                key={`${birthday.id}-${birthday.isWorkAnniversary ? 'work' : 'birthday'}`}
                birthday={birthday}
                showAge={showAge}
              />
            ))}
          </div>
        )}
      </div>
    );
  }
}
