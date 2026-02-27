import * as React from 'react';
import { IBirthday } from '../models/IBirthday';
import { formatBirthdayDate, getDaysUntilMessage } from '../utils/dateUtils';
import styles from './BirthdayCalendar.module.scss';

export interface IBirthdayCardProps {
  birthday: IBirthday;
  showAge: boolean;
}

/**
 * Individual birthday / work anniversary card component
 */
export const BirthdayCard: React.FC<IBirthdayCardProps> = ({ birthday }) => {
  const daysMessage = getDaysUntilMessage(birthday.daysUntil);
  const dateDisplay = formatBirthdayDate(birthday.nextOccurrence);

  const cardClasses = [
    styles.birthdayCard,
    birthday.isWorkAnniversary ? styles.workAnniversary : '',
    birthday.isMilestone ? styles.milestone : '',
    birthday.isToday ? styles.isToday : ''
  ].filter(Boolean).join(' ');

  const todayBadgeText = birthday.isWorkAnniversary ? 'Jubileum!' : 'Verjaardag!';

  const typeLabel = birthday.isWorkAnniversary ? 'Werkjubileum' : 'Verjaardag';
  const typeLabelClass = birthday.isWorkAnniversary
    ? `${styles.typeLabel} ${styles.typeLabelWork}`
    : `${styles.typeLabel} ${styles.typeLabelBirthday}`;

  const yearsText = birthday.isWorkAnniversary
    ? `${birthday.yearsCount} jaar in dienst`
    : `Wordt ${birthday.yearsCount} jaar`;

  return (
    <div className={cardClasses}>
      <div className={styles.cardContent}>
        <span className={typeLabelClass}>{typeLabel}</span>

        {birthday.isToday && <span className={styles.todayBadge}>{todayBadgeText}</span>}

        <span className={styles.name}>{birthday.name}</span>

        <div className={styles.dateInfo}>
          <span className={styles.date}>{dateDisplay}</span>
          <span className={styles.yearsLabel}>
            {birthday.isWorkAnniversary ? '💼' : '🎂'} {yearsText}
          </span>
        </div>

        {!birthday.isToday && (
          <div className={styles.daysUntil}>{daysMessage}</div>
        )}

        {birthday.isLeapYearBirthday && (
          <div className={styles.leapYearNote}>Schrikkeljaarsdag</div>
        )}
      </div>
    </div>
  );
};
