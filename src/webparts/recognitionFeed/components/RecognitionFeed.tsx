import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import styles from './RecognitionFeed.module.scss';
import type { IRecognitionFeedProps } from './IRecognitionFeedProps';

interface IWinnerItem {
  Id: number;
  Name: string;
  PhotoEmail: string;
  AwardType: string;
  Month: string;
  Year: string;
  Justification: string;
  IsTeamAward: boolean;
  TeamName: string;
  TeamMembers: string;
}

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'
];

const monthNumberToName = (raw: string): string => {
  const num = parseInt(raw, 10);
  if (!isNaN(num) && num >= 1 && num <= 12) {
    return MONTH_NAMES[num - 1];
  }
  return raw || '';
};

const getPhotoUrl = (siteAbsoluteUrl: string, email: string): string => {
  if (!email) {
    return '';
  }
  const origin = new URL(siteAbsoluteUrl).origin;
  return `${origin}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(email)}`;
};

const RecognitionFeed: React.FunctionComponent<IRecognitionFeedProps> = (props) => {
  const [winners, setWinners] = useState<IWinnerItem[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>('');
  const [selectedWinner, setSelectedWinner] = useState<IWinnerItem | null>(null);
  const [isPaused, setIsPaused] = useState<boolean>(false);

  const scrollTrackRef = useRef<HTMLDivElement>(null);

  const fieldsConfigured =
    props.listId &&
    props.fieldName &&
    props.fieldPhotoEmail &&
    props.fieldAwardType &&
    props.fieldMonth &&
    props.fieldYear &&
    props.fieldJustification;

  useEffect(() => {
    if (!fieldsConfigured) {
      setLoading(false);
      return;
    }

    const loadWinners = async (): Promise<void> => {
      setLoading(true);
      setError('');

      try {
        const selectFields = [
          'Id',
          `${props.fieldName}`,
          `${props.fieldPhotoEmail}/EMail`,
          `${props.fieldPhotoEmail}/Title`,
          `${props.fieldAwardType}`,
          `${props.fieldMonth}`,
          `${props.fieldYear}`,
          `${props.fieldJustification}`
        ];

        if (props.fieldTeamFlag) {
          selectFields.push(props.fieldTeamFlag);
        }
        if (props.fieldTeamName) {
          selectFields.push(props.fieldTeamName);
        }
        if (props.fieldTeamMembers) {
          selectFields.push(props.fieldTeamMembers);
        }

        const rawItems = await props.sp.web.lists.getById(props.listId).items
          .select(...selectFields)
          .expand(props.fieldPhotoEmail)
          .orderBy(props.fieldYear, false)
          .orderBy(props.fieldMonth, false)
          .top(props.monthsToShow * 4)();

        const mapped: IWinnerItem[] = rawItems.map((raw: Record<string, unknown>) => {
          const photoField = raw[props.fieldPhotoEmail] as { EMail?: string; Title?: string } | undefined;

          const isTeam = props.fieldTeamFlag
            ? !!raw[props.fieldTeamFlag]
            : false;

          return {
            Id: raw.Id as number,
            Name: (raw[props.fieldName] as string) || '',
            PhotoEmail: photoField && photoField.EMail ? photoField.EMail : '',
            AwardType: (raw[props.fieldAwardType] as string) || '',
            Month: monthNumberToName((raw[props.fieldMonth] as string) || ''),
            Year: (raw[props.fieldYear] as string) || '',
            Justification: (raw[props.fieldJustification] as string) || '',
            IsTeamAward: isTeam,
            TeamName: props.fieldTeamName ? ((raw[props.fieldTeamName] as string) || '') : '',
            TeamMembers: props.fieldTeamMembers ? ((raw[props.fieldTeamMembers] as string) || '') : ''
          };
        });

        // Limit to the requested number of months worth of winners.
        // Assumes roughly 2 winners per month (Critical Cog + Employee of the Month);
        // if that ratio changes, this cap may need to be exposed as its own property later.
        const capped = mapped.slice(0, props.monthsToShow * 2);
        setWinners(capped);
      } catch (err) {
        setError('Unable to load recognition data. Check the Data Source and Field Mapping settings.');
        console.error(err);
      } finally {
        setLoading(false);
      }
    };

    loadWinners().catch((err) => console.error(err));
  }, [props.sp, props.listId, props.fieldName, props.fieldPhotoEmail, props.fieldAwardType, props.fieldMonth, props.fieldYear, props.fieldJustification, props.fieldTeamFlag, props.fieldTeamName, props.fieldTeamMembers, props.monthsToShow]);

  const isCurrentMonth = (item: IWinnerItem, index: number): boolean => {
    // The first 1-2 items (most recent, since sorted descending) are treated as "new"
    return index < 2;
  };

  const scrollAnimationDuration = props.scrollSpeed === 'slow' ? '60s' : props.scrollSpeed === 'medium' ? '30s' : '0s';
  const scrollDirectionValue = props.scrollDirection === 'right' ? 'reverse' : 'normal';

  if (loading) {
    return (
      <section className={`${styles.recognitionFeed} ${props.isDarkTheme ? styles.dark : ''}`}>
        <div className={styles.loadingState}>
          <div className={styles.spinner} />
          <span>Loading recognition feed...</span>
        </div>
      </section>
    );
  }

  if (!fieldsConfigured) {
    return (
      <section className={`${styles.recognitionFeed} ${props.isDarkTheme ? styles.dark : ''}`}>
        <div className={styles.errorState}>
          This web part needs to be configured. Open the edit panel and set the Data Source and Field Mapping options.
        </div>
      </section>
    );
  }

  if (error) {
    return (
      <section className={`${styles.recognitionFeed} ${props.isDarkTheme ? styles.dark : ''}`}>
        <div className={styles.errorState}>{error}</div>
      </section>
    );
  }

  if (winners.length === 0) {
    return (
      <section className={`${styles.recognitionFeed} ${props.isDarkTheme ? styles.dark : ''}`}>
        <div className={styles.emptyState}>No recognition data found yet.</div>
      </section>
    );
  }

  // Duplicate the list once so the CSS animation can loop seamlessly.
  const loopedWinners = [...winners, ...winners];

  return (
    <section className={`${styles.recognitionFeed} ${props.isDarkTheme ? styles.dark : ''}`}>
      <div className={styles.header}>
        <h2 className={styles.title}>Recent Recognition</h2>
        <button
          type="button"
          className={styles.pauseButton}
          onClick={() => setIsPaused(!isPaused)}
          aria-label={isPaused ? 'Play scrolling' : 'Pause scrolling'}
        >
          {isPaused ? '▶' : '❚❚'}
        </button>
      </div>

      <div className={styles.scrollViewport}>
        <div
          ref={scrollTrackRef}
          className={styles.scrollTrack}
          style={{
            animationPlayState: isPaused ? 'paused' : 'running',
            animationDuration: scrollAnimationDuration,
            animationDirection: scrollDirectionValue
          }}
          onMouseEnter={() => setIsPaused(true)}
          onMouseLeave={() => setIsPaused(false)}
        >
          {loopedWinners.map((winner, index) => (
            <button
              type="button"
              key={`${winner.Id}-${index}`}
              className={`${styles.card} ${isCurrentMonth(winner, index % winners.length) ? styles.newCard : ''}`}
              onClick={() => setSelectedWinner(winner)}
            >
              {winner.PhotoEmail && (
                <img
                  className={styles.cardPhoto}
                  src={getPhotoUrl(props.siteAbsoluteUrl, winner.PhotoEmail)}
                  alt=""
                />
              )}
              <div className={styles.cardBody}>
                <span className={styles.cardName}>
                  {winner.IsTeamAward && winner.TeamName ? winner.TeamName : winner.Name}
                </span>
                <span className={`${styles.cardBadge} ${winner.AwardType.toLowerCase().includes('critical') ? styles.badgeCriticalCog : styles.badgeEmployee}`}>
                  {winner.AwardType}
                </span>
                <span className={styles.cardMonth}>{winner.Month} {winner.Year}</span>
              </div>
              {isCurrentMonth(winner, index % winners.length) && (
                <span className={styles.newBadge}>New</span>
              )}
            </button>
          ))}

          {props.nominateUrl && (
            
              href={props.nominateUrl}
              target="_blank"
              rel="noreferrer"
              className={`${styles.card} ${styles.nominateCard}`}
            >
              <span className={styles.nominateText}>Know someone who deserves recognition?</span>
              <span className={styles.nominateCta}>Nominate Them →</span>
            </a>
          )}
        </div>
      </div>

      {selectedWinner && (
        <div className={styles.overlayBackdrop} onClick={() => setSelectedWinner(null)}>
          <div className={styles.overlayCard} onClick={(e) => e.stopPropagation()}>
            <button
              type="button"
              className={styles.overlayClose}
              onClick={() => setSelectedWinner(null)}
              aria-label="Close"
            >
              ×
            </button>

            {selectedWinner.PhotoEmail && (
              <img
                className={styles.overlayPhoto}
                src={getPhotoUrl(props.siteAbsoluteUrl, selectedWinner.PhotoEmail)}
                alt=""
              />
            )}

            <h3 className={styles.overlayName}>
              {selectedWinner.IsTeamAward && selectedWinner.TeamName ? selectedWinner.TeamName : selectedWinner.Name}
            </h3>

            <span className={`${styles.cardBadge} ${selectedWinner.AwardType.toLowerCase().includes('critical') ? styles.badgeCriticalCog : styles.badgeEmployee}`}>
              {selectedWinner.AwardType}
            </span>

            <p className={styles.overlayMonth}>{selectedWinner.Month} {selectedWinner.Year}</p>

            {selectedWinner.IsTeamAward && selectedWinner.TeamMembers && (
              <p className={styles.overlayTeamMembers}>
                <strong>Team Members:</strong> {selectedWinner.TeamMembers}
              </p>
            )}

            <p className={styles.overlayJustification}>{selectedWinner.Justification}</p>
          </div>
        </div>
      )}
    </section>
  );
};

export default RecognitionFeed;