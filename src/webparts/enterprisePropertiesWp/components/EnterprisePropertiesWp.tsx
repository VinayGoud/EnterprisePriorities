import * as React from 'react';
import { useState, useEffect } from 'react';
import { Icon } from '@fluentui/react';
import styles from './EnterprisePropertiesWp.module.scss';
import { IEnterprisePropertiesWpProps } from './IEnterprisePropertiesWpProps';
import { EnterprisePrioritiesService, IPriorityCard } from './Services/Enterpriseprioritiesservice';

const getCardUrl = (link?: string): string => {
  if (!link) {
    return '#';
  }

  try {
    const url = new URL(link, window.location.origin);
    return url.protocol === 'http:' || url.protocol === 'https:' ? url.href : '#';
  } catch {
    return '#';
  }
};

// Cycles through 3 approved color variants regardless of how many cards exist
const getVariantClass = (index: number, styleMap: { [key: string]: string }): string => {
  const variants = [styleMap.variant0, styleMap.variant1, styleMap.variant2];
  return variants[index % variants.length];
};

const EnterprisePrioritiesWp: React.FC<IEnterprisePropertiesWpProps> = (props) => {
  const [cards, setCards] = useState<IPriorityCard[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);

  const service = React.useMemo(
    () => new EnterprisePrioritiesService(props.context),
    [props.context]
  );

  useEffect(() => {
    let isMounted = true;

    const loadCards = async (): Promise<void> => {
      const result = await service.getActivePriorityCards();
      if (isMounted) {
        setCards(result);
        setIsLoading(false);
      }
    };

    loadCards().catch((error) => {
      console.error('Failed to load priority cards:', error);
      if (isMounted) {
        setIsLoading(false);
      }
    });

    return () => {
      isMounted = false;
    };
  }, [service]);

  // ---- Render: loading ----
  if (isLoading) {
    return <div className={styles.enterprisePrioritiesWp}>Loading...</div>;
  }

  // ---- Render: empty state (per BRD error/empty states: "no active cards") ----
  if (cards.length === 0) {
    return (
      <div className={styles.enterprisePrioritiesWp}>
        <h2 className={styles.sectionTitle}>Enterprise Priorities</h2>
        <div className={styles.emptyState}>No enterprise priorities to display right now.</div>
      </div>
    );
  }

  // ---- Render: cards ----
  return (
    <div className={styles.enterprisePrioritiesWp}>
      <h2 className={styles.sectionTitle}>Enterprise Priorities</h2>
      <div className={styles.cardGrid}>
        {cards.map((card, index) => (
          <a
            key={card.Id}
            className={`${styles.priorityCard} ${getVariantClass(index, styles)}`}
            href={getCardUrl(card.Link)}
            target="_blank"
            rel="noopener noreferrer"
          >
            <span className={styles.cardIconBox}>
              {/*
                Icon comes directly from the SharePoint list's "Icon" column,
                expected to hold a Fluent UI icon name (e.g. "Compass", "Trophy", "Heart").
                Falls back to a generic icon if the field is empty or the name is invalid.
                NOTE: this assumes Fluent UI icon names as the storage format -- confirm
                this matches the "Icons/visuals" BRD decision once it's resolved; if editors
                will instead upload custom icon images, this will need to switch to an <img> tag.
              */}
         {card.Icon ? (
<img className={styles.cardIconImg} src={card.Icon} alt="" />
) : (
  <Icon iconName="Info" aria-hidden="true" />
)}
            </span>
            <span className={styles.cardArrow} aria-hidden="true">&#8594;</span>
            <h3 className={styles.cardTitle}>{card.Title}</h3>
            <p className={styles.cardDescription}>{card.Description}</p>
          </a>
        ))}
      </div>
    </div>
  );
};

export default EnterprisePrioritiesWp;
