import * as React from "react";
import { useEffect, useState } from "react";
import styles from "./EnterprisePropertiesWp.module.scss";
import { IEnterprisePropertiesWpProps } from "./IEnterprisePropertiesWpProps";
import {
  EnterprisePrioritiesService,
  IPriorityCard,
} from "./Services/Enterpriseprioritiesservice";

const EnterprisePropertiesWp: React.FC<IEnterprisePropertiesWpProps> = (
  props,
) => {
  const [cards, setCards] = useState<IPriorityCard[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);

  useEffect(() => {
    const loadPriorityCards = async (): Promise<void> => {
      try {
        const service = new EnterprisePrioritiesService(props.context);
        const priorityCards = await service.getActivePriorityCards();

        setCards(priorityCards);
      } catch (error) {
        console.error("Error loading Enterprise Priority cards:", error);
      } finally {
        setIsLoading(false);
      }
    };

    loadPriorityCards().catch((error) => {
      console.error("Error loading Enterprise Priority cards:", error);
    });
  }, [props.context]);

  const getBackgroundColor = (backgroundColor: string): string => {
    switch (backgroundColor) {
      case "Red":
        return "linear-gradient(160deg, #D21723 0%, #36090C 100%)";

      case "Dark Gray":
        return "linear-gradient(160deg, #202024 0%, #58585F 100%)";

      case "Dark Red":
        return "linear-gradient(160deg, #66070E 0%, #36090C 100%)";

      default:
        return "linear-gradient(160deg, #D21723 0%, #36090C 100%)";
    }
  };

  if (isLoading || cards.length === 0) {
    return null;
  }

  return (
    <section className={styles.enterprisePrioritiesWp}>
      <h2 className={styles.sectionTitle}>Enterprise Priorities</h2>

      <div className={styles.cardGrid}>
        {cards.map((card) => (
          <a
            key={card.Id}
            className={styles.priorityCard}
            style={{
              background: getBackgroundColor(card.BackgroundColor),
            }}
            href={card.URL}
            target="_blank"
            rel="noopener noreferrer"
            data-interception="off"
          >
            {card.IconURL && (
              <span className={styles.cardIconBox}>
                <img className={styles.cardIconImg} src={card.IconURL} alt="" />
              </span>
            )}

            {card.IsNew && <span className={styles.newBadge}>NEW</span>}

            <div className={styles.cardContent}>
              <h3 className={styles.cardTitle} title={card.Title}>
                {card.Title}
              </h3>

              <p className={styles.cardDescription} title={card.Description}>
                {card.Description}
              </p>
            </div>
          </a>
        ))}
      </div>
    </section>
  );
};

export default EnterprisePropertiesWp;
