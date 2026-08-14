import * as React from 'react';
import { useState } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import styles from './DocumentUploadRouter.module.scss';
import type { IDocumentUploadRouterProps } from './IDocumentUploadRouterProps';

const FALLBACK_ACCENT = '#BF9B30';
const FALLBACK_TILE_BG = '#1E1E1E';

const DocumentUploadRouter: React.FunctionComponent<IDocumentUploadRouterProps> = (props) => {
  const [isPanelOpen, setIsPanelOpen] = useState<boolean>(false);

  const themeColors = props.theme ? props.theme.semanticColors : undefined;
  const themePalette = props.theme ? props.theme.palette : undefined;

  const accentColor = props.accentColorOverride && props.accentColorOverride.trim().length > 0
    ? props.accentColorOverride
    : (themePalette ? themePalette.themePrimary : FALLBACK_ACCENT);

  const tileBackgroundColor = props.tileBackgroundColorOverride && props.tileBackgroundColorOverride.trim().length > 0
    ? props.tileBackgroundColorOverride
    : (themeColors ? themeColors.bodyBackground : FALLBACK_TILE_BG);

  const rootStyle: React.CSSProperties = {
    ...(themeColors && themePalette ? {
      '--dur-text': themeColors.bodyText,
      '--dur-text-secondary': themeColors.bodySubtext || themeColors.bodyText,
      '--dur-border': themeColors.bodyDivider || themePalette.neutralLight,
      '--dur-accent-text': themePalette.white
    } : {}),
    '--dur-accent': accentColor,
    '--dur-tile-bg': tileBackgroundColor
  } as React.CSSProperties;

  const displayTitle = props.customTitle && props.customTitle.trim().length > 0
    ? props.customTitle
    : 'Document Upload Router';

  const displayDescription = props.tileDescription && props.tileDescription.trim().length > 0
    ? props.tileDescription
    : 'Upload a document to the right library, tagged correctly, in a few steps.';

  const iconName = props.tileIconName && props.tileIconName.trim().length > 0
    ? props.tileIconName
    : 'CloudUpload';

  return (
    <section className={styles.documentUploadRouter} style={rootStyle}>
      <button
        type="button"
        className={styles.tile}
        onClick={() => setIsPanelOpen(true)}
      >
        <span className={styles.tileIcon}>
          <Icon iconName={iconName} />
        </span>
        <span className={styles.tileText}>
          {props.showTitle && <span className={styles.tileTitle}>{displayTitle}</span>}
          <span className={styles.tileDescription}>{displayDescription}</span>
        </span>
      </button>

      <Panel
        isOpen={isPanelOpen}
        onDismiss={() => setIsPanelOpen(false)}
        type={PanelType.medium}
        headerText={displayTitle}
        closeButtonAriaLabel="Close"
      >
        <div className={styles.panelPlaceholder}>
          <p>Step 1 of 3: choose where the document belongs.</p>
          <p className={styles.placeholderNote}>
            The site and library pickers, the dynamic metadata form, and the review step come next.
          </p>
        </div>
      </Panel>
    </section>
  );
};

export default DocumentUploadRouter;