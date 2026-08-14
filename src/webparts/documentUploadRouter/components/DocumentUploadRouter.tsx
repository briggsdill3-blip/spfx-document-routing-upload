import * as React from 'react';
import { useState, useEffect } from 'react';
import { Web } from '@pnp/sp/webs';
import { Icon } from '@fluentui/react/lib/Icon';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import styles from './DocumentUploadRouter.module.scss';
import type { IDocumentUploadRouterProps } from './IDocumentUploadRouterProps';
import type { ISiteEntry } from './ISiteEntry';

const FALLBACK_ACCENT = '#BF9B30';
const FALLBACK_TILE_BG = '#1E1E1E';
const LAST_ROUTE_STORAGE_KEY = 'documentUploadRouter.lastRouteUrl';

type WizardStep = 'route' | 'library' | 'form';

const DocumentUploadRouter: React.FunctionComponent<IDocumentUploadRouterProps> = (props) => {
  const [isPanelOpen, setIsPanelOpen] = useState<boolean>(false);
  const [step, setStep] = useState<WizardStep>('route');
  const [selectedSite, setSelectedSite] = useState<ISiteEntry | undefined>(undefined);
  const [selectedLibrary, setSelectedLibrary] = useState<string>('');
  const [availableLibraries, setAvailableLibraries] = useState<string[]>([]);
  const [libraryLoadState, setLibraryLoadState] = useState<'idle' | 'loading' | 'error'>('idle');
  const [lastRouteUrl, setLastRouteUrl] = useState<string>('');

  useEffect(() => {
    try {
      const stored = window.localStorage.getItem(LAST_ROUTE_STORAGE_KEY);
      if (stored) {
        setLastRouteUrl(stored);
      }
    } catch (err) {
      // localStorage unavailable, last-used memory simply won't be offered
    }
  }, []);

  const lastRouteEntry = props.targetSites.filter((entry) => entry.url === lastRouteUrl)[0];

  const resetWizard = (): void => {
    setStep('route');
    setSelectedSite(undefined);
    setSelectedLibrary('');
    setAvailableLibraries([]);
    setLibraryLoadState('idle');
  };

  const openPanel = (): void => {
    resetWizard();
    setIsPanelOpen(true);
  };

  const closePanel = (): void => {
    setIsPanelOpen(false);
  };

  const rememberRoute = (site: ISiteEntry): void => {
    try {
      window.localStorage.setItem(LAST_ROUTE_STORAGE_KEY, site.url);
    } catch (err) {
      // localStorage unavailable, safe to ignore
    }
    setLastRouteUrl(site.url);
  };

  const goToLibraryStep = (site: ISiteEntry): void => {
    setSelectedSite(site);
    setStep('library');
    setLibraryLoadState('loading');
    setAvailableLibraries([]);

    const web = Web([props.sp.web, site.url]);
    web.lists
      .select('Title', 'BaseTemplate', 'Hidden')
      .filter('BaseTemplate eq 101 and Hidden eq false')()
      .then((rawLists) => {
        const names = (rawLists as { Title: string }[])
          .map((raw) => raw.Title)
          .filter((name) => site.hiddenLibraries.indexOf(name) === -1)
          .sort((a, b) => a.localeCompare(b));
        setAvailableLibraries(names);
        setLibraryLoadState('idle');
      })
      .catch((err) => {
        console.error(`Failed to load libraries for ${site.url}`, err);
        setLibraryLoadState('error');
      });
  };

  const handleUseLastRoute = (): void => {
    if (lastRouteEntry) {
      goToLibraryStep(lastRouteEntry);
    }
  };

  const handleSelectSite = (url: string): void => {
    const site = props.targetSites.filter((entry) => entry.url === url)[0];
    if (site) {
      rememberRoute(site);
      goToLibraryStep(site);
    }
  };

  const handleContinueToForm = (): void => {
    if (selectedSite) {
      rememberRoute(selectedSite);
    }
    setStep('form');
  };

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
      '--dur-accent-text': themePalette.white,
      '--dur-bg-surface': themeColors.bodyBackground,
      '--dur-error-text': themeColors.errorText || '#C86A6A',
      '--dur-error-bg': themeColors.errorBackground || 'rgba(163, 61, 61, 0.15)'
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

  const noRoutesConfigured = props.targetSites.length === 0;

  return (
    <section className={styles.documentUploadRouter} style={rootStyle}>
      <button
        type="button"
        className={styles.floatingButton}
        onClick={openPanel}
        title={displayDescription}
      >
        <span className={styles.floatingIcon}>
          <Icon iconName={iconName} />
        </span>
        <span className={styles.floatingLabel}>
          {props.showTitle && displayTitle}
        </span>
      </button>

      <Panel
        isOpen={isPanelOpen}
        onDismiss={closePanel}
        type={PanelType.medium}
        headerText={displayTitle}
        closeButtonAriaLabel="Close"
      >
        {noRoutesConfigured && (
          <div className={styles.errorState}>
            No sites have been configured yet. Open the web part's edit panel and add at least one site under Target Sites.
          </div>
        )}

        {!noRoutesConfigured && step === 'route' && (
          <div className={styles.stepContent}>
            {lastRouteEntry && (
              <div className={styles.lastRouteBox}>
                <span className={styles.lastRouteText}>
                  Last used: {lastRouteEntry.label}
                </span>
                <button type="button" className={styles.primaryButton} onClick={handleUseLastRoute}>
                  Use this again
                </button>
              </div>
            )}

            <label className={styles.fieldLabel} htmlFor="dur-site-select">
              Where does the library live?
            </label>
            <p className={styles.fieldHint}>Pick the site or product office that owns the document.</p>
            <select
              id="dur-site-select"
              className={styles.select}
              value=""
              onChange={(e) => handleSelectSite(e.target.value)}
            >
              <option value="" disabled>Select a site</option>
              {props.targetSites.map((entry) => (
                <option key={entry.url} value={entry.url}>{entry.label}</option>
              ))}
            </select>
          </div>
        )}

        {step === 'library' && selectedSite && (
          <div className={styles.stepContent}>
            <div className={styles.breadcrumb}>
              <span>{selectedSite.label}</span>
              <button type="button" className={styles.changeLink} onClick={() => setStep('route')}>
                Change
              </button>
            </div>

            <label className={styles.fieldLabel} htmlFor="dur-library-select">
              Which document library?
            </label>
            <p className={styles.fieldHint}>These are the libraries you can upload to on that site.</p>

            {libraryLoadState === 'loading' && (
              <div className={styles.loadingState}>Loading libraries...</div>
            )}
            {libraryLoadState === 'error' && (
              <div className={styles.errorState}>
                Couldn't load libraries for this site. Check the URL and confirm you have access, or contact your KM team.
              </div>
            )}
            {libraryLoadState === 'idle' && availableLibraries.length === 0 && (
              <div className={styles.errorState}>
                No document libraries are available on this site.
              </div>
            )}
            {libraryLoadState === 'idle' && availableLibraries.length > 0 && (
              <>
                <select
                  id="dur-library-select"
                  className={styles.select}
                  value={selectedLibrary}
                  onChange={(e) => setSelectedLibrary(e.target.value)}
                >
                  <option value="">Select a library</option>
                  {availableLibraries.map((name) => (
                    <option key={name} value={name}>{name}</option>
                  ))}
                </select>

                <div className={styles.stepActions}>
                  <button type="button" className={styles.secondaryButton} onClick={() => setStep('route')}>
                    Back
                  </button>
                  <button
                    type="button"
                    className={styles.primaryButton}
                    onClick={handleContinueToForm}
                    disabled={selectedLibrary.trim().length === 0}
                  >
                    Continue
                  </button>
                </div>
              </>
            )}
          </div>
        )}

        {step === 'form' && selectedSite && (
          <div className={styles.stepContent}>
            <div className={styles.breadcrumb}>
              <span>{selectedSite.label} &rsaquo; {selectedLibrary}</span>
              <button type="button" className={styles.changeLink} onClick={() => setStep('library')}>
                Change
              </button>
            </div>
            <p>The file picker and metadata form for this library come next.</p>
          </div>
        )}
      </Panel>
    </section>
  );
};

export default DocumentUploadRouter;