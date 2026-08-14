import * as React from 'react';
import { useState, useEffect } from 'react';
import { Web } from '@pnp/sp/webs';
import { Icon } from '@fluentui/react/lib/Icon';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import type { IPersonaProps } from '@fluentui/react/lib/Persona';
import styles from './DocumentUploadRouter.module.scss';
import type { IDocumentUploadRouterProps } from './IDocumentUploadRouterProps';
import type { ISiteEntry } from './ISiteEntry';

const FALLBACK_ACCENT = '#BF9B30';
const FALLBACK_TILE_BG = '#1E1E1E';
const LAST_ROUTE_STORAGE_KEY = 'documentUploadRouter.lastRoute';
const CHUNK_THRESHOLD_BYTES = 10 * 1024 * 1024;

const SUPPORTED_FIELD_TYPES = ['Text', 'Note', 'Choice', 'MultiChoice', 'DateTime', 'Number', 'Boolean', 'User'];

const SYSTEM_FIELD_BLACKLIST = [
  'ContentType', 'Attachments', 'Edit', 'DocIcon', 'LinkTitle', 'LinkTitleNoMenu',
  'LinkFilenameNoMenu', 'LinkFilename', 'ItemChildCount', 'FolderChildCount',
  'AppAuthor', 'AppEditor', 'Author', 'Editor', '_ModerationComments', '_ModerationStatus',
  'WorkflowVersion', 'InstanceID', 'Order', 'GUID', 'FileLeafRef', 'FileRef', 'FileDirRef',
  '_UIVersionString', 'OData__UIVersionString', 'SMTotalSize', 'TemplateUrl',
  'xd_ProgID', 'xd_Signature'
];

type WizardStep = 'route' | 'library' | 'form' | 'review' | 'uploading' | 'success' | 'error';

interface IDynamicField {
  internalName: string;
  title: string;
  typeAsString: string;
  required: boolean;
  choices: string[];
}

interface IRawFieldInfo {
  InternalName: string;
  Title: string;
  TypeAsString: string;
  Required: boolean;
  Choices: string[] | null;
  Hidden: boolean;
  ReadOnlyField: boolean;
}

interface ILastRoute {
  siteUrl: string;
  libraryName: string;
}

interface IPickedPerson extends IPersonaProps {
  loginName?: string;
}

const DocumentUploadRouter: React.FunctionComponent<IDocumentUploadRouterProps> = (props) => {
  const [isPanelOpen, setIsPanelOpen] = useState<boolean>(false);
  const [step, setStep] = useState<WizardStep>('route');
  const [selectedSite, setSelectedSite] = useState<ISiteEntry | undefined>(undefined);
  const [selectedLibrary, setSelectedLibrary] = useState<string>('');
  const [availableLibraries, setAvailableLibraries] = useState<string[]>([]);
  const [libraryLoadState, setLibraryLoadState] = useState<'idle' | 'loading' | 'error'>('idle');
  const [lastRoute, setLastRoute] = useState<ILastRoute | undefined>(undefined);

  const [dynamicFields, setDynamicFields] = useState<IDynamicField[]>([]);
  const [fieldsLoadState, setFieldsLoadState] = useState<'idle' | 'loading' | 'error'>('idle');
  const [fieldValues, setFieldValues] = useState<Record<string, unknown>>({});
  const [selectedFile, setSelectedFile] = useState<File | undefined>(undefined);
  const [formError, setFormError] = useState<string>('');

  const [uploadProgress, setUploadProgress] = useState<number>(0);
  const [uploadedFileUrl, setUploadedFileUrl] = useState<string>('');
  const [uploadErrorMessage, setUploadErrorMessage] = useState<string>('');
  const [isPermissionError, setIsPermissionError] = useState<boolean>(false);

  useEffect(() => {
    try {
      const stored = window.localStorage.getItem(LAST_ROUTE_STORAGE_KEY);
      if (stored) {
        const parsed = JSON.parse(stored) as ILastRoute;
        if (parsed && parsed.siteUrl) {
          setLastRoute(parsed);
        }
      }
    } catch (err) {
      // localStorage unavailable or malformed, last-used memory simply won't be offered
    }
  }, []);

  const lastRouteEntry = lastRoute
    ? props.targetSites.filter((entry) => entry.url === lastRoute.siteUrl)[0]
    : undefined;

  const sortedTargetSites = React.useMemo(() => {
    const functional = props.targetSites.filter((e) => e.label.toUpperCase().indexOf('-PD-') === -1);
    const productDirectorate = props.targetSites.filter((e) => e.label.toUpperCase().indexOf('-PD-') !== -1);
    functional.sort((a, b) => a.label.localeCompare(b.label));
    productDirectorate.sort((a, b) => a.label.localeCompare(b.label));
    return [...functional, ...productDirectorate];
  }, [props.targetSites]);

  const resetWizard = (): void => {
    setStep('route');
    setSelectedSite(undefined);
    setSelectedLibrary('');
    setAvailableLibraries([]);
    setLibraryLoadState('idle');
    setDynamicFields([]);
    setFieldsLoadState('idle');
    setFieldValues({});
    setSelectedFile(undefined);
    setFormError('');
    setUploadProgress(0);
    setUploadedFileUrl('');
    setUploadErrorMessage('');
    setIsPermissionError(false);
  };

  const openPanel = (): void => {
    resetWizard();
    setIsPanelOpen(true);
  };

  const closePanel = (): void => {
    setIsPanelOpen(false);
  };

  const rememberRoute = (site: ISiteEntry, libraryName: string): void => {
    const next: ILastRoute = { siteUrl: site.url, libraryName };
    try {
      window.localStorage.setItem(LAST_ROUTE_STORAGE_KEY, JSON.stringify(next));
    } catch (err) {
      // localStorage unavailable, safe to ignore
    }
    setLastRoute(next);
  };

  const goToLibraryStep = (site: ISiteEntry, preselectLibrary?: string): void => {
    setSelectedSite(site);
    setStep('library');
    setLibraryLoadState('loading');
    setAvailableLibraries([]);
    setSelectedLibrary('');

    const web = Web([props.sp.web, site.url]);
    const hiddenLibraries = site.hiddenLibraries || [];

    web.lists
      .select('Title', 'BaseTemplate', 'Hidden')
      .filter('BaseTemplate eq 101 and Hidden eq false')()
      .then((rawLists) => {
        const names = (rawLists as { Title: string }[])
          .map((raw) => raw.Title)
          .filter((name) => hiddenLibraries.indexOf(name) === -1)
          .sort((a, b) => a.localeCompare(b));
        setAvailableLibraries(names);
        if (preselectLibrary && names.indexOf(preselectLibrary) !== -1) {
          setSelectedLibrary(preselectLibrary);
        }
        setLibraryLoadState('idle');
      })
      .catch((err) => {
        console.error(`Failed to load libraries for ${site.url}`, err);
        setLibraryLoadState('error');
      });
  };

  const handleUseLastRoute = (): void => {
    if (lastRouteEntry && lastRoute) {
      goToLibraryStep(lastRouteEntry, lastRoute.libraryName);
    }
  };

  const handleSelectSite = (url: string): void => {
    const site = props.targetSites.filter((entry) => entry.url === url)[0];
    if (site) {
      goToLibraryStep(site);
    }
  };

  const loadFieldsForLibrary = (site: ISiteEntry, libraryName: string): void => {
    setFieldsLoadState('loading');
    setDynamicFields([]);
    setFieldValues({});

    const web = Web([props.sp.web, site.url]);
    web.lists.getByTitle(libraryName).fields
      .select('InternalName', 'Title', 'TypeAsString', 'Required', 'Choices', 'Hidden', 'ReadOnlyField')()
      .then((rawFields) => {
        const filtered = (rawFields as IRawFieldInfo[])
          .filter((f) => !f.Hidden && !f.ReadOnlyField)
          .filter((f) => SUPPORTED_FIELD_TYPES.indexOf(f.TypeAsString) !== -1)
          .filter((f) => SYSTEM_FIELD_BLACKLIST.indexOf(f.InternalName) === -1)
          .map((f) => ({
            internalName: f.InternalName,
            title: f.Title,
            typeAsString: f.TypeAsString,
            required: f.Required,
            choices: f.Choices || []
          }));

        setDynamicFields(filtered);

        const initialValues: Record<string, unknown> = {};
        filtered.forEach((f) => {
          if (f.typeAsString === 'MultiChoice') {
            initialValues[f.internalName] = [];
          } else if (f.typeAsString === 'User') {
            initialValues[f.internalName] = [];
          } else {
            initialValues[f.internalName] = '';
          }
        });
        setFieldValues(initialValues);
        setFieldsLoadState('idle');
      })
      .catch((err) => {
        console.error(`Failed to load fields for ${libraryName}`, err);
        setFieldsLoadState('error');
      });
  };

  const handleContinueToForm = (): void => {
    if (selectedSite) {
      rememberRoute(selectedSite, selectedLibrary);
      loadFieldsForLibrary(selectedSite, selectedLibrary);
    }
    setStep('form');
  };

  const updateFieldValue = (internalName: string, value: unknown): void => {
    setFieldValues((prev) => ({ ...prev, [internalName]: value }));
  };

  const toggleMultiChoiceValue = (internalName: string, choice: string): void => {
    setFieldValues((prev) => {
      const current = (prev[internalName] as string[]) || [];
      const next = current.indexOf(choice) !== -1
        ? current.filter((c) => c !== choice)
        : [...current, choice];
      return { ...prev, [internalName]: next };
    });
  };

  const validateForm = (): boolean => {
    if (!selectedFile) {
      setFormError('Choose a file before continuing.');
      return false;
    }
    for (const field of dynamicFields) {
      if (!field.required) {
        continue;
      }
      const value = fieldValues[field.internalName];
      if (field.typeAsString === 'MultiChoice' || field.typeAsString === 'User') {
        if (!Array.isArray(value) || value.length === 0) {
          setFormError(`"${field.title}" is required.`);
          return false;
        }
      } else if (value === undefined || value === null || String(value).trim().length === 0) {
        setFormError(`"${field.title}" is required.`);
        return false;
      }
    }
    setFormError('');
    return true;
  };

  const handleReviewClick = (): void => {
    if (validateForm()) {
      setStep('review');
    }
  };

  const doUpload = async (): Promise<void> => {
    if (!selectedSite || !selectedFile) {
      return;
    }

    setStep('uploading');
    setUploadProgress(0);
    setUploadErrorMessage('');
    setIsPermissionError(false);

    try {
      const web = Web([props.sp.web, selectedSite.url]);
      const list = web.lists.getByTitle(selectedLibrary);

      let fileServerRelativeUrl = '';

      if (selectedFile.size > CHUNK_THRESHOLD_BYTES) {
        const chunkResult = await list.rootFolder.files.addChunked(
          selectedFile.name,
          selectedFile,
          (data: any) => {
            if (data && typeof data.blockNumber === 'number' && typeof data.totalBlocks === 'number' && data.totalBlocks > 0) {
              setUploadProgress(Math.round((data.blockNumber / data.totalBlocks) * 100));
            }
          }
        );
        fileServerRelativeUrl = chunkResult.ServerRelativeUrl;
      } else {
        const addResult = await list.rootFolder.files.addUsingPath(selectedFile.name, selectedFile, { Overwrite: false });
        setUploadProgress(100);
        fileServerRelativeUrl = addResult.ServerRelativeUrl;
      }

      const item = await web.getFileByServerRelativePath(fileServerRelativeUrl).getItem();

      const updatePayload: Record<string, unknown> = {};

      for (const field of dynamicFields) {
        const value = fieldValues[field.internalName];

        if (field.typeAsString === 'User') {
          const people = (value as IPickedPerson[]) || [];
          if (people.length > 0 && people[0].loginName) {
            const ensured = await web.ensureUser(people[0].loginName);
            updatePayload[field.internalName + 'Id'] = ensured.Id;
          }
          continue;
        }

        if (field.typeAsString === 'MultiChoice') {
          updatePayload[field.internalName] = { results: value || [] };
          continue;
        }

        if (field.typeAsString === 'Boolean') {
          updatePayload[field.internalName] = value === 'Yes';
          continue;
        }

        if (value !== undefined && value !== null && String(value).trim().length > 0) {
          updatePayload[field.internalName] = value;
        }
      }

      await item.update(updatePayload);

      setUploadedFileUrl(fileServerRelativeUrl);
      setStep('success');
    } catch (err) {
      console.error('Upload failed', err);
      const message = err && (err as Error).message ? (err as Error).message : '';
      const looksLikePermissionError = message.indexOf('403') !== -1
        || message.toLowerCase().indexOf('access denied') !== -1
        || message.toLowerCase().indexOf('forbidden') !== -1;

      setIsPermissionError(looksLikePermissionError);
      setUploadErrorMessage(
        looksLikePermissionError
          ? "You don't have access to upload here. Contact your KM team to request access."
          : 'Something went wrong uploading this file. You can try again.'
      );
      setStep('error');
    }
  };

  const themeColors = props.theme ? props.theme.semanticColors : undefined;
  const themePalette = props.theme ? props.theme.palette : undefined;
  const useThemeColors = props.useThemeColors;

  const accentColor = !useThemeColors && props.accentColorOverride && props.accentColorOverride.trim().length > 0
    ? props.accentColorOverride
    : (themePalette ? themePalette.themePrimary : FALLBACK_ACCENT);

  const tileBackgroundColor = !useThemeColors && props.tileBackgroundColorOverride && props.tileBackgroundColorOverride.trim().length > 0
    ? props.tileBackgroundColorOverride
    : (themeColors ? themeColors.bodyBackground : FALLBACK_TILE_BG);

  const panelBackgroundColor = !useThemeColors && props.panelBackgroundColorOverride && props.panelBackgroundColorOverride.trim().length > 0
    ? props.panelBackgroundColorOverride
    : (themeColors ? themeColors.bodyBackground : '#1E1E1E');

  const panelTextColor = !useThemeColors && props.panelTextColorOverride && props.panelTextColorOverride.trim().length > 0
    ? props.panelTextColorOverride
    : (themeColors ? themeColors.bodyText : '#F5F5F0');

  const panelBorderColor = !useThemeColors && props.panelBorderColorOverride && props.panelBorderColorOverride.trim().length > 0
    ? props.panelBorderColorOverride
    : (themeColors ? (themeColors.bodyDivider || '#3A3A3A') : '#3A3A3A');

  const cssVars: React.CSSProperties = {
    '--dur-text': panelTextColor,
    '--dur-text-secondary': themeColors ? (themeColors.bodySubtext || panelTextColor) : '#B8B8B0',
    '--dur-border': panelBorderColor,
    '--dur-accent-text': themePalette ? themePalette.white : '#1E1E1E',
    '--dur-bg-surface': panelBackgroundColor,
    '--dur-error-text': themeColors ? (themeColors.errorText || '#C86A6A') : '#C86A6A',
    '--dur-error-bg': themeColors ? (themeColors.errorBackground || 'rgba(163, 61, 61, 0.15)') : 'rgba(163, 61, 61, 0.15)',
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

  const renderDynamicField = (field: IDynamicField): JSX.Element => {
    const value = fieldValues[field.internalName];

    switch (field.typeAsString) {
      case 'Note':
        return (
          <textarea
            className={styles.textarea}
            value={(value as string) || ''}
            onChange={(e) => updateFieldValue(field.internalName, e.target.value)}
          />
        );
      case 'Choice':
        return (
          <select
            className={styles.select}
            value={(value as string) || ''}
            onChange={(e) => updateFieldValue(field.internalName, e.target.value)}
          >
            <option value="">Select an option</option>
            {field.choices.map((choice) => (
              <option key={choice} value={choice}>{choice}</option>
            ))}
          </select>
        );
      case 'MultiChoice':
        return (
          <div className={styles.checkboxGroup}>
            {field.choices.map((choice) => (
              <label key={choice} className={styles.checkboxRow}>
                <input
                  type="checkbox"
                  checked={((value as string[]) || []).indexOf(choice) !== -1}
                  onChange={() => toggleMultiChoiceValue(field.internalName, choice)}
                />
                <span>{choice}</span>
              </label>
            ))}
          </div>
        );
      case 'DateTime':
        return (
          <input
            type="date"
            className={styles.select}
            value={(value as string) || ''}
            onChange={(e) => updateFieldValue(field.internalName, e.target.value)}
          />
        );
      case 'Number':
        return (
          <input
            type="number"
            className={styles.select}
            value={(value as string) || ''}
            onChange={(e) => updateFieldValue(field.internalName, e.target.value)}
          />
        );
      case 'Boolean':
        return (
          <select
            className={styles.select}
            value={(value as string) || ''}
            onChange={(e) => updateFieldValue(field.internalName, e.target.value)}
          >
            <option value="">Select an option</option>
            <option value="Yes">Yes</option>
            <option value="No">No</option>
          </select>
        );
      case 'User':
        return (
          <PeoplePicker
            context={props.webPartContext as never}
            personSelectionLimit={1}
            showtooltip={false}
            principalTypes={[PrincipalType.User]}
            resolveDelay={300}
            defaultSelectedUsers={[]}
            onChange={(items: IPersonaProps[]) => updateFieldValue(field.internalName, items)}
          />
        );
      default:
        return (
          <input
            type="text"
            className={styles.select}
            value={(value as string) || ''}
            onChange={(e) => updateFieldValue(field.internalName, e.target.value)}
          />
        );
    }
  };

  const formatReviewValue = (field: IDynamicField): string => {
    const value = fieldValues[field.internalName];
    if (field.typeAsString === 'MultiChoice') {
      const arr = (value as string[]) || [];
      return arr.length > 0 ? arr.join(', ') : '(none)';
    }
    if (field.typeAsString === 'User') {
      const people = (value as IPickedPerson[]) || [];
      return people.length > 0 ? (people[0].text || people[0].loginName || '(selected)') : '(none)';
    }
    if (!value || String(value).trim().length === 0) {
      return '(none)';
    }
    return String(value);
  };

  return (
    <section className={styles.documentUploadRouter} style={cssVars}>
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
        onDismiss={step === 'uploading' ? undefined : closePanel}
        type={PanelType.medium}
        headerText={displayTitle}
        closeButtonAriaLabel="Close"
        isBlocking={step === 'uploading'}
      >
        <div className={styles.panelInner} style={cssVars}>
          {noRoutesConfigured && (
            <div className={styles.errorState}>
              No sites have been configured yet. Open the web part's edit panel and add at least one site under Target Sites.
            </div>
          )}

          {!noRoutesConfigured && step === 'route' && (
            <div className={styles.stepContent}>
              {lastRouteEntry && lastRoute && (
                <div className={styles.lastRouteBox}>
                  <span className={styles.lastRouteText}>
                    Last used: {lastRouteEntry.label}{lastRoute.libraryName ? ` \u203a ${lastRoute.libraryName}` : ''}
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
                {sortedTargetSites.map((entry) => (
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

              <label className={styles.fieldLabel} htmlFor="dur-file-input">File</label>
              <p className={styles.fieldHint}>Choose the document to upload.</p>
              <input
                id="dur-file-input"
                type="file"
                className={styles.fileInput}
                onChange={(e) => setSelectedFile(e.target.files && e.target.files[0] ? e.target.files[0] : undefined)}
              />
              {selectedFile && (
                <p className={styles.fileChosenText}>Selected: {selectedFile.name}</p>
              )}

              {fieldsLoadState === 'loading' && (
                <div className={styles.loadingState}>Loading form fields for this library...</div>
              )}
              {fieldsLoadState === 'error' && (
                <div className={styles.errorState}>
                  Couldn't load this library's columns. Try again, or contact your KM team.
                </div>
              )}

              {fieldsLoadState === 'idle' && dynamicFields.map((field) => (
                <div key={field.internalName} className={styles.formFieldBlock}>
                  <label className={styles.fieldLabel}>
                    {field.title}
                    {field.required && <span className={styles.requiredMark}> *</span>}
                  </label>
                  {renderDynamicField(field)}
                </div>
              ))}

              {formError && (
                <div className={styles.errorState}>{formError}</div>
              )}

              <div className={styles.stepActions}>
                <button type="button" className={styles.secondaryButton} onClick={() => setStep('library')}>
                  Back
                </button>
                <button type="button" className={styles.primaryButton} onClick={handleReviewClick}>
                  Review upload
                </button>
              </div>
            </div>
          )}

          {step === 'review' && selectedSite && (
            <div className={styles.stepContent}>
              <p className={styles.reviewHeading}>Check this before uploading</p>

              <table className={styles.reviewTable}>
                <tbody>
                  <tr>
                    <td className={styles.reviewLabel}>File</td>
                    <td>{selectedFile ? selectedFile.name : ''}</td>
                  </tr>
                  <tr>
                    <td className={styles.reviewLabel}>Going to</td>
                    <td>{selectedSite.label} &rsaquo; {selectedLibrary}</td>
                  </tr>
                  {dynamicFields.map((field) => (
                    <tr key={field.internalName}>
                      <td className={styles.reviewLabel}>{field.title}</td>
                      <td>{formatReviewValue(field)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>

              <div className={styles.stepActions}>
                <button type="button" className={styles.secondaryButton} onClick={() => setStep('form')}>
                  Back to edit
                </button>
                <button type="button" className={styles.primaryButton} onClick={() => { doUpload().catch((err) => console.error(err)); }}>
                  Confirm and upload
                </button>
              </div>
            </div>
          )}

          {step === 'uploading' && (
            <div className={styles.stepContent}>
              <p className={styles.fieldLabel}>Uploading, don't close this window</p>
              <div className={styles.progressTrack}>
                <div className={styles.progressFill} style={{ width: `${uploadProgress}%` }} />
              </div>
              <p className={styles.fieldHint}>{uploadProgress}%</p>
            </div>
          )}

          {step === 'success' && (
            <div className={styles.stepContent}>
              <div className={styles.successState}>
                <Icon iconName="CheckMark" className={styles.successIcon} />
                <p>File uploaded.</p>
                {uploadedFileUrl && (
                  <a href={uploadedFileUrl} target="_blank" rel="noreferrer" className={styles.changeLink}>
                    Open the file
                  </a>
                )}
              </div>
              <div className={styles.stepActions}>
                <button type="button" className={styles.primaryButton} onClick={openPanel}>
                  Upload another file
                </button>
              </div>
            </div>
          )}

          {step === 'error' && (
            <div className={styles.stepContent}>
              <div className={styles.errorState}>{uploadErrorMessage}</div>
              <div className={styles.stepActions}>
                <button type="button" className={styles.secondaryButton} onClick={() => setStep('form')}>
                  Back to edit
                </button>
                {!isPermissionError && (
                  <button type="button" className={styles.primaryButton} onClick={() => { doUpload().catch((err) => console.error(err)); }}>
                    Try again
                  </button>
                )}
              </div>
            </div>
          )}
        </div>
      </Panel>
    </section>
  );
};

export default DocumentUploadRouter;