import * as React from 'react';
import { useState } from 'react';
import { Web } from '@pnp/sp/webs';
import { SPFI } from '@pnp/sp';
import styles from './SiteEntryChipInput.module.scss';
import { ISiteEntry, deriveLabel } from '../components/ISiteEntry';

export interface ISiteEntryChipInputProps {
  label: string;
  entries: ISiteEntry[];
  sp: SPFI;
  accentColor?: string;
  onChange: (entries: ISiteEntry[]) => void;
}

interface ILibraryOption {
  name: string;
}

type LibraryState = ILibraryOption[] | 'loading' | 'error';

const SiteEntryChipInput: React.FunctionComponent<ISiteEntryChipInputProps> = (props) => {
  const [inputValue, setInputValue] = useState<string>('');
  const [expandedUrl, setExpandedUrl] = useState<string>('');
  const [librariesByUrl, setLibrariesByUrl] = useState<Record<string, LibraryState>>({});

  const wrapperStyle: React.CSSProperties = props.accentColor
    ? ({ '--scix-pane-accent': props.accentColor } as React.CSSProperties)
    : {};

  const addEntry = (): void => {
    const trimmed = inputValue.trim();
    if (trimmed.length === 0) {
      return;
    }

    const alreadyExists = props.entries.some(
      (entry) => entry.url.toLowerCase() === trimmed.toLowerCase()
    );
    if (alreadyExists) {
      setInputValue('');
      return;
    }

    const newEntry: ISiteEntry = {
      url: trimmed,
      label: deriveLabel(trimmed),
      hiddenLibraries: []
    };

    props.onChange([...props.entries, newEntry]);
    setInputValue('');
  };

  const removeEntry = (url: string): void => {
    props.onChange(props.entries.filter((entry) => entry.url !== url));
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'Enter') {
      e.preventDefault();
      addEntry();
    }
  };

  const loadLibraries = async (url: string): Promise<void> => {
    setLibrariesByUrl((prev) => ({ ...prev, [url]: 'loading' }));
    try {
      const web = Web([props.sp.web, url]);
      const rawLists = await web.lists
        .select('Title', 'BaseTemplate', 'Hidden')
        .filter('BaseTemplate eq 101 and Hidden eq false')();

      const options: ILibraryOption[] = (rawLists as { Title: string }[])
        .map((raw) => ({ name: raw.Title }))
        .sort((a, b) => a.name.localeCompare(b.name));

      setLibrariesByUrl((prev) => ({ ...prev, [url]: options }));
    } catch (err) {
      console.error(`Failed to load libraries for ${url}`, err);
      setLibrariesByUrl((prev) => ({ ...prev, [url]: 'error' }));
    }
  };

  const toggleManageLibraries = (url: string): void => {
    if (expandedUrl === url) {
      setExpandedUrl('');
      return;
    }
    setExpandedUrl(url);
    if (!librariesByUrl[url]) {
      loadLibraries(url).catch((err) => console.error(err));
    }
  };

  const toggleLibraryHidden = (entryUrl: string, libraryName: string): void => {
    const updated = props.entries.map((entry) => {
      if (entry.url !== entryUrl) {
        return entry;
      }
      const currentHidden = entry.hiddenLibraries || [];
      const isHidden = currentHidden.indexOf(libraryName) !== -1;
      const nextHidden = isHidden
        ? currentHidden.filter((name) => name !== libraryName)
        : [...currentHidden, libraryName];
      return { ...entry, hiddenLibraries: nextHidden };
    });
    props.onChange(updated);
  };

  return (
    <div className={styles.chipInputWrapper} style={wrapperStyle}>
      <div className={styles.fieldLabel}>{props.label}</div>

      <div className={styles.inputRow}>
        <input
          type="text"
          className={styles.textInput}
          placeholder="Paste a site URL and press Enter"
          value={inputValue}
          onChange={(e) => setInputValue(e.target.value)}
          onKeyDown={handleKeyDown}
        />
        <button
          type="button"
          className={styles.addButton}
          onClick={addEntry}
          disabled={inputValue.trim().length === 0}
        >
          Add
        </button>
      </div>

      {props.entries.length > 0 && (
        <ul className={styles.chipList}>
          {props.entries.map((entry) => {
            const isExpanded = expandedUrl === entry.url;
            const libState = librariesByUrl[entry.url];
            const hiddenLibraries = entry.hiddenLibraries || [];
            const hiddenCount = hiddenLibraries.length;

            return (
              <li key={entry.url} className={styles.chipBlock}>
                <div className={styles.chip} title={entry.url}>
                  <span className={styles.chipLabel}>{entry.label}</span>
                  <button
                    type="button"
                    className={styles.chipManageButton}
                    onClick={() => toggleManageLibraries(entry.url)}
                  >
                    {hiddenCount > 0 ? `Libraries (${hiddenCount} hidden)` : 'Manage libraries'}
                  </button>
                  <button
                    type="button"
                    className={styles.chipRemove}
                    onClick={() => removeEntry(entry.url)}
                    aria-label={`Remove ${entry.label}`}
                  >
                    ×
                  </button>
                </div>

                {isExpanded && (
                  <div className={styles.libraryPanel}>
                    {libState === 'loading' && (
                      <div className={styles.libraryStatus}>Loading libraries...</div>
                    )}
                    {libState === 'error' && (
                      <div className={styles.libraryStatus}>
                        Couldn't load libraries for this site. Check the URL and your access.
                      </div>
                    )}
                    {libState && libState !== 'loading' && libState !== 'error' && libState.length === 0 && (
                      <div className={styles.libraryStatus}>No document libraries found on this site.</div>
                    )}
                    {libState && libState !== 'loading' && libState !== 'error' && libState.length > 0 && (
                      <>
                        <p className={styles.libraryHint}>
                          Uncheck any library you don't want available to end users.
                        </p>
                        <div className={styles.libraryScroll}>
                          {libState.map((lib) => (
                            <label key={lib.name} className={styles.libraryRow}>
                              <input
                                type="checkbox"
                                checked={hiddenLibraries.indexOf(lib.name) === -1}
                                onChange={() => toggleLibraryHidden(entry.url, lib.name)}
                              />
                              <span>{lib.name}</span>
                            </label>
                          ))}
                        </div>
                      </>
                    )}
                  </div>
                )}
              </li>
            );
          })}
        </ul>
      )}
    </div>
  );
};

export default SiteEntryChipInput;