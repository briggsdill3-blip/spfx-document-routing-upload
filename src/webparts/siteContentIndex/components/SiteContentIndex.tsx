import * as React from 'react';
import { useState, useEffect } from 'react';
import { Web } from '@pnp/sp/webs';
import styles from './SiteContentIndex.module.scss';
import type { ISiteContentIndexProps } from './ISiteContentIndexProps';

interface IListItem {
  Id: string;
  Title: string;
  Type: string;
  ItemCount: number;
  LastModified: string;
  HasUniqueRoleAssignments: boolean;
  Url: string;
}

interface IRawListInfo {
  Id: string;
  Title: string;
  BaseTemplate: number;
  ItemCount: number;
  LastItemModifiedDate: string;
  HasUniqueRoleAssignments: boolean;
  DefaultViewUrl: string;
  Hidden: boolean;
}

interface ISiteResult {
  siteUrl: string;
  siteTitle: string;
  lists: IListItem[];
  error: string;
}

const BASE_TEMPLATE_NAMES: Record<number, string> = {
  100: 'Custom List',
  101: 'Document Library',
  102: 'Survey',
  103: 'Links',
  104: 'Announcements',
  105: 'Contacts',
  106: 'Calendar',
  107: 'Tasks',
  109: 'Picture Library',
  118: 'Custom List (Datasheet)',
  120: 'Data Sources'
};

const getListTypeName = (baseTemplate: number): string => {
  return BASE_TEMPLATE_NAMES[baseTemplate] || 'List';
};

const MS_PER_DAY = 1000 * 60 * 60 * 24;
const FALLBACK_ACCENT = '#BF9B30';
const FALLBACK_STRIPE = '#2A2A2A';

const SiteContentIndex: React.FunctionComponent<ISiteContentIndexProps> = (props) => {
  const [results, setResults] = useState<ISiteResult[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [expandedSites, setExpandedSites] = useState<Set<string>>(new Set());

  const siteUrls = props.targetSites.map((entry) => entry.url);

  useEffect(() => {
    if (siteUrls.length === 0) {
      setLoading(false);
      return;
    }

    const loadAllSites = async (): Promise<void> => {
      setLoading(true);

      const siteResults: ISiteResult[] = await Promise.all(
        siteUrls.map(async (siteUrl): Promise<ISiteResult> => {
          try {
            const web = Web([props.sp.web, siteUrl]);

            const webInfo = await web.select('Title')();

            let listsQuery = web.lists.select(
              'Id',
              'Title',
              'BaseTemplate',
              'ItemCount',
              'LastItemModifiedDate',
              'HasUniqueRoleAssignments',
              'DefaultViewUrl',
              'Hidden'
            );

            if (!props.includeSystemLists) {
              listsQuery = listsQuery.filter('Hidden eq false');
            }

            const rawLists = (await listsQuery()) as unknown as IRawListInfo[];

            const mapped: IListItem[] = rawLists
              .map((raw: IRawListInfo) => ({
                Id: raw.Id,
                Title: raw.Title,
                Type: getListTypeName(raw.BaseTemplate),
                ItemCount: raw.ItemCount,
                LastModified: raw.LastItemModifiedDate,
                HasUniqueRoleAssignments: raw.HasUniqueRoleAssignments,
                Url: `${new URL(siteUrl).origin}${raw.DefaultViewUrl}`
              }))
              .sort((a, b) => a.Title.localeCompare(b.Title));

            return {
              siteUrl,
              siteTitle: webInfo.Title || siteUrl,
              lists: mapped,
              error: ''
            };
          } catch (err) {
            console.error(`Failed to load lists for ${siteUrl}`, err);
            return {
              siteUrl,
              siteTitle: siteUrl,
              lists: [],
              error: 'Unable to load this site. Check the URL and confirm you have access.'
            };
          }
        })
      );

      setResults(siteResults);
      setExpandedSites(props.expandByDefault ? new Set(siteResults.map((r) => r.siteUrl)) : new Set());
      setLoading(false);
    };

    loadAllSites().catch((err) => console.error(err));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.sp, props.targetSites, props.includeSystemLists, props.expandByDefault]);

  const toggleSite = (siteUrl: string): void => {
    const next = new Set(expandedSites);
    if (next.has(siteUrl)) {
      next.delete(siteUrl);
    } else {
      next.add(siteUrl);
    }
    setExpandedSites(next);
  };

  const matchesFilters = (list: IListItem): boolean => {
    if (searchTerm && list.Title.toLowerCase().indexOf(searchTerm.toLowerCase()) === -1) {
      return false;
    }
    if (props.permissionsFilter === 'unique' && !list.HasUniqueRoleAssignments) {
      return false;
    }
    if (props.permissionsFilter === 'inherited' && list.HasUniqueRoleAssignments) {
      return false;
    }
    return true;
  };

  const isStale = (list: IListItem): boolean => {
    if (!props.staleDaysThreshold || props.staleDaysThreshold <= 0) {
      return false;
    }
    if (!list.LastModified) {
      return false;
    }
    const ageInDays = (Date.now() - new Date(list.LastModified).getTime()) / MS_PER_DAY;
    return ageInDays > props.staleDaysThreshold;
  };

  const formatDate = (raw: string): string => {
    if (!raw) {
      return '';
    }
    return new Date(raw).toLocaleDateString();
  };

  const themeColors = props.theme ? props.theme.semanticColors : undefined;
  const themePalette = props.theme ? props.theme.palette : undefined;

  const accentColor = props.accentColorOverride && props.accentColorOverride.trim().length > 0
    ? props.accentColorOverride
    : (themePalette ? themePalette.themePrimary : FALLBACK_ACCENT);

  const stripeColor = props.stripeColorOverride && props.stripeColorOverride.trim().length > 0
    ? props.stripeColorOverride
    : (themePalette ? (themePalette.neutralLighterAlt || themePalette.neutralLight) : FALLBACK_STRIPE);

  const rootStyle: React.CSSProperties = {
    ...(themeColors && themePalette ? {
      '--scix-text': themeColors.bodyText,
      '--scix-text-secondary': themeColors.bodySubtext || themeColors.bodyText,
      '--scix-bg-surface': themeColors.bodyBackground,
      '--scix-bg-hover': themeColors.bodyBackgroundHovered || themeColors.bodyBackground,
      '--scix-border': themeColors.bodyDivider || themePalette.neutralLight,
      '--scix-accent-text': themePalette.white,
      '--scix-error-bg': themeColors.errorBackground || 'rgba(163, 61, 61, 0.15)',
      '--scix-error-text': themeColors.errorText || '#C86A6A'
    } : {}),
    '--scix-accent': accentColor,
    '--scix-stripe-bg': stripeColor
  } as React.CSSProperties;

  const tableClass = `${styles.table} ${props.enableRowStriping ? styles.striped : ''}`;

  const renderListRow = (list: IListItem, showSite: boolean, siteTitle: string): JSX.Element => (
    <tr key={`${siteTitle}-${list.Id}`} className={styles.row}>
      {showSite && <td className={styles.cellSite}>{siteTitle}</td>}
      <td className={styles.cellName}>
        <a href={list.Url} target="_blank" rel="noreferrer">{list.Title}</a>
        {isStale(list) && <span className={styles.staleFlag}>Stale</span>}
      </td>
      <td className={styles.cellType}>{list.Type}</td>
      <td className={styles.cellCount}>{list.ItemCount}</td>
      <td className={styles.cellDate}>{formatDate(list.LastModified)}</td>
      <td className={styles.cellPerms}>
        {list.HasUniqueRoleAssignments ? (
          <span className={styles.uniquePermsFlag}>Unique</span>
        ) : (
          <span className={styles.inheritedPermsFlag}>Inherited</span>
        )}
      </td>
    </tr>
  );

  const flatRows: JSX.Element[] = [];
  results.forEach((site) => {
    site.lists.filter(matchesFilters).forEach((list) => {
      flatRows.push(renderListRow(list, true, site.siteTitle));
    });
  });

  if (siteUrls.length === 0) {
    return (
      <section className={styles.siteContentIndex} style={rootStyle}>
        <div className={styles.errorState}>
          This web part needs to be configured. Open the edit panel and add one or more site URLs under Target Sites.
        </div>
      </section>
    );
  }

  if (loading) {
    return (
      <section className={styles.siteContentIndex} style={rootStyle}>
        <div className={styles.loadingState}>
          <div className={styles.spinner} />
          <span>Loading site content...</span>
        </div>
      </section>
    );
  }

  const displayTitle = props.customTitle && props.customTitle.trim().length > 0
    ? props.customTitle
    : 'Site Content Index';

  return (
    <section className={styles.siteContentIndex} style={rootStyle}>
      <div className={styles.header}>
        {props.showTitle && <h2 className={styles.title}>{displayTitle}</h2>}
        <input
          type="text"
          className={styles.searchBox}
          placeholder="Search lists and libraries..."
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
        />
      </div>

      {props.groupBySite ? (
        <div className={styles.groupedView}>
          {results.map((site) => {
            const filteredLists = site.lists.filter(matchesFilters);
            const isExpanded = expandedSites.has(site.siteUrl);

            if ((searchTerm || props.permissionsFilter !== 'all') && filteredLists.length === 0 && !site.error) {
              return null;
            }

            return (
              <div key={site.siteUrl} className={styles.siteGroup}>
                <button
                  type="button"
                  className={styles.siteGroupHeader}
                  onClick={() => toggleSite(site.siteUrl)}
                  aria-expanded={isExpanded}
                >
                  <span className={styles.expandIcon}>{isExpanded ? '▾' : '▸'}</span>
                  <span className={styles.siteGroupTitle}>{site.siteTitle}</span>
                  <span className={styles.siteGroupCount}>{filteredLists.length}</span>
                </button>

                {isExpanded && (
                  <>
                    {site.error ? (
                      <div className={styles.siteError}>{site.error}</div>
                    ) : (
                      <table className={tableClass}>
                        <thead>
                          <tr>
                            <th>Name</th>
                            <th>Type</th>
                            <th>Items</th>
                            <th>Last Modified</th>
                            <th>Permissions</th>
                          </tr>
                        </thead>
                        <tbody>
                          {filteredLists.map((list) => renderListRow(list, false, site.siteTitle))}
                        </tbody>
                      </table>
                    )}
                  </>
                )}
              </div>
            );
          })}
        </div>
      ) : (
        <table className={tableClass}>
          <thead>
            <tr>
              <th>Site</th>
              <th>Name</th>
              <th>Type</th>
              <th>Items</th>
              <th>Last Modified</th>
              <th>Permissions</th>
            </tr>
          </thead>
          <tbody>
            {flatRows}
          </tbody>
        </table>
      )}
    </section>
  );
};

export default SiteContentIndex;