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

const parseTargetSites = (raw: string): string[] => {
  return raw
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
};

const SiteContentIndex: React.FunctionComponent<ISiteContentIndexProps> = (props) => {
  const [results, setResults] = useState<ISiteResult[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [expandedSites, setExpandedSites] = useState<Set<string>>(new Set());

  const siteUrls = parseTargetSites(props.targetSites);

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
      setExpandedSites(new Set(siteResults.map((r) => r.siteUrl)));
      setLoading(false);
    };

    loadAllSites().catch((err) => console.error(err));
  }, [props.sp, props.targetSites, props.includeSystemLists]);

  const toggleSite = (siteUrl: string): void => {
    const next = new Set(expandedSites);
    if (next.has(siteUrl)) {
      next.delete(siteUrl);
    } else {
      next.add(siteUrl);
    }
    setExpandedSites(next);
  };

  const matchesSearch = (list: IListItem): boolean => {
    if (!searchTerm) {
      return true;
    }
    return list.Title.toLowerCase().indexOf(searchTerm.toLowerCase()) !== -1;
  };

  const formatDate = (raw: string): string => {
    if (!raw) {
      return '';
    }
    return new Date(raw).toLocaleDateString();
  };

  if (siteUrls.length === 0) {
    return (
      <section className={styles.siteContentIndex}>
        <div className={styles.errorState}>
          This web part needs to be configured. Open the edit panel and add one or more site URLs under Target Sites.
        </div>
      </section>
    );
  }

  if (loading) {
    return (
      <section className={styles.siteContentIndex}>
        <div className={styles.loadingState}>
          <div className={styles.spinner} />
          <span>Loading site content...</span>
        </div>
      </section>
    );
  }

  const renderListRow = (list: IListItem, showSite: boolean, siteTitle: string): JSX.Element => (
    <tr key={`${siteTitle}-${list.Id}`} className={styles.row}>
      {showSite && <td className={styles.cellSite}>{siteTitle}</td>}
      <td className={styles.cellName}>
        <a href={list.Url} target="_blank" rel="noreferrer">{list.Title}</a>
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
    site.lists.filter(matchesSearch).forEach((list) => {
      flatRows.push(renderListRow(list, true, site.siteTitle));
    });
  });

  return (
    <section className={styles.siteContentIndex}>
      <div className={styles.header}>
        <h2 className={styles.title}>Site Content Index</h2>
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
            const filteredLists = site.lists.filter(matchesSearch);
            const isExpanded = expandedSites.has(site.siteUrl);

            if (searchTerm && filteredLists.length === 0 && !site.error) {
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
                      <table className={styles.table}>
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
        <table className={styles.table}>
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