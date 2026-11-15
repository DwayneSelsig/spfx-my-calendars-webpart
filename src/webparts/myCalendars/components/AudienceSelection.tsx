import * as React from 'react';
import {
  Button, DataGrid, DataGridBody, DataGridCell, DataGridHeader, DataGridHeaderCell,
  DataGridRow, Field, FluentProvider, Radio, RadioGroup, SearchBox, Spinner,
  Text, createTableColumn, tokens, webDarkTheme, webLightTheme
} from '@fluentui/react-components';
import { useTheme } from '@fluentui/react';
import * as strings from 'MyCalendarsWebPartStrings';
import type { AudienceService, IEntraAudienceGroup } from '../services/AudienceService';
import { audienceTypeLabel } from './AssignmentControls';
import { formatLocalizedString } from '../utils/localization';

export interface IAudienceSelectionProps {
  groups: IEntraAudienceGroup[];
  loading: boolean;
  loaded?: boolean;
  error?: string;
  search: string;
  everyone: boolean;
  single: boolean;
  selected: Record<string, string>;
  service?: AudienceService;
  actionLabel: string;
  onSearchChange: (value: string) => void;
  onEveryoneChange: (everyone: boolean) => void;
  onSelectionChange: (selected: Record<string, string>) => void;
  onApply: () => void;
}

const pageSize = 10;

/** Presentation and ephemeral count state; the administrator panel retains ownership of the draft. */
export function AudienceSelection(props: IAudienceSelectionProps): React.ReactElement {
  const theme = useTheme();
  const [page, setPage] = React.useState(0);
  const [counts, setCounts] = React.useState<Record<string, number | undefined>>({});
  const cache = React.useRef<Record<string, number>>({});
  const query = props.search.trim().toLocaleLowerCase();
  const filteredGroups = props.groups.filter(group => group.displayName.toLocaleLowerCase().indexOf(query) !== -1);
  const pageCount = Math.max(1, Math.ceil(filteredGroups.length / pageSize));
  const currentPage = Math.min(page, pageCount - 1);
  const visible = filteredGroups.slice(currentPage * pageSize, (currentPage + 1) * pageSize);
  const visibleIds = visible.map(group => group.id).join(',');

  React.useEffect(() => { setPage(0); }, [props.groups, props.search]);
  React.useEffect(() => {
    let active = true;
    if (props.everyone || props.loading || !visibleIds) return;
    const ids = visibleIds.split(',');
    const missing = ids.filter(id => cache.current[id] === undefined);
    const unavailable: Record<string, undefined> = {};
    missing.forEach(id => { unavailable[id] = undefined; });
    setCounts({ ...cache.current });
    if (missing.length) {
      const request = props.service ? props.service.getAudienceMemberCounts(missing)
        : Promise.resolve(unavailable);
      request.then(result => {
        if (!active) return;
        Object.keys(result).forEach(id => { if (result[id] !== undefined) cache.current[id] = result[id] as number; });
        setCounts({ ...cache.current, ...result });
      }).catch(() => {
        if (active) setCounts(previous => ({ ...previous, ...unavailable }));
      });
    }
    return () => { active = false; };
  }, [visibleIds, props.everyone, props.loading, props.service]);

  const columns = [
    createTableColumn<IEntraAudienceGroup>({ columnId: 'name', renderHeaderCell: () => strings.NameLabel,
      renderCell: group => <Text wrap>{group.displayName}</Text> }),
    createTableColumn<IEntraAudienceGroup>({ columnId: 'type', renderHeaderCell: () => strings.AudienceGroupTypeLabel,
      renderCell: group => audienceTypeLabel({ groupId: group.id, displayName: group.displayName, groupType: group.groupType }) }),
    createTableColumn<IEntraAudienceGroup>({ columnId: 'members', renderHeaderCell: () => <span title={strings.AudienceMemberCountHelpLabel}>{strings.AudienceMembersLabel}</span>,
      renderCell: group => Object.prototype.hasOwnProperty.call(counts, group.id)
        ? counts[group.id] === undefined ? strings.AudienceCountUnavailableLabel : counts[group.id]?.toLocaleString()
        : <Spinner size="extra-tiny" aria-label={strings.AudienceCountLoadingLabel} /> })
  ];

  return <FluentProvider theme={{ ...(theme.isInverted ? webDarkTheme : webLightTheme),
    colorBrandBackground: theme.palette.themePrimary, colorBrandBackgroundHover: theme.palette.themeDarkAlt,
    colorBrandBackgroundPressed: theme.palette.themeDark, colorBrandForeground1: theme.palette.themePrimary,
    colorBrandStroke1: theme.palette.themePrimary }} style={{ background: 'transparent' }}>
    <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
      <Field label={strings.SelectAdminAudienceGroupsLabel}>
        <RadioGroup value={props.everyone ? 'everyone' : 'groups'} onChange={(_, data) => props.onEveryoneChange(data.value === 'everyone')}>
          <Radio value="everyone" label={strings.EveryoneAudienceLabel} />
          <Radio value="groups" label={props.single ? strings.SpecificAudienceGroupLabel : strings.SpecificAudienceGroupsLabel} />
        </RadioGroup>
      </Field>
      <Text size={200}>{strings.AudiencePermissionsHelpLabel}</Text>
      {!props.everyone && <>
        <SearchBox aria-label={strings.SearchSecurityGroupsPlaceholder} placeholder={strings.SearchSecurityGroupsPlaceholder}
            value={props.search} style={{ flex: 1, minWidth: 0, maxWidth: 'none' }}
            onKeyDown={event => { if (event.key === 'Enter') event.preventDefault(); }}
            onChange={(_, data) => props.onSearchChange(data.value)} />
        <Text size={200}>{strings.AudienceMemberCountHelpLabel}</Text>
        {props.error && <Text role="alert">{props.error}</Text>}
        {props.loading ? <Spinner label={strings.LoadingSecurityGroupsLabel} /> : <>
          {props.loaded && !props.error && !filteredGroups.length && <Text role="status">{strings.NoSecurityGroupsLabel}</Text>}
          {!!visible.length && <div style={{ overflowX: 'auto' }}>
            <DataGrid size="small" items={visible} columns={columns} getRowId={group => group.id}
              aria-label={strings.SelectAdminAudienceGroupsLabel} selectionMode={props.single ? 'single' : 'multiselect'}
              selectedItems={new Set(Object.keys(props.selected))}
              onSelectionChange={(_, data) => {
                const selected = props.single ? {} : { ...props.selected };
                visible.forEach(group => { delete selected[group.id]; });
                visible.forEach(group => { if (data.selectedItems.has(group.id)) selected[group.id] = group.displayName; });
                props.onSelectionChange(selected);
              }}>
              <DataGridHeader><DataGridRow selectionCell={{ 'aria-label': strings.SelectAdminAudienceGroupsLabel }}>
                {({ renderHeaderCell }) => <DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>}
              </DataGridRow></DataGridHeader>
              <DataGridBody<IEntraAudienceGroup>>{({ item, rowId }) => <DataGridRow<IEntraAudienceGroup> key={rowId}
                selectionCell={{ 'aria-label': item.displayName }}>
                {({ renderCell }) => <DataGridCell>{renderCell(item)}</DataGridCell>}
              </DataGridRow>}</DataGridBody>
            </DataGrid>
          </div>}
          {pageCount > 1 && <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }}>
            <Button disabled={currentPage === 0} onClick={() => setPage(currentPage - 1)}>{strings.AudiencePreviousPageLabel}</Button>
            <Text role="status">{formatLocalizedString(strings.AudiencePageLabel, String(currentPage + 1), String(pageCount))}</Text>
            <Button disabled={currentPage === pageCount - 1} onClick={() => setPage(currentPage + 1)}>{strings.NextLabel}</Button>
          </div>}
        </>}
        {!!Object.keys(props.selected).length && <Text role="status">{formatLocalizedString(strings.SelectedGroupsLabel, Object.keys(props.selected).map(id => props.selected[id]).join(', '))}</Text>}
      </>}
      <div><Button appearance="primary" onClick={props.onApply} disabled={!props.everyone && !Object.keys(props.selected).length}>{props.actionLabel}</Button></div>
    </div>
  </FluentProvider>;
}
