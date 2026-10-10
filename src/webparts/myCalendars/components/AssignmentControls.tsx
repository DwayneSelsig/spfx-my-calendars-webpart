import * as React from 'react';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { Stack } from '@fluentui/react/lib/Stack';
import * as strings from 'MyCalendarsWebPartStrings';
import { IAdminAllowedOverrides, IAssignmentPolicy, IAudienceGroup } from '../models/ICalendarSettings';
import { normalizeAllowedOverrides, normalizeAssignmentPolicy } from '../services/CalendarSettingsService';

export function assignmentPolicyLabel(policy: Partial<IAssignmentPolicy>): string {
  return policy.isMandatory ? strings.MandatoryPolicyLabel : policy.defaultEnabled !== false ? strings.DefaultPolicyLabel : strings.AvailablePolicyLabel;
}

export function audienceTypeLabel(group: IAudienceGroup): string {
  return group.groupType === 'microsoft365' ? strings.GroupTypeMicrosoft365Label : group.groupType === 'mailEnabledSecurity' ? strings.GroupTypeMailSecurityLabel : group.groupType === 'security' ? strings.GroupTypeSecurityLabel : '';
}

export const AssignmentPolicyControl: React.FC<{ policy: Partial<IAssignmentPolicy>; onChange: (policy: IAssignmentPolicy) => void }> = ({ policy, onChange }) => (
  <Dropdown label={strings.PolicyLabel} selectedKey={policy.isMandatory ? 'mandatory' : policy.defaultEnabled !== false ? 'default' : 'available'}
    options={[{ key: 'mandatory', text: strings.MandatoryPolicyLabel }, { key: 'default', text: strings.DefaultPolicyLabel }, { key: 'available', text: strings.AvailablePolicyLabel }]}
    onChange={(_, option) => onChange(normalizeAssignmentPolicy({ isMandatory: option?.key === 'mandatory', defaultEnabled: option?.key !== 'available' }))} />
);

export const AllowedOverrideControls: React.FC<{ value?: IAdminAllowedOverrides; planner?: boolean; onChange: (value: IAdminAllowedOverrides) => void }> = ({ value, planner, onChange }) => {
  const allowed = normalizeAllowedOverrides(value);
  const fields: Array<{ key: keyof IAdminAllowedOverrides; label: string }> = [
    { key: 'name', label: strings.AllowNameOverrideLabel }, { key: 'color', label: strings.AllowColorOverrideLabel }, { key: 'showSourceLogo', label: strings.AllowLogoOverrideLabel }
  ];
  if (planner) fields.push({ key: 'plannerAssignedToMeOnly', label: strings.AllowAssignedFilterOverrideLabel }, { key: 'showCompletedTasks', label: strings.AllowCompletedFilterOverrideLabel });
  return <Stack tokens={{ childrenGap: 8 }}>{fields.map(field => <Checkbox key={field.key} label={field.label} checked={allowed[field.key]} onChange={(_, checked) => onChange({ ...allowed, [field.key]: !!checked })} />)}</Stack>;
};

export const AssignmentGroup: React.FC<{ title: string; description?: string }> = ({ title, description, children }) => (
  <details open style={{ marginBottom: 12 }}><summary style={{ cursor: 'pointer', padding: '8px 0' }}><strong>{title}</strong>{description && <span style={{ marginLeft: 8 }}>{description}</span>}</summary><div style={{ paddingLeft: 16 }}>{children}</div></details>
);
