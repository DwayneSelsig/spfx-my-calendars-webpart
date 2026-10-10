import * as React from 'react';
import { IBasePickerSuggestionsProps, NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import * as strings from 'MyCalendarsWebPartStrings';
import { IMailboxSearchResult } from '../services/ExchangeCalendarService';

export interface IMailboxPeoplePickerProps {
  selectedMailbox?: IMailboxSearchResult;
  placeholder?: string;
  disabled?: boolean;
  onResolveSuggestions: (query: string) => Promise<IMailboxSearchResult[]>;
  onInputChange: (query: string) => void;
  onChange: (mailbox?: IMailboxSearchResult) => void;
}

type IMailboxPersona = IPersonaProps & { mailbox: IMailboxSearchResult };

function getInitials(displayName: string): string {
  const words = displayName.trim().split(/\s+/).filter(Boolean);
  return words.length > 1 ? `${words[0].charAt(0)}${words[1].charAt(0)}` : (words[0] || '?').charAt(0);
}

function toPersona(mailbox: IMailboxSearchResult): IMailboxPersona {
  return {
    key: mailbox.id,
    text: mailbox.displayName,
    secondaryText: mailbox.mail || mailbox.userPrincipalName,
    tertiaryText: mailbox.jobTitle,
    imageUrl: mailbox.imageUrl,
    imageInitials: getInitials(mailbox.displayName),
    mailbox
  };
}

/**
 * Small local mailbox picker built on Fluent UI primitives.
 * It intentionally has no dependency on @pnp/spfx-controls-react.
 */
export class MailboxPeoplePicker extends React.Component<IMailboxPeoplePickerProps> {
  private selectionInProgress = false;

  private readonly suggestionProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: strings.MailboxPickerSuggestionsLabel,
    noResultsFoundText: strings.MailboxPickerNoResultsLabel,
    loadingText: strings.MailboxPickerLoadingLabel,
    searchingText: strings.MailboxPickerLoadingLabel,
    resultsMaximumNumber: 5
  };

  private resolveSuggestions = async (query: string): Promise<IPersonaProps[]> => {
    if (query.trim().length < 2) return [];
    try {
      const mailboxes = await this.props.onResolveSuggestions(query);
      return mailboxes.map(toPersona);
    } catch (error) {
      console.error('Mailbox picker search failed:', error);
      return [];
    }
  };

  private handleChange = (items?: IPersonaProps[]): void => {
    const selected = items && items[0] as IMailboxPersona | undefined;
    this.props.onChange(selected?.mailbox);
  };

  private handleItemSelected = (item?: IPersonaProps): IPersonaProps | null => {
    const selected = item as IMailboxPersona | undefined;
    if (selected?.mailbox) {
      // Notify the owner before Fluent UI clears the input and completes the
      // controlled selectedItems update.
      this.selectionInProgress = true;
      this.props.onChange(selected.mailbox);
    }
    return item || null;
  };

  private handleInputChange = (query: string): string => {
    if (!query && this.selectionInProgress) {
      this.selectionInProgress = false;
      return query;
    }
    this.props.onInputChange(query);
    return query;
  };

  public render(): React.ReactElement {
    const selectedItems = this.props.selectedMailbox ? [toPersona(this.props.selectedMailbox)] : [];

    return (
      <NormalPeoplePicker
        pickerSuggestionsProps={this.suggestionProps}
        onResolveSuggestions={this.resolveSuggestions}
        onEmptyInputFocus={() => []}
        getTextFromItem={item => item.text || item.secondaryText || ''}
        selectedItems={selectedItems}
        itemLimit={1}
        disabled={this.props.disabled}
        resolveDelay={250}
        onItemSelected={this.handleItemSelected}
        onChange={this.handleChange}
        onInputChange={this.handleInputChange}
        removeButtonAriaLabel={strings.DeleteLabel}
        inputProps={{
          'aria-label': strings.EnterMailboxEmailLabel,
          placeholder: this.props.placeholder
        }}
      />
    );
  }
}

