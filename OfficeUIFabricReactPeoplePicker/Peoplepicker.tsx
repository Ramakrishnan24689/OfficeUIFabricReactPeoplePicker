import * as React from 'react';
import { BaseComponent } from 'office-ui-fabric-react/lib/Utilities';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { IInputs } from "./generated/ManifestTypes";

import {
  IBasePickerSuggestionsProps,
  IBasePicker,
  NormalPeoplePicker,
  ValidationState
} from 'office-ui-fabric-react/lib/Pickers';

export interface IPeoplePersona {
  text?: string;
  secondaryText?: string;
}

export interface IPeopleProps {
  people?: any;
  preselectedpeople?: any;
  context?: ComponentFramework.Context<IInputs>;
  peopleList?: (newValue: any) => void;
  isPickerDisabled?: boolean;
}

export interface IPeoplePickerState {
  currentPicker?: number | string;
  delayResults?: boolean;
  peopleList: IPersonaProps[];
  mostRecentlyUsed: IPersonaProps[];
  currentSelectedItems?: IPersonaProps[];
}

const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: 'Suggested People',
  mostRecentlyUsedHeaderText: 'Suggested Contacts',
  noResultsFoundText: 'No results found',
  loadingText: 'Loading',
  showRemoveButtons: true,
  suggestionsAvailableAlertText: 'People Picker Suggestions available',
  suggestionsContainerAriaLabel: 'Suggested contacts'
};


export class PeoplePickerTypes extends BaseComponent<any, IPeoplePickerState> {
  // All pickers extend from BasePicker specifying the item type.
  private _picker = React.createRef<IBasePicker<IPersonaProps>>();

  constructor(props: IPeopleProps) {
    super(props);

    this.state = {
      currentPicker: 1,
      delayResults: false,
      peopleList: this.props.people,
      mostRecentlyUsed: [],
      currentSelectedItems: []
    };
    initializeIcons();
    this.handleChange = this.handleChange.bind(this);
  }

  public render() {
    return (
      <NormalPeoplePicker
        onResolveSuggestions={this._onFilterChanged}
        onEmptyInputFocus={this._returnMostRecentlyUsed}
        getTextFromItem={this._getTextFromItem}
        pickerSuggestionsProps={suggestionProps}
        className={'ms-PeoplePicker'}
        key={'normal'}
        onRemoveSuggestion={this._onRemoveSuggestion}
        onValidateInput={this._validateInput}
        removeButtonAriaLabel={'Remove'}
        defaultSelectedItems={this.props.preselectedpeople}
        onItemSelected={this._onItemSelected}
        inputProps={{
          onBlur: (ev: React.FocusEvent<HTMLInputElement>) => {
          },
          onFocus: (ev: React.FocusEvent<HTMLInputElement>) => {
          },
          'aria-label': 'People Picker'
        }}
        componentRef={this._picker}
        onInputChange={this._onInputChange}
        resolveDelay={300}
        disabled={this.props.isPickerDisabled}
        onChange={() => this.handleChange()}
      />
    );

  }

  private handleChange() {
    if (this._picker !== undefined) {
      this.getPeopleRelatedVal();
    }
  }

  private _onItemSelected = (item: any): Promise<IPersonaProps> => {
    const processedItem = { ...item };
    processedItem.text = `${item.text}`;
    return new Promise<IPersonaProps>((resolve, reject) =>
      resolve(processedItem));
  };



  private async getPeopleRelatedVal() {
    try {
      let tempPeople: IPeoplePersona[] = [];
      await Promise.all(this._picker.current!.items!.map((item) => {
        tempPeople.push({ "text": item.text, "secondaryText": item.secondaryText });
      }));
      if (this.props.peopleList) {
        this.props.peopleList(tempPeople);
      }
    }
    catch (err) {
      console.log(err);
    }
  }

  private _getTextFromItem(persona: IPersonaProps): string {
    return persona.text as string;
  }

  private _onRemoveSuggestion = (item: IPersonaProps): void => {
    const { peopleList, mostRecentlyUsed: mruState } = this.state;
    const indexPeopleList: number = peopleList.indexOf(item);
    const indexMostRecentlyUsed: number = mruState.indexOf(item);

    if (indexPeopleList >= 0) {
      const newPeople: IPersonaProps[] = peopleList.slice(0, indexPeopleList).concat(peopleList.slice(indexPeopleList + 1));
      this.setState({ peopleList: newPeople });
    }

    if (indexMostRecentlyUsed >= 0) {
      const newSuggestedPeople: IPersonaProps[] = mruState
        .slice(0, indexMostRecentlyUsed)
        .concat(mruState.slice(indexMostRecentlyUsed + 1));
      this.setState({ mostRecentlyUsed: newSuggestedPeople });
    }
  }

  private _onFilterChanged = (
    filterText: string,
    currentPersonas: IPersonaProps[] | undefined,
    limitResults?: number
  ): any => {
    if (filterText) {
      if (filterText.length > 2) {
        return this._searchUsers(filterText);
      }
    } else {
      return [];
    }
  };


  private _searchUsers(filterText: string): IPersonaProps[] | Promise<IPersonaProps[]> {
    return new Promise(async (resolve: any, reject: any) => {
      let People: IPersonaProps[] = [];
      try {
        let tempPeople: any = [];
        tempPeople = await this.props.context.webAPI.retrieveMultipleRecords(this.props.context.parameters.entityName.raw!, "?$select=fullname,internalemailaddress&$filter=startswith(fullname,'" + filterText + "')");
        await Promise.all(tempPeople.entities.map((entity: any) => {
          People.push({ "text": entity.fullname, "secondaryText": entity.internalemailaddress }); //change fieldname if values are different
        }));
        resolve(People);
      }
      catch (err) {
        console.log(err);
        reject(People);
      }
    });
  }


  private _returnMostRecentlyUsed = (currentPersonas: any): IPersonaProps[] | Promise<IPersonaProps[]> => {
    let { mostRecentlyUsed } = this.state;
    mostRecentlyUsed = this._removeDuplicates(mostRecentlyUsed, currentPersonas);
    return this._filterPromise(mostRecentlyUsed);
  };

  private _filterPromise(personasToReturn: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> {
    if (this.state.delayResults) {
      return this._convertResultsToPromise(personasToReturn);
    } else {
      return personasToReturn;
    }
  }

  private _listContainsPersona(persona: IPersonaProps, personas: IPersonaProps[]) {
    if (!personas || !personas.length || personas.length === 0) {
      return false;
    }
    return personas.filter(item => item.text === persona.text).length > 0;
  }

  private _filterPersonasByText(filterText: string): IPersonaProps[] {
    return this.state.peopleList.filter(item => this._doesTextStartWith(item.text as string, filterText));
  }

  private _doesTextStartWith(text: string, filterText: string): boolean {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  }

  private _convertResultsToPromise(results: IPersonaProps[]): Promise<IPersonaProps[]> {
    return new Promise<IPersonaProps[]>((resolve, reject) => setTimeout(() => resolve(results), 2000));
  }

  private _removeDuplicates(personas: IPersonaProps[], possibleDupes: IPersonaProps[]) {
    return personas.filter(persona => !this._listContainsPersona(persona, possibleDupes));
  }

  private _validateInput = (input: string): ValidationState => {
    if (input.indexOf('@') !== -1) {
      return ValidationState.valid;
    } else if (input.length > 1) {
      return ValidationState.warning;
    } else {
      return ValidationState.invalid;
    }
  }

  /**
   * Takes in the picker input and modifies it in whichever way
   * the caller wants, i.e. parsing entries copied from Outlook (sample
   * input: "Aaron Reid <aaron>").
   *
   * @param input The text entered into the picker.
   */
  private _onInputChange(input: string): string {
    const outlookRegEx = /<.*>/g;
    const emailAddress = outlookRegEx.exec(input);

    if (emailAddress && emailAddress[0]) {
      return emailAddress[0].substring(1, emailAddress[0].length - 1);
    }
    return input;
  }

}

