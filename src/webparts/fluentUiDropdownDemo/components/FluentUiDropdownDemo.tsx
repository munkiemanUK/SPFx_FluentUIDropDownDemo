import * as React from 'react';
import styles from './FluentUiDropdownDemo.module.scss';
import { Guid } from '@microsoft/sp-core-library';
import { IFluentUiDropdownDemoProps } from './IFluentUiDropdownDemoProps';
import { Web } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { escape } from '@microsoft/sp-lodash-subset';
import {Dropdown, PrimaryButton, IDropdownOption} from '@fluentui/react';

var arr = [];
export interface IDropdownStates
{
  singleValueDropdown:string;
  multiValueDropdown:any;
}

export default class FluentUiDropdownDemo extends React.Component<IFluentUiDropdownDemoProps, IDropdownStates{}> {
  constructor(props)
  {
    super(props);
    this.state={
      singleValueDropdown:"",
      multiValueDropdown:[]
    };
    
  }

  public onDropdownChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.setState({ singleValueDropdown: item.key as string});
  }

  public onDropdownMultiChange = async (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): Promise<void> => {
    if (item.selected) {
      await arr.push(item.key as string);
    }
    else {
      await arr.indexOf(item.key) !== -1 && arr.splice(arr.indexOf(item.key), 1);
    }
    await this.setState({ multiValueDropdown: arr });
  }

  private async Save(e) {
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("FluentUIDropdown").items.add({
      Title: Guid.newGuid().toString(),
      SingleValueDropdown: this.state.singleValueDropdown,
      MultiValueDropdown: { results: this.state.multiValueDropdown }

    }).then(i => {
      console.log(i);
    });
    alert("Submitted Successfully");
  }

  public render(): React.ReactElement<IFluentUiDropdownDemoProps> {
      return (
        <div className={ styles.fluentUiDropdown }>
          <h1>Fluent UI Dropdown</h1>
          <Dropdown
            placeholder="Single Select Dropdown..."
            selectedKey={this.state.singleValueDropdown}
            label="Single Select Dropdown"
            options={this.props.singleValueOptions}
            onChange={this.onDropdownChange}
          />
          <br />
          <Dropdown
            placeholder="Multi Select Dropdown..."
            defaultSelectedKeys={this.state.multiValueDropdown}
            label="Multi Select Dropdown"
            multiSelect
            options={this.props.multiValueOptions}
            onChange={this.onDropdownMultiChange}
          />
          <div>
            <br />
            <br />
            <PrimaryButton onClick={e => this.Save(e)}>Submit</PrimaryButton>
          </div>
        </div>
    );
  }
}