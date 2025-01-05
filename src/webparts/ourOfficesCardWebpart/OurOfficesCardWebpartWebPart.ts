import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneDropdown, PropertyPaneButton, PropertyPaneButtonType } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import OurOfficesCardWebpart from './components/OurOfficesCardWebpart';
import { IOurOfficesCardWebpartProps } from './components/IOurOfficesCardWebpartProps';

export interface IOurOfficesCardWebpartWebPartProps {
  flagUrl1: string;
  country1: string;
  address1: string;
  email1: string;
  phone1: string;
  flagUrl2: string;
  country2: string;
  address2: string;
  email2: string;
  phone2: string;
  flagUrl3: string;
  country3: string;
  address3: string;
  email3: string;
  phone3: string;
  headquarters: string;
}

export default class OurOfficesCardWebpartWebPart extends BaseClientSideWebPart<IOurOfficesCardWebpartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IOurOfficesCardWebpartProps> = React.createElement(
      OurOfficesCardWebpart,
      {
        flagUrl1: this.properties.flagUrl1,
        country1: this.properties.country1,
        address1: this.properties.address1,
        email1: this.properties.email1,
        phone1: this.properties.phone1,
        flagUrl2: this.properties.flagUrl2,
        country2: this.properties.country2,
        address2: this.properties.address2,
        email2: this.properties.email2,
        phone2: this.properties.phone2,
        flagUrl3: this.properties.flagUrl3,
        country3: this.properties.country3,
        address3: this.properties.address3,
        email3: this.properties.email3,
        phone3: this.properties.phone3,
        headquarters: this.properties.headquarters
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    // Initialize default values
    this.properties.flagUrl1 = this.properties.flagUrl1 || 'https://flagcdn.com/w320/us.png';
    this.properties.country1 = this.properties.country1 || 'United States';
    this.properties.address1 = this.properties.address1 || '1234 Elm Street, Springfield, USA';
    this.properties.email1 = this.properties.email1 || 'contact.us@example.com';
    this.properties.phone1 = this.properties.phone1 || '+1 234 567 890';

    this.properties.flagUrl2 = this.properties.flagUrl2 || 'https://flagcdn.com/w320/gb.png';
    this.properties.country2 = this.properties.country2 || 'United Kingdom';
    this.properties.address2 = this.properties.address2 || '5678 Oak Avenue, London, UK';
    this.properties.email2 = this.properties.email2 || 'contact.uk@example.com';
    this.properties.phone2 = this.properties.phone2 || '+44 20 7946 0958';

    this.properties.flagUrl3 = this.properties.flagUrl3 || 'https://flagcdn.com/w320/ca.png';
    this.properties.country3 = this.properties.country3 || 'Canada';
    this.properties.address3 = this.properties.address3 || '910 Maple Street, Toronto, Canada';
    this.properties.email3 = this.properties.email3 || 'contact.ca@example.com';
    this.properties.phone3 = this.properties.phone3 || '+1 416 555 0123';

    this.properties.headquarters = this.properties.headquarters || '1';

    return super.onInit();
  }

  private saveProperties(): void {
    this.context.propertyPane.close();
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Configure Office Details" },
          groups: [
            {
              groupName: "Office 1",
              groupFields: [
                PropertyPaneTextField('flagUrl1', { label: 'Flag URL' }),
                PropertyPaneTextField('country1', { label: 'Country' }),
                PropertyPaneTextField('address1', { label: 'Address' }),
                PropertyPaneTextField('email1', { label: 'Email' }),
                PropertyPaneTextField('phone1', { label: 'Phone' }),
              ]
            },
            {
              groupName: "Office 2",
              groupFields: [
                PropertyPaneTextField('flagUrl2', { label: 'Flag URL' }),
                PropertyPaneTextField('country2', { label: 'Country' }),
                PropertyPaneTextField('address2', { label: 'Address' }),
                PropertyPaneTextField('email2', { label: 'Email' }),
                PropertyPaneTextField('phone2', { label: 'Phone' }),
              ]
            },
            {
              groupName: "Office 3",
              groupFields: [
                PropertyPaneTextField('flagUrl3', { label: 'Flag URL' }),
                PropertyPaneTextField('country3', { label: 'Country' }),
                PropertyPaneTextField('address3', { label: 'Address' }),
                PropertyPaneTextField('email3', { label: 'Email' }),
                PropertyPaneTextField('phone3', { label: 'Phone' }),
              ]
            },
            {
              groupName: "Settings",
              groupFields: [
                PropertyPaneDropdown('headquarters', {
                  label: 'Headquarters',
                  options: [
                    { key: '1', text: 'Office 1' },
                    { key: '2', text: 'Office 2' },
                    { key: '3', text: 'Office 3' }
                  ],
                  selectedKey: this.properties.headquarters
                }),
                PropertyPaneButton('saveButton', {
                  text: 'Save',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: () => this.saveProperties()
                })
              ]
            }
          ]
        }
      ]
    };
  }
}