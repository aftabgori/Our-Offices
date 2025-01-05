// import * as React from 'react';
// import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
// import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import styles from './OurOfficesCardWebpart.module.scss';
import * as React from 'react';

export interface IOurOfficesCardWebpartProps {
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

export default class OurOfficesCardWebpart extends React.Component<IOurOfficesCardWebpartProps, {}> {
  public render(): React.ReactElement<IOurOfficesCardWebpartProps> {
    const offices = [
      {
        flagUrl: this.props.flagUrl1,
        country: this.props.country1,
        address: this.props.address1,
        email: this.props.email1,
        phone: this.props.phone1,
        isHeadquarters: this.props.headquarters === '1'
      },
      {
        flagUrl: this.props.flagUrl2,
        country: this.props.country2,
        address: this.props.address2,
        email: this.props.email2,
        phone: this.props.phone2,
        isHeadquarters: this.props.headquarters === '2'
      },
      {
        flagUrl: this.props.flagUrl3,
        country: this.props.country3,
        address: this.props.address3,
        email: this.props.email3,
        phone: this.props.phone3,
        isHeadquarters: this.props.headquarters === '3'
      }
    ];

    return (
      <div className={styles.officesContainer}>
        {offices.map((office, index) => (
          <div key={index} className={styles.officeCard}>
            {office.isHeadquarters && <div className={styles.headquartersTitle}>Headquarters</div>}
            <img src={office.flagUrl} alt={`Flag of ${office.country}`} className={styles.flagImage} />
            <h3 className={styles.countryName}>{office.country}</h3>
            <p className={styles.address}>{office.address}</p>
            <p className={styles.email}>
              <i className={`ms-Icon ms-Icon--Mail`} aria-hidden="true"></i> {office.email}
            </p>
            <p className={styles.phone}>
              <i className={`ms-Icon ms-Icon--Phone`} aria-hidden="true"></i> {office.phone}
            </p>
          </div>
        ))}
      </div>
    );
  }
}