import * as React from 'react';
import styles from './PersonaCard.module.scss';
import { IPersonaCardProps } from './IPersonaCardProps';
import { IPersonaCardState } from './IPersonaCardState';
import {
  Log, Environment, EnvironmentType,
} from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { Icon } from '@fluentui/react/lib/Icon';
import {
  Persona,
  PersonaSize,
  DocumentCard,
  DocumentCardType,
} from 'office-ui-fabric-react';

const EXP_SOURCE: string = 'SPFxDirectory';
const LIVE_PERSONA_COMPONENT_ID: string =
  '914330ee-2df2-4f6e-a858-30c23a812408';

export class PersonaCard extends React.Component<
  IPersonaCardProps,
  IPersonaCardState
> {
  constructor(props: IPersonaCardProps) {
    super(props);

    this.state = { livePersonaCard: undefined, pictureUrl: undefined };
  }
  /**
   *
   *
   * @memberof PersonaCard
   */
  public async componentDidMount() {
    if (Environment.type !== EnvironmentType.Local) {
      const sharedLibrary = await this._loadSPComponentById(
        LIVE_PERSONA_COMPONENT_ID
      );
      const livePersonaCard: any = sharedLibrary.LivePersonaCard;
      this.setState({ livePersonaCard: livePersonaCard });
    }
  }

  /**
   *
   *
   * @param {IPersonaCardProps} prevProps
   * @param {IPersonaCardState} prevState
   * @memberof PersonaCard
   */
  public componentDidUpdate(
    prevProps: IPersonaCardProps,
    prevState: IPersonaCardState
  ): void { }

  /**
   *
   *
   * @private
   * @returns
   * @memberof PersonaCard
   */
  private _LivePersonaCard() {
    return React.createElement(
      this.state.livePersonaCard,
      {
        serviceScope: this.props.context.serviceScope,
        upn: this.props.profileProperties.Email,
        onCardOpen: () => {
          console.log('LivePersonaCard Open');
        },
        onCardClose: () => {
          console.log('LivePersonaCard Close');
        },
      },
      this._PersonaCard()
    );
  }

  /**
   *
   *
   * @private
   * @returns {JSX.Element}
   * @memberof PersonaCard
   */
  private _PersonaCard(): JSX.Element {
    return (
      <DocumentCard
        className={styles.documentCard}
        type={DocumentCardType.normal}
      >
        <div className={styles.persona}>
          <Persona
            text={this.props.profileProperties.DisplayName}
            imageUrl={this.props.profileProperties.PictureUrl}
            size={PersonaSize.size72}
            imageShouldFadeIn={false}
            imageShouldStartVisible={true}
          >
          </Persona>
          {/* {this.props.profileProperties.DisplayName ? (
            <div style={{ marginTop: '5em' }}>
              <span style={{ marginLeft: 5, fontSize: '20px' }}>
                {' '}
                {this.props.profileProperties.DisplayName}
              </span>
            </div>
          ) : (
            ''
          )} */}
          {this.props.profileProperties.Title ? (
            <div className={styles.flex} style={{ marginTop: '1em', display: 'flex !important' }}>
              <Icon iconName="UserOptional" style={{ fontSize: '14px' }} />
              <span style={{ marginLeft: 5, fontSize: '14px' }}>
                {' '}
                {this.props.profileProperties.Title}
              </span>
            </div>
          ) : (
            ''
          )}
          {this.props.profileProperties.Department ? (
            <div className={styles.flex} style={{ marginTop: '0.5em', display: 'flex !important' }}>
              <Icon iconName="CityNext" style={{ fontSize: '14px' }} />
              <span style={{ marginLeft: 5, fontSize: '14px' }}>
                {' '}
                {this.props.profileProperties.Department}
              </span>
            </div>
          ) : (
            ''
          )}
          {this.props.profileProperties.Email ? (
            <div style={{ marginTop: '0.5em' }}>
              <Icon iconName="Mail" style={{ fontSize: '14px' }} />
              <span style={{ marginLeft: 5, fontSize: '14px' }}>
                {' '}
                {this.props.profileProperties.Email}
              </span>
            </div>
          ) : (
            ''
          )}
          {this.props.profileProperties.WorkPhone ? (
            <div style={{ marginTop: '0.5em' }}>
              <Icon iconName="Phone" style={{ fontSize: '14px' }} />
              <span style={{ marginLeft: 5, fontSize: '14px' }}>
                {' '}
                {this.props.profileProperties.WorkPhone}
              </span>
            </div>
          ) : (
            ''
          )}
          {this.props.profileProperties.MobilePhone ? (
            <div className={styles.textOverflow} style={{ marginTop: '0.5em' }}>
              <Icon iconName="CellPhone" style={{ fontSize: '14px' }} />
              <span style={{ marginLeft: 5, fontSize: '14px' }}>
                {' '}
                {this.props.profileProperties.MobilePhone}
              </span>
            </div>
          ) : (
            ''
          )}
          {this.props.profileProperties.JoinDate ? (
            <div className={styles.textOverflow} style={{ marginTop: '0.5em' }}>
              <Icon iconName="DateTime" style={{ fontSize: '14px' }} />
              <span style={{ marginLeft: 5, fontSize: '14px' }}>
                {' '}
                {this.props.profileProperties.JoinDate}
              </span>
            </div>
          ) : (
            ''
          )}
          {this.props.profileProperties.TransferDate ? (
            <div className={styles.textOverflow} style={{ marginTop: '0.5em' }}>
              <Icon iconName="DateTime" style={{ fontSize: '14px' }} />
              <span style={{ marginLeft: 5, fontSize: '14px' }}>
                {' '}
                {this.props.profileProperties.TransferDate}
              </span>
            </div>
          ) : (
            ''
          )}
          {this.props.profileProperties.LastServiceDate ? (
            <div className={styles.textOverflow} style={{ marginTop: '0.5em' }}>
              <Icon iconName="DateTime" style={{ fontSize: '14px' }} />
              <span style={{ marginLeft: 5, fontSize: '14px' }}>
                {' '}
                {this.props.profileProperties.LastServiceDate}
              </span>
            </div>
          ) : (
            ''
          )}
          {this.props.profileProperties.ReportingOfficer ? (
            <div className={styles.textOverflow} style={{ marginTop: '0.5em' }}>
              <Icon iconName="ManagerSelfService" style={{ fontSize: '14px' }} />
              <span style={{ marginLeft: 5, fontSize: '14px' }}>
                {' '}
                {this.props.profileProperties.ReportingOfficer}
              </span>
            </div>
          ) : (
            ''
          )}

        </div>
      </DocumentCard>
    );
  }
  /**
   * Load SPFx component by id, SPComponentLoader is used to load the SPFx components
   * @param componentId - componentId, guid of the component library
   */
  private async _loadSPComponentById(componentId: string): Promise<any> {
    try {
      const component: any = await SPComponentLoader.loadComponentById(
        componentId
      );
      return component;
    } catch (error) {
      Promise.reject(error);
      Log.error(EXP_SOURCE, error, this.props.context.serviceScope);
    }
  }

  /**
   *
   *
   * @returns {React.ReactElement<IPersonaCardProps>}
   * @memberof PersonaCard
   */
  public render(): React.ReactElement<IPersonaCardProps> {
    return (
      <div className={styles.personaContainer}>
        {this.state.livePersonaCard
          ? this._PersonaCard()
          : this._PersonaCard()}
      </div>
    );
  }
}
