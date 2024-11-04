import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneSlider } from '@microsoft/sp-property-pane';
import * as strings from 'BlurbWebPartStrings';
import { Blurb } from './components/Blurb';
import { IBlurbProps } from './components/IBlurbProps';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldIconPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldIconPicker';
import { initializeIcons } from '@fluentui/react/lib/Icons';

export interface IBlurbWebPartProps {
  description: string;
  containerCount: number;
  containers: Array<{
    fontColor: string;
    icon: string;
    backgroundColor: string;
    borderColor: string;
    borderRadius: string;
    title: string;
    text: string;
  }>;
}

export default class BlurbWebPart extends BaseClientSideWebPart<IBlurbWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private selectedContainerIndex: number = -1;
  private _isEditMode: boolean = false;

  public render(): void {
    const element: React.ReactElement<IBlurbProps> = React.createElement(
      Blurb,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        containers: this.properties.containers || [],
        containerCount: this.properties.containerCount || 1,

        onContainerClick: async (index: number) => {
          // Close the property pane if it's already open
          if (this.context.propertyPane.isRenderedByWebPart()) {
            this.context.propertyPane.close();
          }
          // Delay to ensure the pane has closed
          await new Promise(resolve => setTimeout(resolve, 10));

          // Set the selected container and enter edit mode
          this.selectedContainerIndex = index;
          this._isEditMode = true;
          this.context.propertyPane.refresh();
          this.context.propertyPane.open();
        },
        onEditClick: async (index: number) => {
          // Handle edit click by closing and reopening the property pane
          this.selectedContainerIndex = index;
          this._isEditMode = true;

          if (this.context.propertyPane.isRenderedByWebPart()) {
            this.context.propertyPane.close();
          }

          await new Promise(resolve => setTimeout(resolve, 10)); // Delay for smooth reopening
          this.context.propertyPane.refresh();
          this.context.propertyPane.open();
        },
        onMoveClick: (index: number, direction: 'up' | 'down') => {
          if (direction === 'up' && index > 0) {
            const temp = this.properties.containers[index];
            this.properties.containers[index] = this.properties.containers[index - 1];
            this.properties.containers[index - 1] = temp;
          } else if (direction === 'down' && index < this.properties.containers.length - 1) {
            const temp = this.properties.containers[index];
            this.properties.containers[index] = this.properties.containers[index + 1];
            this.properties.containers[index + 1] = temp;
          }
          this.render(); // Re-render to reflect the new order
        },
        onRemoveClick: (index: number, updatedCount: number) => {
          this.properties.containers.splice(index, 1);
          this.properties.containerCount = updatedCount; // Update the container count
          this.render(); // Re-render the component to reflect the changes
        },
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    initializeIcons();
    if (!this.properties.containerCount) {
      this.properties.containerCount = 1;
    }

    if (!this.properties.containers) {
      this.properties.containers = [];
    }

    const currentContainerCount = this.properties.containers.length;
    if (this.properties.containerCount > currentContainerCount) {
      for (let i = currentContainerCount; i < this.properties.containerCount; i++) {
        this.properties.containers.push({
          icon: 'CannedChat',
          backgroundColor: '#FAF9F8',
          borderColor: '#EDEBE9',
          borderRadius: '0',
          fontColor: '#323130',
          title: ``,
          text: ''
        });
      }
    } else if (this.properties.containerCount < currentContainerCount) {
      this.properties.containers.splice(this.properties.containerCount);
    }

    const message = await this._getEnvironmentMessage();
    this._environmentMessage = message;

    return super.onInit();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onPropertyPaneConfigurationComplete(): void {
    this._isEditMode = false;
    this.selectedContainerIndex = -1;
  }

  private async _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      const context = await this.context.sdks.microsoftTeams.teamsJs.app.getContext();
      let environmentMessage: string = '';
      switch (context.app.host.name) {
        case 'Office':
          environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
          break;
        case 'Outlook':
          environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
          break;
        case 'Teams':
          environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
          break;
        default:
          environmentMessage = strings.UnknownEnvironment;
      }
      return environmentMessage;
    }
    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string | number, newValue: string | number): void {
    if (propertyPath === 'containerCount' && newValue !== oldValue) {
      const newContainerCount = newValue as number;
      const currentContainerCount = this.properties.containers.length;

      if (newContainerCount > currentContainerCount) {
        for (let i = currentContainerCount; i < newContainerCount; i++) {
          this.properties.containers.push({
            icon: 'CannedChat',
            backgroundColor: '#FAF9F8',
            borderColor: '#EDEBE9',
            borderRadius: '0',
            fontColor: '#323130',
            title: ``,
            text: ''
          });
        }
      } else if (newContainerCount < currentContainerCount) {
        this.properties.containers.splice(newContainerCount);
      }
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    if (this._isEditMode && this.selectedContainerIndex !== -1) {
      const selectedContainer = this.properties.containers[this.selectedContainerIndex] || {};
      return {
        pages: [
          {
            header: { description: `Configure Blurb ${this.selectedContainerIndex + 1}` },
            groups: [
              {
                groupName: "Blurb Settings",
                groupFields: [
                  PropertyFieldIconPicker(`containers[${this.selectedContainerIndex}].icon`, {
                    label: "Select Icon",
                    currentIcon: selectedContainer.icon,
                    onSave: (iconName: string) => {
                      this.properties.containers[this.selectedContainerIndex].icon = iconName;
                      this.render();
                    },
                    properties: this.properties,
                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                    buttonLabel: "Select Icon",
                    renderOption: "panel",
                    key: `iconPicker-${this.selectedContainerIndex}`
                  }),
                  PropertyPaneTextField(`containers[${this.selectedContainerIndex}].title`, {
                    label: `Blurb Title ${this.selectedContainerIndex + 1}`,
                    value: selectedContainer.title || ''
                  }),
                  PropertyPaneTextField(`containers[${this.selectedContainerIndex}].text`, {
                    label: `Blurb Text ${this.selectedContainerIndex + 1}`,
                    value: selectedContainer.text || '',
                    multiline: true, // Multi-line text input
                    resizable: true // Allows vertical resizing
                  }),
                  PropertyFieldColorPicker(`containers[${this.selectedContainerIndex}].fontColor`, {
                    label: "Font Color",
                    selectedColor: selectedContainer.fontColor,
                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                    properties: this.properties,
                    style: PropertyFieldColorPickerStyle.Inline,
                    showPreview: true,
                    key: `fontColor-${this.selectedContainerIndex}`
                  }),
                  PropertyFieldColorPicker(`containers[${this.selectedContainerIndex}].backgroundColor`, {
                    label: "Background Color",
                    selectedColor: selectedContainer.backgroundColor,
                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                    properties: this.properties,
                    style: PropertyFieldColorPickerStyle.Inline,
                    showPreview: true,
                    key: `backgroundColor-${this.selectedContainerIndex}`
                  }),
                  PropertyFieldColorPicker(`containers[${this.selectedContainerIndex}].borderColor`, {
                    label: "Border Color",
                    selectedColor: selectedContainer.borderColor,
                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                    properties: this.properties,
                    style: PropertyFieldColorPickerStyle.Inline,
                    showPreview: true,
                    key: `borderColor-${this.selectedContainerIndex}`
                  }),
                  PropertyPaneSlider(`containers[${this.selectedContainerIndex}].borderRadius`, {
                    label: "Border Radius",
                    min: 0,
                    max: 50,
                    step: 1,
                    value: selectedContainer.borderRadius ? parseInt(selectedContainer.borderRadius, 10) : 0,
                    showValue: true,
                  }),
                ]
              }
            ]
          }
        ]
      };
    }

    return {
      pages: [
        {
          header: { description: "Select a Blurb to configure its properties" },
          groups: [
            {
              groupName: "Settings",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Description",
                  value: this.properties.description,
                }),
                PropertyPaneSlider('containerCount', {
                  label: "Number of Blurbs",
                  min: 1,
                  max: 4,
                  value: this.properties.containerCount,
                  showValue: true,
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
