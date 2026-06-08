import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { PropertyPaneTagCategoryPicker } from "./controls/TagCategoryPicker/PropertyPaneTagCategoryPicker";
import {
  type IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import WordPressFeed from "./components/WordPressFeed";
import { ITagOrCategory, IWordPressFeedWebPartProps } from "./interfaces";
import { defaultColorSettings, defaultDisplaySettings, defaultFeedRequestSettings } from "./defaults";
import { filterJoinOperators, layouts } from "./dropdownOptions";
import { PropertyPaneWordPressUrlField } from "./controls/WordPressUrlField/PropertyPaneWordPressUrlField";
import { getColorDropdownOptions, validateUrl } from "./util";
import { PropertyPaneNumericTextField } from "./controls/NumericTextField/PropertyPaneNumericTextField";
import { PropertyPaneThemePaletteColorPicker } from "./controls/ThemePaletteColorPicker/PropertyPaneNumericTextField";
import { IDropdownOption } from "@fluentui/react";
export default class WordPressFeedWebPart extends BaseClientSideWebPart<IWordPressFeedWebPartProps> {
  // store data that you need to fetch asynchronously within
  // configuration panel as private variables, and load them
  // onInit.
  private _tagOptions: { key: number; text: string }[] = [];
  private _categoryOptions: { key: number; text: string }[] = [];
  private _colorOptions: IPropertyPaneDropdownOption[] = [];
  private _siteName: string = "";
  public render(): void {
    const element: React.ReactElement<IWordPressFeedWebPartProps> = React.createElement(WordPressFeed, {
      displaySettings: this.properties.displaySettings,
      feedFilterSettings: this.properties.feedFilterSettings,
      colorSettings: this.properties.colorSettings,
      url: this.properties.url,
    });
    ReactDom.render(element, this.domElement);
  }
  private loadSiteDataFromURL: () => Promise<void> = async () => {
    const apiBase = this.properties.url;
    const [general, tags, categories] = await Promise.all([
      fetch(`${apiBase}/wp-json`).then((res) => res.json()),
      // per_page ensures we're pulling all categories instead of the default (10)
      fetch(`${apiBase}/wp-json/wp/v2/tags?per_page=100`).then((res) => res.json()),
      fetch(`${apiBase}/wp-json/wp/v2/categories?per_page=100`).then((res) => res.json()),
    ]);
    this._siteName = general.name;
    this._tagOptions = tags.map((tag: ITagOrCategory) => ({ key: tag.id, text: tag.name }));
    this._categoryOptions = categories.map((cat: ITagOrCategory) => ({ key: cat.id, text: cat.name }));
    return Promise.resolve();
  };
  protected onInit(): Promise<void> {
    // initialize values to defaults if database instance has nothing stored
    if (!this.properties.url) {
      this.properties.url = "";
    }
    if (!this.properties.feedFilterSettings) {
      this.properties.feedFilterSettings = defaultFeedRequestSettings;
    }
    if (!this.properties.displaySettings) {
      this.properties.displaySettings = defaultDisplaySettings;
    }
    if (!this.properties.colorSettings) {
      this.properties.colorSettings = defaultColorSettings;
    }
    if (validateUrl(this.properties.url)) {
      this.loadSiteDataFromURL().catch((e) => console.error(e));
    }
    this._colorOptions = getColorDropdownOptions();
    return super.onInit();
  }
  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty("--bodyText", semanticColors.bodyText || null);
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty("--linkHovered", semanticColors.linkHovered || null);
    }
  }
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
  public getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Configure your WordPress feed" },
          groups: [
            {
              groupName: "WordPress Site Connection",
              groupFields: [
                new PropertyPaneWordPressUrlField("url", {
                  label: "WordPress Site URL",
                  key: "url",
                  value: this.properties.url,
                  siteName: this._siteName,
                  onChange: (s: string) => {
                    this.properties.url = s;
                    this.context.propertyPane.refresh();
                  },
                  siteFetchHandler: async () => {
                    await this.loadSiteDataFromURL()
                      .then(() => {
                        this.context.propertyPane.refresh();
                        this.render();
                      })
                      .catch((e) => console.error(e));
                  },
                }),
              ],
            },
            {
              groupName: "Display Settings",
              groupFields: [
                PropertyPaneTextField("displaySettings.title", {
                  label: "Feed Title",
                  value: this.properties.displaySettings.title,
                }),
                PropertyPaneTextField("displaySettings.description", {
                  label: "Feed Description",
                  multiline: true,
                  value: this.properties.displaySettings.description,
                }),
                PropertyPaneToggle("displaySettings.readMoreLink.include", {
                  label: "Include Link Below Feed?",
                  checked: this.properties.displaySettings.readMoreLink.include,
                }),
                PropertyPaneTextField("displaySettings.readMoreLink.linkText", {
                  label: "Link Text",
                  disabled: !this.properties.displaySettings.readMoreLink.include,
                  value: this.properties.displaySettings.readMoreLink.linkText,
                }),
                PropertyPaneTextField("displaySettings.readMoreLink.linkUrl", {
                  label: "Link URL",
                  disabled: !this.properties.displaySettings.readMoreLink.include,
                  value: this.properties.displaySettings.readMoreLink.linkUrl,
                }),
                PropertyPaneToggle("displaySettings.readMoreLink.linkNewTab", {
                  label: "Open link in new tab",
                  checked: this.properties.displaySettings.readMoreLink.linkNewTab,
                }),
                PropertyPaneToggle("displaySettings.showAuthor", {
                  label: "Show Author on Posts",
                  checked: this.properties.displaySettings.showAuthor,
                }),
                PropertyPaneToggle("displaySettings.showMedia", {
                  label: "Show Media / Thumbnail (if available)",
                  checked: this.properties.displaySettings.showMedia,
                }),
                PropertyPaneDropdown("displaySettings.layoutType", {
                  label: "Feed Layout",
                  options: layouts,
                  selectedKey: this.properties.displaySettings.layoutType,
                }),
                new PropertyPaneNumericTextField("displaySettings.itemsPerPage", {
                  label: "Pagination - Items Per Page (set equal to Number of Posts for no pagination)",
                  min: 1,
                  max: 100, // obviously the real limit will be the number of posts in the feed
                  key: "displaySettings.itemsPerPage",
                  value: this.properties.displaySettings.itemsPerPage.toString(),
                  onChange: (s: string) => {
                    this.properties.displaySettings.itemsPerPage = parseInt(s);
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                new PropertyPaneNumericTextField("displaySettings.excerptLength", {
                  label: "Excerpt Preview Length (set to 0 to hide excerpt)",
                  min: 0,
                  max: 300,
                  key: "displaySettings.excerptLength",
                  value: this.properties.displaySettings.excerptLength.toString(),
                  onChange: (s: string) => {
                    this.properties.displaySettings.excerptLength = parseInt(s);
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
              ],
            },
            {
              groupName: "Color Settings",
              groupFields: [
                new PropertyPaneThemePaletteColorPicker("colorSettings.mainBackgroundColor", {
                  options: this._colorOptions,
                  selectedKey: this.properties.colorSettings.mainBackgroundColor,
                  label: "Main Background Color",
                  key: "colorSettings.mainBackgroundColor",
                  onChange: (newSelectedOption: IDropdownOption) => {
                    this.properties.colorSettings.mainBackgroundColor = newSelectedOption.key as string;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                new PropertyPaneThemePaletteColorPicker("colorSettings.titleTextColor", {
                  options: this._colorOptions,
                  selectedKey: this.properties.colorSettings.titleTextColor,
                  label: "Title Text Color",
                  key: "colorSettings.titleTextColor",
                  onChange: (newSelectedOption: IDropdownOption) => {
                    this.properties.colorSettings.titleTextColor = newSelectedOption.key as string;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                new PropertyPaneThemePaletteColorPicker("colorSettings.descriptionTextColor", {
                  options: this._colorOptions,
                  selectedKey: this.properties.colorSettings.descriptionTextColor,
                  label: "Description Text Color",
                  key: "colorSettings.descriptionTextColor",
                  onChange: (newSelectedOption: IDropdownOption) => {
                    this.properties.colorSettings.descriptionTextColor = newSelectedOption.key as string;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                new PropertyPaneThemePaletteColorPicker("colorSettings.postBackgroundColor", {
                  options: this._colorOptions,
                  selectedKey: this.properties.colorSettings.postBackgroundColor,
                  label: "Post Background Color",
                  key: "colorSettings.postBackgroundColor",
                  onChange: (newSelectedOption: IDropdownOption) => {
                    this.properties.colorSettings.postBackgroundColor = newSelectedOption.key as string;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                new PropertyPaneThemePaletteColorPicker("colorSettings.postTitleTextColor", {
                  options: this._colorOptions,
                  selectedKey: this.properties.colorSettings.postTitleTextColor,
                  label: "Post Title Text Color",
                  key: "colorSettings.postTitleTextColor",
                  onChange: (newSelectedOption: IDropdownOption) => {
                    this.properties.colorSettings.postTitleTextColor = newSelectedOption.key as string;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                new PropertyPaneThemePaletteColorPicker("colorSettings.postBodyTextColor", {
                  options: this._colorOptions,
                  selectedKey: this.properties.colorSettings.postBodyTextColor,
                  label: "Post Body Text Color",
                  key: "colorSettings.postBodyTextColor",
                  onChange: (newSelectedOption: IDropdownOption) => {
                    this.properties.colorSettings.postBodyTextColor = newSelectedOption.key as string;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                new PropertyPaneThemePaletteColorPicker("colorSettings.postReadMoreLinkTextColor", {
                  options: this._colorOptions,
                  selectedKey: this.properties.colorSettings.postReadMoreLinkTextColor,
                  label: "Post Read More Link Text Color",
                  key: "colorSettings.postReadMoreLinkTextColor",
                  onChange: (newSelectedOption: IDropdownOption) => {
                    this.properties.colorSettings.postReadMoreLinkTextColor = newSelectedOption.key as string;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                new PropertyPaneThemePaletteColorPicker("colorSettings.pageButtonBackgroundColor", {
                  options: this._colorOptions,
                  selectedKey: this.properties.colorSettings.pageButtonBackgroundColor,
                  label: "Page Button Background Color",
                  key: "colorSettings.pageButtonBackgroundColor",
                  onChange: (newSelectedOption: IDropdownOption) => {
                    this.properties.colorSettings.pageButtonBackgroundColor = newSelectedOption.key as string;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                new PropertyPaneThemePaletteColorPicker("colorSettings.pageButtonTextColor", {
                  options: this._colorOptions,
                  selectedKey: this.properties.colorSettings.pageButtonTextColor,
                  label: "Page Button Text Color",
                  key: "colorSettings.pageButtonTextColor",
                  onChange: (newSelectedOption: IDropdownOption) => {
                    this.properties.colorSettings.pageButtonTextColor = newSelectedOption.key as string;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                new PropertyPaneThemePaletteColorPicker("colorSettings.pageButtonHoverBackgroundColor", {
                  options: this._colorOptions,
                  selectedKey: this.properties.colorSettings.pageButtonHoverBackgroundColor,
                  label: "Page Button Hover Background Color",
                  key: "colorSettings.pageButtonHoverBackgroundColor",
                  onChange: (newSelectedOption: IDropdownOption) => {
                    this.properties.colorSettings.pageButtonHoverBackgroundColor = newSelectedOption.key as string;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                new PropertyPaneThemePaletteColorPicker("colorSettings.pageButtonHoverTextColor", {
                  options: this._colorOptions,
                  selectedKey: this.properties.colorSettings.pageButtonHoverTextColor,
                  label: "Page Button Hover Text Color",
                  key: "colorSettings.pageButtonHoverTextColor",
                  onChange: (newSelectedOption: IDropdownOption) => {
                    this.properties.colorSettings.pageButtonHoverTextColor = newSelectedOption.key as string;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                new PropertyPaneThemePaletteColorPicker("colorSettings.feedReadMoreButtonBackgroundColor", {
                  options: this._colorOptions,
                  selectedKey: this.properties.colorSettings.feedReadMoreButtonBackgroundColor,
                  label: "(Bottom of feed) Read More Button Background Color",
                  key: "colorSettings.feedReadMoreButtonBackgroundColor",
                  onChange: (newSelectedOption: IDropdownOption) => {
                    this.properties.colorSettings.feedReadMoreButtonBackgroundColor = newSelectedOption.key as string;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                new PropertyPaneThemePaletteColorPicker("colorSettings.feedReadMoreButtonTextColor", {
                  options: this._colorOptions,
                  selectedKey: this.properties.colorSettings.feedReadMoreButtonTextColor,
                  label: "(Bottom of feed) Read More Button Text Color",
                  key: "colorSettings.feedReadMoreButtonTextColor",
                  onChange: (newSelectedOption: IDropdownOption) => {
                    this.properties.colorSettings.feedReadMoreButtonTextColor = newSelectedOption.key as string;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
              ],
            },
            {
              groupName: "Feed Filter Settings",
              groupFields: [
                new PropertyPaneNumericTextField("feedFilterSettings.numPosts", {
                  label: "Number of Posts",
                  min: 1,
                  max: 5000, // realistically won't need this high of a max, but should not hurt to have.
                  key: "feedFilterSettings.numPosts",
                  value: this.properties.feedFilterSettings.numPosts.toString(),
                  onChange: (s: string) => {
                    this.properties.feedFilterSettings.numPosts = parseInt(s);
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                new PropertyPaneNumericTextField("feedFilterSettings.pastDays", {
                  label: "Past Days",
                  min: 1,
                  max: 36500, // 100 years seems sufficient
                  key: "feedFilterSettings.pastDays",
                  value: this.properties.feedFilterSettings.pastDays.toString(),
                  onChange: (s: string) => {
                    this.properties.feedFilterSettings.pastDays = parseInt(s);
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                // Tag/category filtering would require dynamic loading,
                // which SPFX doesn't support directly in the property pane.
                // I am using a customPropertyPaneField for this if needed.
                new PropertyPaneTagCategoryPicker("feedFilterSettings.tagIds", {
                  label: "Select Tags",
                  key: "feedFilterSettings.tagIds",
                  selectedKeys: this.properties.feedFilterSettings.tagIds || [],
                  options: this._tagOptions,
                  onChange: (newTagIds: number[]) => {
                    this.properties.feedFilterSettings.tagIds = newTagIds;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                new PropertyPaneTagCategoryPicker("feedFilterSettings.categoryIds", {
                  label: "Select Categories",
                  key: "feedFilterSettings.categoryIds",
                  selectedKeys: this.properties.feedFilterSettings.categoryIds || [],
                  options: this._categoryOptions,
                  onChange: (newCategoryIds: number[]) => {
                    this.properties.feedFilterSettings.categoryIds = newCategoryIds;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
                PropertyPaneTextField("feedFilterSettings.postPattern", {
                  label: "Post Pattern (Wildcard OK)",
                  value: this.properties.feedFilterSettings.postPattern,
                }),
                PropertyPaneDropdown("feedFilterSettings.filterJoinOperator", {
                  label: "Join Operator",
                  options: filterJoinOperators,
                  selectedKey: this.properties.feedFilterSettings.filterJoinOperator,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
