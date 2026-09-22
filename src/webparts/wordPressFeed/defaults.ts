import { IWordPressFeedFilterSettings, IReadMoreLink, IDisplaySettings } from "./interfaces";

export const defaultFeedRequestSettings = {
  numPosts: 1000,
  pastDays: 3000,
  postPattern: "",
  filterJoinOperator: "AND",
  categoryIds: [],
  tagIds: [],
} as IWordPressFeedFilterSettings;

export const defaultDisplaySettings = {
  layoutType: "list",
  showAuthor: false,
  showMedia: true,
  itemsPerPage: 4,
  excerptLength: 0,
  title: "Practice News",
  titleFontSize: 20,
  titleFontWeight: "400",
  description: "Read the latest from this Practice.",
  readMoreLink: {
    linkText: "Read more:",
    linkUrl: "https://wadetrim.com",
    linkNewTab: false,
  } as IReadMoreLink,
} as IDisplaySettings;

export const defaultColorSettings = {
  mainBackgroundColor: "[theme:white, default: #fefefe]", // Very light background
  titleTextColor: "[theme:neutralDark, default: #201f1e]", // Dark text for contrast
  descriptionTextColor: "[theme:neutralPrimary, default: #323130]", // Darker text
  postBackgroundColor: "[theme:neutralLighter, default: #f4f4f4]", // Light background for posts
  postTitleTextColor: "[theme:neutralDark, default: #201f1e]", // Dark post title
  postBodyTextColor: "[theme:neutralPrimaryAlt, default: #3b3a39]", // Darker post body text
  postReadMoreLinkTextColor: "[theme:wtPrimary, default: #43B02A]", // Distinctive link color
  pageButtonBackgroundColor: "[theme:wtPrimary, default: #43B02A]", // Primary theme color for button
  pageButtonTextColor: "[theme:white, default: #fefefe]", // Light text on dark button
  pageButtonHoverBackgroundColor: "[theme:wtPrimaryHover, default: #369126]", // Slightly darker hover
  pageButtonHoverTextColor: "[theme:white, default: #fefefe]", // Light text on hover

  feedReadMoreButtonBackgroundColor: "[theme:wtPrimary, default: #43B02A]", // Primary theme color for button
  feedReadMoreButtonTextColor: "[theme:white, default: #fefefe]", // Light text on dark button
};
