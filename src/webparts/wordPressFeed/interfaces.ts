interface IWordPressPostAuthor {
  id: number;
  name: string; // full name
  link: string; // link to author on site
  slug: string;
  avatar_urls:
    | {
        "24": string;
      }
    | undefined;
}

interface IWordPressPostEmbeddedData {
  author: Array<IWordPressPostAuthor>;
  "wp:featuredmedia": Array<IWordPressMediaItem>;
}

interface IWordPressPost {
  id: number;
  title: { rendered: string };
  excerpt: { rendered: string };
  _embedded: IWordPressPostEmbeddedData;
  link: string;
  date: string;
  layoutType?: "grid" | "list"; // Optional layoutType prop
}

interface IPostsLayout {
  posts: Array<IWordPressPost>;
  displaySettings: IDisplaySettings;
  colorSettings: IColorSettings;
}

interface IPostComponent {
  post: IWordPressPost;
  displaySettings: IDisplaySettings;
  colorSettings: IColorSettings;
}

interface ITagOrCategory {
  id: number;
  name: string;
}

interface IColorSettings {
  mainBackgroundColor: string;
  titleTextColor: string;
  descriptionTextColor: string;
  postBackgroundColor: string;
  postTitleTextColor: string;
  postBodyTextColor: string; //excerpt, author, date
  postReadMoreLinkTextColor: string;
  pageButtonBackgroundColor: string;
  pageButtonTextColor: string;
  pageButtonHoverBackgroundColor: string;
  pageButtonHoverTextColor: string;
  feedReadMoreButtonBackgroundColor: string;
  feedReadMoreButtonTextColor: string;
}

interface IDisplaySettings {
  title?: string;
  titleFontSize: number;
  titleFontWeight: string;
  description?: string;
  postFontSize: number;
  postFontWeight: string;
  readMoreLink: IReadMoreLink;
  layoutType: "list" | "grid";
  showAuthor: boolean;
  mediaHeight: number;
  showMedia: boolean;
  itemsPerPage: number;
  excerptLength: number;
}

interface IWordPressFeedFilterSettings {
  numPosts: number;
  pastDays: number;
  postPattern: string;
  filterJoinOperator: string;
  tagIds: Array<number>;
  categoryIds: Array<number>;
}

type WordPressFeedType = "news" | "projectHighlights";

interface IReadMoreLink {
  linkUrl: string;
  linkText: string;
  linkNewTab: boolean;
  include: boolean;
  colorSettings: IColorSettings;
}

interface IWordPressFeedWebPartProps {
  feedFilterSettings: IWordPressFeedFilterSettings;
  displaySettings: IDisplaySettings;
  colorSettings: IColorSettings;
  url: string;
  feedType: WordPressFeedType;
}

interface IFeedRenderProps {
  displaySettings: IDisplaySettings;
  colorSettings: IColorSettings;
  posts: Array<IWordPressPost>;
}

// wordpress media items with ?embed=true in the wp:featuredmedia array
interface IWordPressMediaTitle {
  rendered: string;
}

interface IWordPressMediaCaption {
  rendered: string;
}

interface IWordPressMediaSize {
  file: string;
  width: number;
  height: number;
  filesize: number;
  mime_type: string;
  source_url: string;
}

interface IWordPressMediaImageMeta {
  aperture: string;
  credit: string;
  camera: string;
  caption: string;
  created_timestamp: string;
  copyright: string;
  focal_length: string;
  iso: string;
  shutter_speed: string;
  title: string;
  orientation: string;
  keywords: string[];
  resized_images: Record<string, string>;
}

interface IWordPressMediaDetails {
  width: number;
  height: number;
  file: string;
  filesize: number;
  sizes: Record<string, IWordPressMediaSize>;
  image_meta: IWordPressMediaImageMeta;
}

interface IWordPressMediaLink {
  href: string;
  targetHints?: {
    allow: string[];
  };
  embeddable?: boolean;
}

interface IWordPressMediaLinks {
  self: IWordPressMediaLink[];
  collection: IWordPressMediaLink[];
  about: IWordPressMediaLink[];
  author: IWordPressMediaLink[];
  replies: IWordPressMediaLink[];
}

interface IWordPressMediaItem {
  id: number;
  date: string;
  slug: string;
  type: string;
  link: string;
  title: IWordPressMediaTitle;
  author: number;
  featured_media: number;
  caption: IWordPressMediaCaption;
  alt_text: string;
  media_type: string;
  mime_type: string;
  media_details: IWordPressMediaDetails;
  source_url: string;
  _links: IWordPressMediaLinks;
}

interface IPagination {
  totalPages: number;
  currentPage: number;
  paginate: (pageNumber: number) => void;
  colorSettings: IColorSettings;
}
export {
  IPagination,
  IFeedRenderProps,
  IReadMoreLink,
  IWordPressFeedWebPartProps,
  IWordPressPost,
  IPostComponent,
  ITagOrCategory,
  IWordPressFeedFilterSettings,
  IWordPressMediaItem,
  IWordPressPostAuthor,
  IWordPressPostEmbeddedData,
  IDisplaySettings,
  IPostsLayout,
  IColorSettings,
  WordPressFeedType,
};
