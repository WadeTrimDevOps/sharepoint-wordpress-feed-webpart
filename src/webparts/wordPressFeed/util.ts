import { IPropertyPaneDropdownOption } from "@microsoft/sp-property-pane";
import {
  IWordPressPost,
  IWordPressFeedFilterSettings,
  IReadMoreLink,
  IWordPressPostAuthor,
  IWordPressMediaItem,
  IWordPressPostEmbeddedData,
} from "./interfaces";
import { colorPalette } from "./colorPalette";

const fetchPostsWithAndFilters: (
  fetchUrl: string,
  settings: IWordPressFeedFilterSettings,
) => Promise<Array<IWordPressPost>> = async (fetchUrl, settings) => {
  try {
    if (settings.tagIds.length > 0) fetchUrl += `&tags=${settings.tagIds.join(",")}`;
    if (settings.categoryIds.length > 0) fetchUrl += `&categories=${settings.categoryIds.join(",")}`;
    if (settings.pastDays) {
      const sinceDate = new Date();
      sinceDate.setDate(sinceDate.getDate() - settings.pastDays);
      fetchUrl += `&after=${sinceDate.toISOString()}`;
    }
    const response = await fetch(fetchUrl);
    return await response.json();
  } catch (error) {
    if (error instanceof Error) {
      throw new Error("Error fetching posts: " + error.message);
    }
    throw new Error("Error fetching posts: unknown error");
  }
};

const fetchPostsWithOrFilters: (
  fetchUrl: string,
  settings: IWordPressFeedFilterSettings,
) => Promise<Array<IWordPressPost>> = async (fetchUrl, settings) => {
  try {
    if (settings.pastDays) {
      const sinceDate = new Date();
      sinceDate.setDate(sinceDate.getDate() - settings.pastDays);
      fetchUrl += `&after=${sinceDate.toISOString()}`;
    }
    const withTags: IWordPressPost[] = [];
    const withCategories: IWordPressPost[] = [];
    let tagsFilterUrl = fetchUrl;
    let categoriesFilterUrl = fetchUrl;
    // get post set matching tags filter
    if (settings.tagIds.length > 0) {
      tagsFilterUrl += `&tags=${settings.tagIds.join(",")}`;

      const tagFilterResponse = await fetch(tagsFilterUrl);
      withTags.push(...(await tagFilterResponse.json()));
    }

    if (settings.categoryIds.length > 0) {
      categoriesFilterUrl += `&categories=${settings.categoryIds.join(",")}`;

      const categoryFilterResponse = await fetch(categoriesFilterUrl);
      withCategories.push(...(await categoryFilterResponse.json()));
    }

    const dedupedUnion = [...withTags, ...withCategories].reduce<IWordPressPost[]>((accum, post) => {
      // Check if the post is already in the accumulator based on the post id
      if (!accum.some((item) => item.id === post.id)) {
        accum.push(post); // Add unique post to the accumulator
      }
      return accum;
    }, []);

    return dedupedUnion;
  } catch (error) {
    console.error("FAILED");
    if (error instanceof Error) {
      throw new Error("Error fetching posts: " + error.message);
    }
    throw new Error("Error fetching posts: unknown error");
  }
};

const getAuthorIfPresent: (embedded: IWordPressPostEmbeddedData) => IWordPressPostAuthor | undefined = (embedded) => {
  if (embedded.author && embedded.author.length > 0) {
    const author = embedded.author[0];
    return {
      id: author.id ? author.id : 0,
      name: author.name ? author.name : "",
      slug: author.slug ? author.slug : "",
      link: author.link ? author.link : "",
      avatar_urls: author.avatar_urls && "24" in author.avatar_urls ? author.avatar_urls : undefined,
    };
  }
  return undefined;
};

const fetchPosts: (url: string, settings: IWordPressFeedFilterSettings) => Promise<Array<IWordPressPost>> = async (
  url,
  settings,
) => {
  let posts: Array<IWordPressPost> = [];
  if (!url || url === "") return posts;
  try {
    const fetchUrl = `${url}/wp-json/wp/v2/posts?per_page=${settings.numPosts}&_embed=true`;
    if (settings.filterJoinOperator === "AND") {
      posts = await fetchPostsWithAndFilters(fetchUrl, settings);
    } else {
      posts = await fetchPostsWithOrFilters(fetchUrl, settings);
    }

    // filter with regex pattern and sort most to least recent
    return posts
      .filter((post) => new RegExp(settings.postPattern, "i").test(post.title.rendered))
      .slice()
      .sort((a, b) => {
        // most recent first
        const dateA = new Date(a.date);
        const dateB = new Date(b.date);
        return dateB.getTime() - dateA.getTime();
      })
      .slice(0, settings.numPosts);
  } catch (e) {
    console.error("FAILED");
    console.error(e);
    if (e instanceof Error) {
      throw new Error(e.message);
    }
    throw new Error("Error fetching posts: unknown error");
  }
};

const validateUrl: (url: string) => boolean = (url) => {
  let valid = false;
  if (url && url !== "") {
    const urlRegex = /^(https?:\/\/)?(www\.)?([a-zA-Z0-9.-]+)\.([a-z]{2,})\/?$/;
    valid = urlRegex.test(url);
  }
  return valid;
};

const readMoreLinkNotEmpty: (readMoreLink: IReadMoreLink) => boolean = (readMoreLink) => {
  return readMoreLink && readMoreLink.linkText.trim() !== "" && readMoreLink.linkUrl.trim() !== "";
};

function getColorDropdownOptions(): IPropertyPaneDropdownOption[] {
  return Object.keys(colorPalette).map((key) =>
    ({
      key: `[theme:${key}, default: ${colorPalette[key as keyof typeof colorPalette]}]`,
      text: key,
    } as IPropertyPaneDropdownOption),
  );
}

function extractDefaultColor(themeString: string): string {
  if (!themeString) {
    return "";
  }
  const parts = themeString.split("default:");
  if (parts.length > 1) {
    return parts[1].replace("]", "").trim();
  }
  return themeString;
}

const getPostFeaturedMediaIfPresent: (featuredMedia: Array<IWordPressMediaItem>) => IWordPressMediaItem | undefined = (
  featuredMedia,
) => {
  return featuredMedia && featuredMedia.length > 0 ? featuredMedia[0] : undefined;
};

export {
  extractDefaultColor,
  getColorDropdownOptions,
  readMoreLinkNotEmpty,
  fetchPostsWithAndFilters,
  fetchPostsWithOrFilters,
  fetchPosts,
  validateUrl,
  getAuthorIfPresent,
  getPostFeaturedMediaIfPresent,
};
