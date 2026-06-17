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

const MAX_WORDPRESS_POSTS_PER_PAGE = 100;
const MAX_WORDPRESS_PAGES = 50;

const getPaginatedFetchUrl: (fetchUrl: string, page: number) => string = (fetchUrl, page) => {
  const url = new URL(fetchUrl);
  url.searchParams.set("per_page", MAX_WORDPRESS_POSTS_PER_PAGE.toString());
  url.searchParams.set("page", page.toString());
  url.searchParams.set("_embed", "true");
  return url.toString();
};

const fetchPaginatedPosts: (fetchUrl: string, maxPosts: number) => Promise<Array<IWordPressPost>> = async (
  fetchUrl,
  maxPosts,
) => {
  const posts: IWordPressPost[] = [];
  let page = 1;
  while (posts.length < maxPosts && page <= MAX_WORDPRESS_PAGES) {
    const response = await fetch(getPaginatedFetchUrl(fetchUrl, page));
    if (!response.ok) {
      throw new Error(`Error fetching posts: ${response.status} ${response.statusText}`);
    }
    const pagePosts = (await response.json()) as Array<IWordPressPost>;
    if (!Array.isArray(pagePosts) || pagePosts.length === 0) {
      break;
    }
    posts.push(...pagePosts);
    if (pagePosts.length < MAX_WORDPRESS_POSTS_PER_PAGE) {
      break;
    }
    page += 1;
  }
  return posts.slice(0, maxPosts);
};

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
    return await fetchPaginatedPosts(fetchUrl, settings.numPosts);
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
      withTags.push(...(await fetchPaginatedPosts(tagsFilterUrl, settings.numPosts)));
    }

    if (settings.categoryIds.length > 0) {
      categoriesFilterUrl += `&categories=${settings.categoryIds.join(",")}`;
      withCategories.push(...(await fetchPaginatedPosts(categoriesFilterUrl, settings.numPosts)));
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
