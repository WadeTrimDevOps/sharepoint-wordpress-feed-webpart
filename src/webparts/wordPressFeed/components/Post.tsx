import * as React from "react";
import styles from "./WordPressFeed.module.scss";
import { IWordPressPostAuthor, IPostComponent, IWordPressMediaItem } from "../interfaces";
import { Link } from "@fluentui/react";
import { extractDefaultColor } from "../util";
import { PropertyPaneDynamicFieldSet } from "@microsoft/sp-property-pane";

const Post: React.FC<IPostComponent> = ({ post, displaySettings, colorSettings }) => {
  const [media, setMedia] = React.useState<IWordPressMediaItem | undefined>(undefined);
  const [author, setAuthor] = React.useState<IWordPressPostAuthor | undefined>(undefined);
  React.useEffect(() => {
    if (post && post._embedded && post._embedded["wp:featuredmedia"] && post._embedded["wp:featuredmedia"].length > 0) {
      setMedia(post._embedded["wp:featuredmedia"][0]);
    }
  }, [displaySettings.showMedia, post?._embedded?.["wp:featuredmedia"]?.length, post?.id]);

  React.useEffect(() => {
    if (post && post._embedded && post._embedded.author && post._embedded.author.length > 0) {
      setAuthor(post._embedded.author[0]);
    }
  }, [displaySettings.showAuthor, post?._embedded?.author?.length]);

  const authorHasAllData: () => boolean = () => {
    return (
      author !== undefined &&
      author.name !== "" &&
      author.link !== "" &&
      author.avatar_urls !== undefined &&
      author.avatar_urls["24"] !== undefined &&
      author.avatar_urls["24"] !== ""
    );
  };

  const formattedDate = new Date(post.date).toLocaleDateString("en-US", {
    month: "long",
    day: "numeric",
    year: "numeric",
  });

  const extractText = (html: string): string => {
  if (!html) {
    return "";
  }
  const doc = new DOMParser().parseFromString(html, "text/html");
  const text = doc.body.textContent || "";
  const cleanedText = text
    .replace(/\[[^\]]+\]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (cleanedText.length <= displaySettings.excerptLength) {
    return cleanedText;
  }
  const truncated = cleanedText.slice(0, displaySettings.excerptLength);
  const lastSpace = truncated.lastIndexOf(" ");
  return `${truncated.slice(0, lastSpace > 0 ? lastSpace : truncated.length)}...`;
};

  const decodeHtml = (html: string): string => {
    const doc = new DOMParser().parseFromString(html, "text/html");
    return doc.documentElement.textContent || "";
  };

  const mediaHasAllData: () => boolean = () => {
    return (
      media !== undefined &&
      media.id !== undefined &&
      media.date !== undefined &&
      media.alt_text !== undefined &&
      media.source_url !== undefined &&
      media.caption !== undefined
    );
  };

  return (
    <article
      style={{ backgroundColor: extractDefaultColor(colorSettings.postBackgroundColor) }}
      className={displaySettings.layoutType === "grid" ? styles.gridPost : styles.listPost}
    >
      {displaySettings.showMedia && media && mediaHasAllData() && (
        <img
          id={media.id.toString()}
          className={styles.thumbnail}
          src={media.source_url}
          alt={decodeHtml(media.title.rendered)}
          style={{ height: displaySettings.mediaHeight, width: displaySettings.layoutType === "grid" ? "100%" : "auto" }}
        />
      )}
      <div className={styles.content}>
        {post.title && (
          <h2
            style={{
              color: extractDefaultColor(colorSettings.postTitleTextColor),
              fontSize: `${displaySettings.postFontSize}px`,
              fontWeight: Number(displaySettings.postFontWeight),
            }}
            className={styles.title}>
            {decodeHtml(post.title.rendered)}
          </h2>
        )}
        {displaySettings.showAuthor && author && authorHasAllData() && (
          <div className={styles.postAuthorContainer}>
            <img src={author.avatar_urls?.["24"]} alt={`avatar for ${author.name}`} />
            <p style={{ color: extractDefaultColor(colorSettings.postBodyTextColor) }}>
              Published by{" "}
              <Link href={author.link} target={"_blank"}>
                {decodeHtml(author.name)}
              </Link>
            </p>
          </div>
        )}
        <p style={{ color: extractDefaultColor(colorSettings.postBodyTextColor) }} className={styles.date}>
          {formattedDate}
        </p>
        {displaySettings.excerptLength !== 0 && (
          <p style={{ color: extractDefaultColor(colorSettings.postBodyTextColor) }} className={styles.excerpt}>
            {extractText(post.content?.rendered || post.excerpt?.rendered || "")}
          </p>
        )}
        <a
          style={{ color: extractDefaultColor(colorSettings.postReadMoreLinkTextColor) }}
          href={post.link}
          target="_blank"
          rel="noopener noreferrer"
          className={styles.readMore}
        >
          Read More →
        </a>
      </div>
    </article>
  );
};

export default Post;
