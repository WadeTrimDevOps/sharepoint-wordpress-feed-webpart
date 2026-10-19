import * as React from "react";
import { IFeedRenderProps } from "../interfaces";
import { Alert, MESSAGES } from "./Alert";
import PostsLayout from "./PostsLayout";
import styles from "./WordPressFeed.module.scss";
import { readMoreLinkNotEmpty } from "../util";
import ReadMoreLink from "./ReadMoreLink";
import { extractDefaultColor } from "../util";

const FeedRender: React.FC<IFeedRenderProps> = ({ posts, displaySettings, colorSettings }) => {
  return (
    <div
      style={{ backgroundColor: extractDefaultColor(colorSettings.mainBackgroundColor) }}
      className={styles.feedContainer}
    >
      {displaySettings?.title?.trim() !== "" && (
        <h2
          style={{
            color: extractDefaultColor(colorSettings.titleTextColor),
            fontSize: `${displaySettings.titleFontSize}px`,
fontWeight: Number(displaySettings.titleFontWeight),
          }}
          className={styles.feedTitle}
        >
          {displaySettings.title}
        </h2>
      )}
      {displaySettings.description?.trim() !== "" && (
        <p
          style={{ color: extractDefaultColor(colorSettings.descriptionTextColor) }}
          className={styles.feedDescription}
        >
          {displaySettings.description}
        </p>
      )}
      {!posts || posts.length === 0 ? (
        <Alert type={"warning"} msg={MESSAGES.WARNING.noPostsToDisplay} />
      ) : (
        <PostsLayout posts={posts} displaySettings={displaySettings} colorSettings={colorSettings} />
      )}
      {displaySettings.readMoreLink.include && readMoreLinkNotEmpty(displaySettings.readMoreLink) && (
        <ReadMoreLink {...displaySettings.readMoreLink} colorSettings={colorSettings} />
      )}
    </div>
  );
};

export default FeedRender;
