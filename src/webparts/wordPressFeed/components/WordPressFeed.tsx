/**
 * SharePoint Framework (SPFx) Web Part to fetch and display a WordPress feed
 * with filtering by tag, category, or post name pattern.
 */

import * as React from "react";
import { useEffect } from "react";
import { IWordPressPost, IWordPressFeedWebPartProps } from "../interfaces";
import { fetchPosts, validateUrl } from "../util";
import FeedRender from "./FeedRender";
import { MESSAGES, Alert } from "./Alert";

const WordPressFeed: React.FC<IWordPressFeedWebPartProps> = ({
  feedFilterSettings,
  displaySettings,
  colorSettings,
  url,
  feedType,
}) => {
  const [posts, setPosts] = React.useState<Array<IWordPressPost>>([]);

  const updatePosts: () => void = () => {
    fetchPosts(url, feedFilterSettings, feedType)
      .then((posts) => {
        setPosts(posts);
      })
      .catch((e) => {
        console.error(e);
      });
  };

  useEffect(() => {
    if (validateUrl(url) && feedFilterSettings) {
      updatePosts();
    }
  }, [
    url,
    feedType,
    feedFilterSettings.filterJoinOperator,
    feedFilterSettings.numPosts,
    feedFilterSettings.pastDays,
    feedFilterSettings.postPattern,
    feedFilterSettings.categoryIds.join(","),
    feedFilterSettings.tagIds.join(","),
  ]);

  return url && posts && feedFilterSettings && displaySettings ? (
    <FeedRender posts={posts} displaySettings={displaySettings} colorSettings={colorSettings} />
  ) : (
    <Alert msg={MESSAGES.WARNING.noPostsToDisplay} type={"warning"} />
  );
};

export default WordPressFeed;
