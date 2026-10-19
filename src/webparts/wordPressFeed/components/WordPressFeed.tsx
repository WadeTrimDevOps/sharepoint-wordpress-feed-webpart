/**
 * SharePoint Framework (SPFx) Web Part to fetch and display a WordPress feed
 * with filtering by tag, category, or post name pattern.
 */

import * as React from "react";
import { useEffect } from "react";
import { IWordPressPost, IWordPressFeedComponentProps } from "../interfaces";
import { fetchPosts, validateUrl } from "../util";
import FeedRender from "./FeedRender";
import { MESSAGES, Alert } from "./Alert";

const WordPressFeed: React.FC<IWordPressFeedComponentProps> = ({
  feedFilterSettings,
  displaySettings,
  colorSettings,
  url,
  feedType,
  refreshToken,
}) => {
  const selectedFeedType = feedType ?? "news";
  const [posts, setPosts] = React.useState<Array<IWordPressPost>>([]);
  const [status, setStatus] = React.useState<"idle" | "loading" | "success" | "error">("loading");

  useEffect(() => {
    if (!validateUrl(url) || !feedFilterSettings) {
      setStatus("idle");
      return;
    }

    let isCurrentRequest = true;
    setStatus("loading");

    fetchPosts(url, feedFilterSettings, selectedFeedType)
      .then((nextPosts) => {
        if (isCurrentRequest) {
          setPosts(nextPosts);
          setStatus("success");
        }
      })
      .catch((e) => {
        if (isCurrentRequest) {
          console.error(e);
          setStatus("error");
        }
      });

    return () => {
      isCurrentRequest = false;
    };
  }, [
    url,
    selectedFeedType,
    refreshToken,
    feedFilterSettings?.filterJoinOperator,
    feedFilterSettings?.numPosts,
    feedFilterSettings?.pastDays,
    feedFilterSettings?.postPattern,
    feedFilterSettings?.categoryIds.join(","),
    feedFilterSettings?.tagIds.join(","),
  ]);

  if (!validateUrl(url)) {
    return <Alert msg={MESSAGES.WARNING.provideSiteUrl} type={"warning"} />;
  }

  if (status === "loading") {
    return <Alert msg={MESSAGES.WARNING.postsLoading} type={"warning"} />;
  }

  if (status === "error") {
    return <Alert msg={MESSAGES.WARNING.failedToFetchFeed} type={"error"} />;
  }

  if (status === "success" && feedFilterSettings && displaySettings) {
    return <FeedRender posts={posts} displaySettings={displaySettings} colorSettings={colorSettings} />;
  }

  return <Alert msg={MESSAGES.WARNING.noPostsToDisplay} type={"warning"} />;
};

export default WordPressFeed;
