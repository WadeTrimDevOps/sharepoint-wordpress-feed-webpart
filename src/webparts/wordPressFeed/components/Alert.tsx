import * as React from "react";
import styles from "./WordPressFeed.module.scss";

interface AlertProps {
  msg: string;
  type: "success" | "warning" | "error";
}
const MESSAGES = {
  ERROR: {
    failedSaveSettings: "Failed to save settings. Please try modifying your settings and resaving.",
  },
  SUCCESS: {
    connectedToSite: "Connected to site",
    savedSettings: "Settings saved successfully.",
  },
  WARNING: {
    noPostsToDisplay: "There are no posts to display currently",
    postsLoading: "Posts loading...",
    provideSiteUrl: "Please provide a valid site URL.",
    failedSiteInfoFetch: "Failed to fetch site information. Please try updating your site URL.",
    failedToFetchFeed: "Failed to fetch the feed. Please try updating your site URL.",
  },
};
const Alert: (props: AlertProps) => JSX.Element = (props: AlertProps) => {
  const classMap = {
    success: styles.success,
    warning: styles.warning,
    error: styles.error,
  };
  return (
    <div className={`${styles.alert} ${classMap[props.type]}`}>
      <span>{props.msg}</span>
    </div>
  );
};

export { MESSAGES, Alert };
