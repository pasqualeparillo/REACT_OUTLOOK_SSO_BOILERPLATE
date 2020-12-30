import * as React from "react";
import Test from "./test";
import Progress from "./Progress";
/* global Button, Header, HeroList, HeroListItem, Progress */

export default function App(props) {
  const { title, isOfficeInitialized } = props;
  if (!isOfficeInitialized) {
    return (
      <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
    );
  }

  return (
    <div className="ms-welcome">
      <Test />
    </div>
  );
}
