import * as React from "react";
import WorksheetTools from "./WorksheetTools";
import { makeStyles } from "@fluentui/react-components";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC<AppProps> = () => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <WorksheetTools />
    </div>
  );
};

export default App;
