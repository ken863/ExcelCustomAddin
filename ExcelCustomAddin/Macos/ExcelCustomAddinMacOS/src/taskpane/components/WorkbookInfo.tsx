import * as React from "react";
import {
  makeStyles,
  Card,
  CardHeader,
  Text,
  Button,
  tokens,
} from "@fluentui/react-components";
import {
  Info24Regular,
  ArrowClockwise24Regular,
} from "@fluentui/react-icons";
import StorageService from "../services/StorageService";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  statsGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalS,
  },
  statItem: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },
  statValue: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorBrandForeground1,
  },
  statLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  refreshButton: {
    alignSelf: "flex-end",
  },
});

interface WorkbookStats {
  totalSheets: number;
  pinnedSheets: number;
  activeSheet: string;
  hasUnsavedChanges: boolean;
  lastModified?: string;
}

interface WorkbookInfoProps {
  onRefresh: () => void;
}

const WorkbookInfo: React.FC<WorkbookInfoProps> = ({ onRefresh }) => {
  const styles = useStyles();
  const [stats, setStats] = React.useState<WorkbookStats>({
    totalSheets: 0,
    pinnedSheets: 0,
    activeSheet: "",
    hasUnsavedChanges: false,
  });

  const refreshStats = async () => {
    try {
      await Excel.run(async (context) => {
        const workbook = context.workbook;
        workbook.load("name");
        
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        
        const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
        activeWorksheet.load("name");
        
        await context.sync();

        // Đếm pinned sheets từ storage
        const pinnedSheets = StorageService.getPinnedSheets(workbook.name);
        const pinnedCount = pinnedSheets.size;

        setStats({
          totalSheets: worksheets.items.length,
          pinnedSheets: pinnedCount,
          activeSheet: activeWorksheet.name,
          hasUnsavedChanges: false, // Excel Online doesn't expose this easily
          lastModified: new Date().toLocaleTimeString(),
        });
      });
    } catch (error) {
      console.error("Error getting workbook stats:", error);
    }
  };

  React.useEffect(() => {
    refreshStats();

    // Auto-refresh mỗi 10 giây
    const autoRefreshInterval = setInterval(() => {
      refreshStats();
    }, 10000);

    // Refresh khi window focus
    const handleWindowFocus = () => {
      refreshStats();
    };

    window.addEventListener('focus', handleWindowFocus);

    return () => {
      clearInterval(autoRefreshInterval);
      window.removeEventListener('focus', handleWindowFocus);
    };
  }, []);

  const handleRefresh = () => {
    refreshStats();
    onRefresh();
  };

  return (
    <Card className={styles.container}>
      <CardHeader
        header={
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", width: "100%" }}>
            <div style={{ display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS }}>
              <Info24Regular />
              <Text weight="semibold">Workbook Info</Text>
            </div>
            <Button
              size="small"
              icon={<ArrowClockwise24Regular />}
              onClick={handleRefresh}
              className={styles.refreshButton}
              title="Refresh info"
            />
          </div>
        }
      />
      
      <div className={styles.statsGrid}>
        <div className={styles.statItem}>
          <Text className={styles.statValue}>{stats.totalSheets}</Text>
          <Text className={styles.statLabel}>Total Sheets</Text>
        </div>
        
        <div className={styles.statItem}>
          <Text className={styles.statValue}>{stats.pinnedSheets}</Text>
          <Text className={styles.statLabel}>Pinned Sheets</Text>
        </div>
        
        <div className={styles.statItem}>
          <Text className={styles.statValue} style={{ fontSize: tokens.fontSizeBase300 }}>
            {stats.activeSheet || "None"}
          </Text>
          <Text className={styles.statLabel}>Active Sheet</Text>
        </div>
        
        <div className={styles.statItem}>
          <Text className={styles.statValue} style={{ fontSize: tokens.fontSizeBase300 }}>
            {stats.lastModified || "Unknown"}
          </Text>
          <Text className={styles.statLabel}>Last Refreshed</Text>
        </div>
      </div>
    </Card>
  );
};

export default WorkbookInfo;
