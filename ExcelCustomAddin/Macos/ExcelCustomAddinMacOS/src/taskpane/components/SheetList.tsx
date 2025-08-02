import * as React from "react";
import {
  makeStyles,
  Button,
  Text,
  tokens,
  mergeClasses,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
} from "@fluentui/react-components";
import {
  Pin24Regular,
  Pin24Filled,
  ArrowClockwise24Regular,
  Rename24Regular,
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100%", // Chiếm toàn bộ chiều cao của container cha
    gap: tokens.spacingVerticalXS,
  },
  header: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalXS}`,
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusSmall,
    marginBottom: tokens.spacingVerticalXS,
    flexShrink: 0, // Prevent header from shrinking
  },
  headerActions: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalXS,
  },
  listContainer: {
    flex: 1, // Chiếm toàn bộ không gian còn lại
    overflowY: "auto", // Enable scrolling khi cần
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusSmall,
    padding: tokens.spacingVerticalXXS,
    minHeight: 0, // Cho phép shrink
  },
  sheetItem: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`, // Tăng padding cho rõ ràng hơn
    margin: `${tokens.spacingVerticalXXS} ${tokens.spacingHorizontalXXS}`, // Thêm margin để tạo khoảng cách
    borderRadius: tokens.borderRadiusSmall, // Thêm border radius cho đẹp
    cursor: "pointer",
    minHeight: "25px", // Tăng height để phù hợp với padding
    transition: "background-color 0.2s ease", // Smooth transition
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  selectedItem: {
    // Styling handled in getItemTextStyle function
  },
  pinnedItem: {
    // Styling handled in getItemTextStyle function
  },
  sheetInfo: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalXXS, // Giảm gap
    flex: 1,
  },
  sheetName: {
    flex: 1,
    textAlign: "left",
    fontSize: tokens.fontSizeBase200, // Giảm font size
    lineHeight: "1.2",
  },
  pinButton: {
    minWidth: "auto",
    padding: tokens.spacingVerticalXXS,
    width: "24px", // Fixed width
    height: "24px",
  },
});

interface SheetInfo {
  name: string;
  tabColor?: string;
  isPinned: boolean;
}

interface SheetListProps {
  sheets: SheetInfo[];
  selectedSheet: string;
  onSheetSelect: (sheetName: string) => void;
  onTogglePin: (sheetName: string) => void;
  onRefresh: () => void;
  onRenameSheet: (sheetName: string) => void;
}

const SheetList: React.FC<SheetListProps> = ({
  sheets,
  selectedSheet,
  onSheetSelect,
  onTogglePin,
  onRefresh,
  onRenameSheet,
}) => {
  const styles = useStyles();
  const [contextMenuOpen, setContextMenuOpen] = React.useState<string | null>(null);

  // Close context menu when clicking outside
  React.useEffect(() => {
    const handleClickOutside = () => {
      setContextMenuOpen(null);
    };

    if (contextMenuOpen) {
      document.addEventListener('click', handleClickOutside);
      return () => document.removeEventListener('click', handleClickOutside);
    }
    
    return undefined;
  }, [contextMenuOpen]);

  // Sắp xếp sheets: pinned sheets lên đầu
  const sortedSheets = React.useMemo(() => {
    return [...sheets].sort((a, b) => {
      if (a.isPinned && !b.isPinned) return -1;
      if (!a.isPinned && b.isPinned) return 1;
      return 0;
    });
  }, [sheets]);

  const handleSheetClick = (sheetName: string) => {
    onSheetSelect(sheetName);
  };

  const handlePinClick = (e: React.MouseEvent, sheetName: string) => {
    e.stopPropagation(); // Prevent sheet selection when clicking pin
    onTogglePin(sheetName);
  };

  const handleContextMenu = (e: React.MouseEvent, sheetName: string) => {
    e.preventDefault(); // Prevent default browser context menu
    setContextMenuOpen(sheetName);
  };

  const handleRenameClick = (sheetName: string) => {
    onRenameSheet(sheetName);
    setContextMenuOpen(null); // Close menu after action
  };

  const getItemBackgroundStyle = () => {
    // No background colors, keep default background
    return {};
  };

  const getItemTextStyle = (sheet: SheetInfo) => {
    // Use text colors to indicate different states
    if (sheet.name === selectedSheet) {
      return { 
        color: tokens.colorBrandForeground1, // Brand color for selected
        fontWeight: tokens.fontWeightSemibold
      };
    }
    if (sheet.isPinned) {
      return { 
        color: tokens.colorNeutralForeground1, // Normal color for pinned
        fontWeight: tokens.fontWeightSemibold
      };
    }
    if (sheet.tabColor) {
      // Use the tab color for text color
      return {
        color: sheet.tabColor,
        fontWeight: tokens.fontWeightMedium
      };
    }
    // Default text color
    return {
      color: tokens.colorNeutralForeground1,
      fontWeight: tokens.fontWeightRegular
    };
  };

  return (
    <div className={styles.container}>
      {/* Compact Header */}
      <div className={styles.header}>
        <div className={styles.headerActions}>
          <Text weight="semibold" size={200}>
            Worksheets ({sheets.length})
          </Text>
          <Button
            size="small"
            appearance="subtle"
            icon={<ArrowClockwise24Regular />}
            onClick={onRefresh}
            title="Refresh sheet list"
          />
        </div>
      </div>
      
      <div className={styles.listContainer}>
        {sortedSheets.length === 0 ? (
          <div style={{ padding: tokens.spacingVerticalS, textAlign: "center" }}>
            <Text size={200}>No worksheets found</Text>
          </div>
        ) : (
          sortedSheets.map((sheet) => (
            <div key={sheet.name} style={{ position: 'relative' }}>
              <div
                className={mergeClasses(
                  styles.sheetItem,
                  sheet.name === selectedSheet && styles.selectedItem,
                  sheet.isPinned && styles.pinnedItem
                )}
                style={{
                  ...getItemBackgroundStyle(),
                  ...getItemTextStyle(sheet)
                }}
                onClick={() => handleSheetClick(sheet.name)}
                onContextMenu={(e) => handleContextMenu(e, sheet.name)}
                title={`Left click to activate '${sheet.name}' | Right click for options`}
              >
                <div className={styles.sheetInfo}>
                  <Text className={styles.sheetName}>
                    {sheet.name}
                    {sheet.isPinned && " 📌"}
                  </Text>
                </div>
                
                <Button
                  size="small"
                  appearance="subtle"
                  icon={sheet.isPinned ? <Pin24Filled /> : <Pin24Regular />}
                  className={styles.pinButton}
                  onClick={(e) => handlePinClick(e, sheet.name)}
                  title={sheet.isPinned ? "Unpin sheet" : "Pin sheet"}
                />
              </div>
              
              {contextMenuOpen === sheet.name && (
                <Menu 
                  open={true}
                  onOpenChange={(_, data) => {
                    if (!data.open) setContextMenuOpen(null);
                  }}
                  positioning="below-end"
                >
                  <MenuTrigger disableButtonEnhancement>
                    <div style={{ position: 'absolute', top: 0, left: 0, width: 0, height: 0 }} />
                  </MenuTrigger>
                  <MenuPopover>
                    <MenuList>
                      <MenuItem 
                        icon={<Rename24Regular />}
                        onClick={() => handleRenameClick(sheet.name)}
                      >
                        Rename Sheet
                      </MenuItem>
                    </MenuList>
                  </MenuPopover>
                </Menu>
              )}
            </div>
          ))
        )}
      </div>
    </div>
  );
};

export default SheetList;
