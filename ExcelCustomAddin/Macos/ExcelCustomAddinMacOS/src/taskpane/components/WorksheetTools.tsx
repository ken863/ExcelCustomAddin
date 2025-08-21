import * as React from "react";
import {
  makeStyles,
  Button,
  Card,
  CardHeader,
  Text,
  Input,
  Label,
  SpinButton,
  Divider,
  tokens,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogContent,
  DialogActions,
  DialogBody,
  MessageBar,
  Checkbox,
} from "@fluentui/react-components";
import {
  DocumentAdd24Regular,
  Image24Regular,
  DocumentBulletListMultiple24Regular,
  Copy24Regular,
} from "@fluentui/react-icons";
import SheetList from "./SheetList";
import StorageService from "../services/StorageService";
import ExcelUtils, { SheetInfo } from "../utils/ExcelUtils";
import WorksheetOperations from "../utils/WorksheetOperations";
import ImageUtils from "../utils/ImageUtils";
import EventHandlers from "../utils/EventHandlers";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100vh", // Chiếm toàn bộ chiều cao viewport
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalS,
    overflow: "hidden", // Prevent container from scrolling
  },
  toolSection: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },
  buttonRow: {
    display: "flex",
    gap: tokens.spacingHorizontalXS,
    flexWrap: "wrap",
  },
  pathDisplay: {
    backgroundColor: tokens.colorNeutralBackground2,
    padding: tokens.spacingVerticalXXS,
    borderRadius: tokens.borderRadiusSmall,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    wordBreak: "break-all",
  },
  inputGroup: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXXS,
  },
  numberInput: {
    width: "100px",
  },
  compactCard: {
    padding: tokens.spacingVerticalXS,
    flexShrink: 0, // Prevent cards from shrinking
  },
  compactCardHeader: {
    marginBottom: tokens.spacingVerticalXXS,
  },
  horizontalGroup: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
    flexWrap: "wrap",
  },
  scaleGroup: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalXS,
  },
  sheetListContainer: {
    flex: 1, // Chiếm toàn bộ không gian còn lại
    minHeight: 0, // Cho phép shrink xuống dưới content size
    display: "flex",
    flexDirection: "column",
  },
  fixedSection: {
    flexShrink: 0, // Prevent fixed sections from shrinking
  },
});

const WorksheetTools: React.FC = () => {
  const styles = useStyles();
  const [currentWorkbookPath, setCurrentWorkbookPath] = React.useState<string>("");
  const [scalePercent, setScalePercent] = React.useState<number>(85);
  const [allowReimport, setAllowReimport] = React.useState<boolean>(false);
  const [sheets, setSheets] = React.useState<SheetInfo[]>([]);
  const [selectedSheet, setSelectedSheet] = React.useState<string>("");
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [lastRefresh, setLastRefresh] = React.useState<Date>(new Date());
  const [message, setMessage] = React.useState<string>("");

  // Rename dialog state
  const [renameDialogOpen, setRenameDialogOpen] = React.useState<boolean>(false);
  const [newSheetName, setNewSheetName] = React.useState<string>("");

  // Auto-save settings khi thay đổi
  React.useEffect(() => {
    StorageService.saveScalePercent(scalePercent);
  }, [scalePercent]);

  React.useEffect(() => {
    StorageService.setAllowReimport(allowReimport);
  }, [allowReimport]);

  // Auto-hide message sau 3 giây
  React.useEffect(() => {
    if (message) {
      const timer = setTimeout(() => {
        setMessage("");
      }, 3000);

      return () => clearTimeout(timer);
    }

    return undefined;
  }, [message]);

  // Auto-refresh khi window được focus lại
  React.useEffect(() => {
    const handleWindowFocus = () => {
      console.log("Window focused, refreshing sheet list...");
      refreshSheetList();
    };

    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible') {
        console.log("Document visible, refreshing sheet list...");
        refreshSheetList();
      }
    };

    window.addEventListener('focus', handleWindowFocus);
    document.addEventListener('visibilitychange', handleVisibilityChange);

    return () => {
      window.removeEventListener('focus', handleWindowFocus);
      document.removeEventListener('visibilitychange', handleVisibilityChange);
    };
  }, []);

  // Khởi tạo và lấy danh sách sheets
  React.useEffect(() => {
    // Load saved settings from storage
    setScalePercent(StorageService.getScalePercent());
    setAllowReimport(StorageService.getAllowReimport());

    refreshWorkbookInfo();
    refreshSheetList();

    // Khởi tạo event handlers sử dụng EventHandlers utility
    const initializeEvents = async () => {
      try {
        // Đăng ký handlers
        const unregisterActivated = EventHandlers.registerSheetActivatedHandler(refreshSheetList);
        const unregisterAdded = EventHandlers.registerSheetAddedHandler(refreshSheetList);
        const unregisterDeleted = EventHandlers.registerSheetDeletedHandler(refreshSheetList);

        // Đăng ký workbook closed handler để clear storage
        const unregisterClosed = EventHandlers.registerWorkbookClosedHandler((workbookName) => {
          console.log(`Cleared pinned sheets storage for workbook: ${workbookName}`);
          setMessage(`Cleared pinned sheets for: ${workbookName}`);
        });

        // Khởi tạo Excel events
        await EventHandlers.initializeExcelEvents();

        // Khởi tạo workbook monitoring
        EventHandlers.initializeWorkbookMonitoring();

        // Auto-refresh fallback mỗi 5 giây
        EventHandlers.startPollingFallback(5000);

        // Return cleanup function
        return () => {
          unregisterActivated();
          unregisterAdded();
          unregisterDeleted();
          unregisterClosed();
          EventHandlers.cleanup();
        };
      } catch (error) {
        console.error("Error initializing events:", error);
        // Fallback to polling only
        EventHandlers.startPollingFallback(5000);

        return () => {
          EventHandlers.cleanup();
        };
      }
    };

    let cleanup: (() => void) | undefined;

    initializeEvents().then((cleanupFn) => {
      cleanup = cleanupFn;
    });

    // Component cleanup
    return () => {
      if (cleanup) {
        cleanup();
      }
      console.log("WorksheetTools component unmounting");
    };
  }, []);

  const refreshWorkbookInfo = async () => {
    try {
      const workbookInfo = await ExcelUtils.getWorkbookInfo();
      setCurrentWorkbookPath(workbookInfo.path);
    } catch (error) {
      console.error("Error getting workbook info:", error);
    }
  };

  const refreshSheetList = async () => {
    try {
      const sheetData = await ExcelUtils.getWorksheetList();

      // Cập nhật pin status cho mỗi sheet
      const sheetList: SheetInfo[] = sheetData.worksheets.map(sheet => ({
        ...sheet,
        isPinned: StorageService.isSheetPinned(sheetData.workbookName, sheet.name),
      }));

      // Sắp xếp: sheets được pin lên đầu
      const sortedSheets = sheetList.sort((a, b) => {
        if (a.isPinned && !b.isPinned) return -1;
        if (!a.isPinned && b.isPinned) return 1;
        return 0;
      });

      setSheets(sortedSheets);
      setSelectedSheet(sheetData.activeSheet);
      setLastRefresh(new Date());
    } catch (error) {
      console.error("Error getting sheet list:", error);
    }
  };

  // Tạo Evidence Sheet (tương đương CreateEvidence)
  const createEvidenceSheet = async () => {
    if (isLoading) return;
    setIsLoading(true);

    try {
      const result = await WorksheetOperations.createEvidenceSheet();

      if (result.success) {
        setMessage(result.message);
        refreshSheetList();
      } else {
        alert(result.message);
      }
    } catch (error) {
      console.error("Error creating evidence sheet:", error);
      alert(`Có lỗi xảy ra: ${error.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  // Format Document (tương đương FormatDocument)
  const formatDocument = async () => {
    if (isLoading) return;
    setIsLoading(true);

    try {
      const result = await WorksheetOperations.formatDocument();

      if (result.success) {
        alert(result.message);
      } else {
        alert(result.message);
      }
    } catch (error) {
      console.error("Error formatting document:", error);
      alert(`Có lỗi xảy ra: ${error.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  // Change Sheet Name với dialog thay vì prompt
  const changeSheetName = (sheetNameOrEvent?: string | React.MouseEvent) => {
    let targetSheet: string;

    if (typeof sheetNameOrEvent === 'string') {
      // Called from context menu with sheet name
      targetSheet = sheetNameOrEvent;
    } else {
      // Called from button click, use selectedSheet
      targetSheet = selectedSheet;
    }

    if (!targetSheet) {
      setMessage("Vui lòng chọn một sheet từ danh sách để đổi tên.");
      return;
    }

    setNewSheetName(targetSheet);
    setSelectedSheet(targetSheet); // Ensure the sheet is selected
    setRenameDialogOpen(true);
  };

  // Handle Enter key in rename dialog
  const handleRenameKeyDown = (event: React.KeyboardEvent) => {
    if (event.key === 'Enter' && newSheetName.trim() && !isLoading) {
      handleRenameConfirm();
    }
  };

  // Handle rename confirmation
  const handleRenameConfirm = async () => {
    if (!newSheetName || newSheetName.trim() === "" || newSheetName === selectedSheet) {
      setRenameDialogOpen(false);
      return;
    }

    try {
      if (isLoading) return;
      setIsLoading(true);

      const result = await WorksheetOperations.renameWorksheet(selectedSheet, newSheetName);

      if (result.success) {
        setMessage(result.message);
        refreshSheetList();
        // Update selected sheet to new name
        setSelectedSheet(newSheetName);
      } else {
        setMessage(result.message);
      }
    } catch (error) {
      console.error("Error changing sheet name:", error);
      setMessage(`Có lỗi xảy ra: ${error.message}`);
    } finally {
      setIsLoading(false);
      setRenameDialogOpen(false);
    }
  };

  // Insert Multiple Images - tự động import từ folder path hoặc file picker
  const insertMultipleImages = async () => {
    if (isLoading) return;
    setIsLoading(true);
    
    try {
      const result = await ImageUtils.importImagesFromFolder();
      
      // Sử dụng setMessage thay vì alert
      if (result.message) {
        setMessage(result.message);
      }
    } catch (error) {
      console.error("Error inserting images:", error);
      setMessage(`Có lỗi xảy ra: ${error.message}`);
    } finally {
      setIsLoading(false);
    }
  };  // Pin/Unpin Sheet - sử dụng StorageService
  const togglePinSheet = async (sheetName: string) => {
    try {
      await Excel.run(async (context) => {
        const workbook = context.workbook;
        workbook.load("name");
        await context.sync();

        const workbookName = workbook.name;
        const isPinned = StorageService.togglePinSheet(workbookName, sheetName);

        // Cập nhật state local
        setSheets(prevSheets =>
          prevSheets.map(sheet =>
            sheet.name === sheetName
              ? { ...sheet, isPinned }
              : sheet
          )
        );

        // Refresh để sắp xếp lại danh sách
        setTimeout(() => refreshSheetList(), 100);
      });
    } catch (error) {
      console.error("Error toggling pin sheet:", error);
    }
  };

  // Sheet selection handler
  const handleSheetSelect = async (sheetName: string) => {
    setSelectedSheet(sheetName);

    try {
      await ExcelUtils.activateWorksheet(sheetName);
    } catch (error) {
      console.error("Error activating sheet:", error);
    }
  };

  const copyFilePath = async () => {
    try {
      if (navigator.clipboard) {
        await navigator.clipboard.writeText(currentWorkbookPath);
        setMessage("File path copied to clipboard");
      } else {
        setMessage("Clipboard not available");
      }
    } catch (error) {
      console.error("Error copying to clipboard:", error);
      setMessage("Failed to copy file path");
    }
  };

  return (
    <div className={styles.container}>
      {/* Message Bar - Fixed */}
      {message && (
        <div className={styles.fixedSection}>
          <MessageBar>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', width: '100%' }}>
              <div style={{ flex: 1 }}>
                {message.split('\n').map((line, index) => (
                  <div key={index}>
                    <Text size={200}>{line}</Text>
                  </div>
                ))}
              </div>
              <Button
                appearance="transparent"
                size="small"
                onClick={() => setMessage("")}
                style={{ marginLeft: '8px' }}
              >
                ✕
              </Button>
            </div>
          </MessageBar>
        </div>
      )}

      {/* Copy File Path Button - Fixed */}
      <div className={styles.fixedSection}>
        <Button
          size="small"
          icon={<Copy24Regular />}
          onClick={copyFilePath}
          appearance="subtle"
          style={{
            width: '100%',
            border: `1px solid ${tokens.colorNeutralStroke1}`,
            borderRadius: tokens.borderRadiusSmall
          }}
        >
          <Text size={200}>Copy file path</Text>
        </Button>
      </div>

      <div className={styles.fixedSection}>
        <Divider />
      </div>

      {/* Main Tools - Fixed */}
      <Card className={`${styles.compactCard} ${styles.fixedSection}`}>
        <CardHeader
          className={styles.compactCardHeader}
          header={<Text weight="semibold" size={200}>Worksheet Tools</Text>}
        />
        <div className={styles.toolSection}>
          <div className={styles.buttonRow}>
            <Button
              size="small"
              icon={<DocumentAdd24Regular />}
              onClick={createEvidenceSheet}
              disabled={isLoading}
              style={{ flex: 1 }}
            >
              {isLoading ? "Creating..." : "Create Evidence"}
            </Button>
            <Button
              size="small"
              icon={<DocumentBulletListMultiple24Regular />}
              onClick={formatDocument}
              style={{ flex: 1 }}
            >
              Format Document
            </Button>
          </div>
        </div>
      </Card>

      {/* Image Tools - Fixed */}
      <Card className={`${styles.compactCard} ${styles.fixedSection}`}>
        <CardHeader
          className={styles.compactCardHeader}
          header={<Text weight="semibold" size={200}>Image Tools</Text>}
        />
        <div className={styles.toolSection}>
          <div className={styles.horizontalGroup}>
            <div className={styles.scaleGroup}>
              <Label htmlFor="scalePercent" size="small">
                <Text size={200}>Scale:</Text>
              </Label>
              <SpinButton
                size="small"
                id="scalePercent"
                value={scalePercent}
                onChange={(_, data) => setScalePercent(data.value || 85)}
                min={10}
                max={200}
                step={5}
                style={{ width: '60px' }}
              />
            </div>

            <Button
              size="small"
              icon={<Image24Regular />}
              onClick={insertMultipleImages}
              appearance="primary"
              disabled={isLoading}
            >
              {isLoading ? "Đang import..." : "Import Images"}
            </Button>
          </div>

          <div className={styles.horizontalGroup}>
            <Checkbox
              checked={allowReimport}
              onChange={(_, data) => setAllowReimport(data.checked === true)}
              label={<Text size={200}>Allow reimport</Text>}
            />
          </div>
        </div>
      </Card>

      {/* Sheet List - Flexible */}
      <div className={styles.sheetListContainer}>
        <SheetList
          sheets={sheets}
          selectedSheet={selectedSheet}
          onSheetSelect={handleSheetSelect}
          onTogglePin={togglePinSheet}
          onRefresh={refreshSheetList}
          onRenameSheet={changeSheetName}
        />
      </div>

      {/* Rename Sheet Dialog */}
      <Dialog open={renameDialogOpen} onOpenChange={(_, data) => setRenameDialogOpen(data.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Đổi tên Sheet</DialogTitle>
            <DialogContent>
              <div style={{ marginBottom: "16px" }}>
                <Label>Tên sheet hiện tại: <strong>{selectedSheet}</strong></Label>
              </div>
              <div>
                <Label htmlFor="newSheetNameInput">Tên mới:</Label>
                <Input
                  id="newSheetNameInput"
                  value={newSheetName}
                  onChange={(_, data) => setNewSheetName(data.value)}
                  onKeyDown={handleRenameKeyDown}
                  placeholder="Nhập tên mới cho sheet"
                  autoComplete="off"
                  autoFocus={renameDialogOpen}
                />
              </div>
            </DialogContent>
            <DialogActions>
              <Button
                appearance="secondary"
                onClick={() => setRenameDialogOpen(false)}
                disabled={isLoading}
              >
                Hủy
              </Button>
              <Button
                appearance="primary"
                onClick={handleRenameConfirm}
                disabled={isLoading || !newSheetName.trim()}
              >
                {isLoading ? "Đang đổi tên..." : "Đồng ý"}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
};

export default WorksheetTools;
