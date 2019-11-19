import * as React from 'react'
import { VirtualizedComboBox, IComboBoxOption, Icon, mergeStyles, IComboBox } from 'office-ui-fabric-react';

export interface IIconPickerProps {
    value: string;
    onChange: (event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string) => void;
    color?: string;
    label?: string;
}

export const IconPicker: React.FunctionComponent<IIconPickerProps> = props => {

    const color = props.color || "#0078d4";
    const label = props.label || "Type";

    const allIcons: IComboBoxOption[] = getIcons().map((item) => ({
        key: item,
        text: item.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/\b([A-Z]+)([A-Z])([a-z])/, '$1 $2$3').replace(/^./, function(str){ return str.toUpperCase(); }),
        data: { icon: item }
    }));

    const iconClass = mergeStyles({ marginRight: '8px', color: color });

    const renderOption = (option: IComboBoxOption) => {
        return (
            <div>
                {option.data && option.data.icon && (
                    <Icon className={iconClass} iconName={option.data.icon} aria-hidden="true" title={option.data.icon} />
                )}
                <span>{option.text}</span>
            </div>
        );
    };

    return <VirtualizedComboBox label={label} selectedKey={props.value} onChange={props.onChange} options={allIcons} autoComplete="on" allowFreeform={false} onRenderOption={renderOption} />;

}

function getIcons() {

    //https://stackoverflow.com/questions/43100718/typescript-enum-to-object-array
    //React.IconNames

    return ["PageLink", "CommentSolid", "ChangeEntitlements", "Installation", "WebAppBuilderModule", "WebAppBuilderFragment", "WebAppBuilderSlot", "BullseyeTargetEdit",
        "WebAppBuilderFragmentCreate", "PageData", "PageHeaderEdit", "ProductList", "UnpublishContent", "DependencyAdd", "DependencyRemove", "EntitlementPolicy", "EntitlementRedemption",
        "GlobalNavButton", "InternetSharing", "Brightness", "MapPin", "Airplane", "Tablet", "QuickNote", "ChevronDown", "ChevronUp", "Edit", "Add", "Cancel", "More", "Settings", "Video",
        "Mail", "People", "Phone", "Pin", "Shop", "Stop", "Link", "Filter", "AllApps", "Zoom", "ZoomOut", "Microphone", "Search", "Camera", "Attach", "Send", "FavoriteList", "PageSolid",
        "Forward", "Back", "Refresh", "Share", "Lock", "BlockedSite", "ReportHacked", "EMI", "MiniLink", "Blocked", "FavoriteStar", "FavoriteStarFill", "ReadingMode", "Favicon", "Remove",
        "Checkbox", "CheckboxComposite", "CheckboxFill", "CheckboxIndeterminate", "CheckboxCompositeReversed", "CheckMark", "BackToWindow", "FullScreen", "Print", "Up", "Down", "OEM",
        "Delete", "Save", "ReturnKey", "Cloud", "Flashlight", "CommandPrompt", "Sad", "RealEstate", "SIPMove", "EraseTool", "GripperTool", "Dialpad", "PageLeft", "PageRight", "MultiSelect",
        "KeyboardClassic", "Play", "Pause", "ChevronLeft", "ChevronRight", "InkingTool", "Emoji2", "GripperBarHorizontal", "System", "Personalize", "SearchAndApps", "Globe", "EaseOfAccess",
        "ContactInfo", "Unpin", "Contact", "Memo", "IncomingCall", "Paste", "WindowsLogo", "Error", "GripperBarVertical", "Unlock", "Slideshow", "Calendar", "Megaphone", "AutoEnhanceOn",
        "AutoEnhanceOff", "Color", "SaveAs", "Light", "Filters", "AspectRatio", "Contrast", "Redo", "Undo", "Crop", "PhotoCollection", "Album", "Rotate", "PanoIndicator", "Translate",
        "RedEye", "ViewOriginal", "ThumbnailView", "Package", "Telemarketer", "Warning", "Financial", "Education", "ShoppingCart", "Train", "Flag", "Move", "Page", "TouchPointer", "Merge",
        "TurnRight", "Ferry", "Highlight", "PowerButton", "Tab", "Admin", "TVMonitor", "Speakers", "Game", "UnstackSelected", "StackIndicator", "Nav2DMapView", "StreetsideSplitMinimize",
        "Car", "Bus", "EatDrink", "SeeDo", "LocationCircle", "Home", "SwitcherStartEnd", "ParkingLocation", "IncidentTriangle", "Touch", "MapDirections", "CaretHollow", "CaretSolid",
        "History", "Location", "MapLayers", "SearchNearby", "Work", "Recent", "Hotel", "Bank", "LocationDot", "Dictionary", "ChromeBack", "FolderOpen", "Pinned", "PinnedFill",
        "RevToggleKey", "USB", "View", "Previous", "Next", "Clear", "Sync", "Download", "Help", "Upload", "Emoji", "MailForward", "ClosePane", "OpenPane", "PreviewLink", "ZoomIn",
        "Bookmarks", "Document", "ProtectedDocument", "OpenInNewWindow", "MailFill", "ViewAll", "Switch", "Rename", "Go", "Remote", "SelectAll", "Orientation", "Import", "Folder", "Picture",
        "ChromeClose", "ShowResults", "Message", "CalendarDay", "CalendarWeek", "MailReplyAll", "Read", "Cut", "PaymentCard", "Copy", "Important", "MailReply", "Sort", "GotoToday", "Font",
        "FontColor", "FolderFill", "Permissions", "DisableUpdates", "Unfavorite", "Italic", "Underline", "Bold", "MoveToFolder", "Dislike", "Like", "AlignRight", "AlignCenter", "AlignLeft",
        "OpenFile", "ClearSelection", "FontDecrease", "FontIncrease", "FontSize", "CellPhone", "Tag", "RepeatOne", "RepeatAll", "Calculator", "Library", "PostUpdate", "NewFolder",
        "CalendarReply", "UnsyncFolder", "SyncFolder", "BlockContact", "AddFriend", "Accept", "BulletedList", "Preview", "News", "Chat", "Group", "World", "Comment", "DockLeft", "DockRight",
        "Repair", "Accounts", "Street", "RadioBullet", "Stopwatch", "Clock", "WorldClock", "AlarmClock", "Photo", "ActionCenter", "Hospital", "Timer", "FullCircleMask", "LocationFill",
        "ChromeMinimize", "ChromeRestore", "Annotation", "Fingerprint", "Handwriting", "ChromeFullScreen", "Completed", "Label", "FlickDown", "FlickUp", "FlickLeft", "FlickRight",
        "MiniExpand", "MiniContract", "Streaming", "MusicInCollection", "OneDriveLogo", "CompassNW", "Code", "LightningBolt", "Info", "CalculatorMultiply", "CalculatorAddition",
        "CalculatorSubtract", "CalculatorPercentage", "CalculatorEqualTo", "PrintfaxPrinterFile", "StorageOptical", "Communications", "Headset", "Health", "FrontCamera", "ChevronUpSmall",
        "ChevronDownSmall", "ChevronLeftSmall", "ChevronRightSmall", "ChevronUpMed", "ChevronDownMed", "ChevronLeftMed", "ChevronRightMed", "Devices2", "PC1", "PresenceChickletVideo",
        "Reply", "HalfAlpha", "ConstructionCone", "DoubleChevronLeftMed", "Volume0", "Volume1", "Volume2", "Volume3", "Chart", "Robot", "Manufacturing", "LockSolid", "FitPage", "FitWidth",
        "BidiLtr", "BidiRtl", "RightDoubleQuote", "Sunny", "CloudWeather", "Cloudy", "PartlyCloudyDay", "PartlyCloudyNight", "ClearNight", "RainShowersDay", "Rain", "Thunderstorms",
        "RainSnow", "Snow", "BlowingSnow", "Frigid", "Fog", "Squalls", "Duststorm", "Unknown", "Precipitation", "SortLines", "Ribbon", "AreaChart", "Assign", "FlowChart", "CheckList",
        "Diagnostic", "Generate", "LineChart", "Equalizer", "BarChartHorizontal", "BarChartVertical", "Freezing", "FunnelChart", "Processing", "Quantity", "ReportDocument", "StackColumnChart",
        "SnowShowerDay", "HailDay", "WorkFlow", "HourGlass", "StoreLogoMed20", "TimeSheet", "TriangleSolid", "UpgradeAnalysis", "VideoSolid", "RainShowersNight", "SnowShowerNight", "Teamwork",
        "HailNight", "PeopleAdd", "Glasses", "DateTime2", "Shield", "Header1", "PageAdd", "NumberedList", "PowerBILogo", "Info2", "MusicInCollectionFill", "List", "Asterisk", "ErrorBadge",
        "CircleRing", "CircleFill", "Record2", "AllAppsMirrored", "BookmarksMirrored", "BulletedListMirrored", "CaretHollowMirrored", "CaretSolidMirrored", "ChromeBackMirrored",
        "ClearSelectionMirrored", "ClosePaneMirrored", "DockLeftMirrored", "DoubleChevronLeftMedMirrored", "GoMirrored", "HelpMirrored", "ImportMirrored", "ImportAllMirrored",
        "ListMirrored", "MailForwardMirrored", "MailReplyMirrored", "MailReplyAllMirrored", "MiniContractMirrored", "MiniExpandMirrored", "OpenPaneMirrored", "ParkingLocationMirrored",
        "SendMirrored", "ShowResultsMirrored", "ThumbnailViewMirrored", "Media", "Devices3", "Focus", "VideoLightOff", "Lightbulb", "StatusTriangle", "VolumeDisabled", "Puzzle",
        "EmojiNeutral", "EmojiDisappointed", "HomeSolid", "Ringer", "PDF", "HeartBroken", "StoreLogo16", "MultiSelectMirrored", "Broom", "AddToShoppingList", "Cocktails", "Wines",
        "Articles", "Cycling", "DietPlanNotebook", "Pill", "ExerciseTracker", "HandsFree", "Medical", "Running", "Weights", "Trackers", "AddNotes", "AllCurrency", "BarChart4", "CirclePlus",
        "Coffee", "Cotton", "Market", "Money", "PieDouble", "PieSingle", "RemoveFilter", "Savings", "Sell", "StockDown", "StockUp", "Lamp", "Source", "MSNVideos", "Cricket", "Golf",
        "Baseball", "Soccer", "MoreSports", "AutoRacing", "CollegeHoops", "CollegeFootball", "ProFootball", "ProHockey", "Rugby", "SubstitutionsIn", "Tennis", "Arrivals", "Design",
        "Website", "Drop", "HistoricalWeather", "SkiResorts", "Snowflake", "BusSolid", "FerrySolid", "AirplaneSolid", "TrainSolid", "Heart", "HeartFill", "Ticket", "WifiWarning4",
        "Devices4", "AzureLogo", "BingLogo", "MSNLogo", "OutlookLogoInverse", "OfficeLogo", "SkypeLogo", "Door", "EditMirrored", "GiftCard", "DoubleBookmark", "StatusErrorFull",
        "Certificate", "FastForward", "Rewind", "Photo2", "OpenSource", "Movers", "CloudDownload", "Family", "WindDirection", "Bug", "SiteScan", "BrowserScreenShot", "F12DevTools", "CSS",
        "JS", "DeliveryTruck", "ReminderPerson", "ReminderGroup", "TabletMode", "Umbrella", "NetworkTower", "CityNext", "CityNext2", "Section", "OneNoteLogoInverse", "ToggleFilled",
        "ToggleBorder", "SliderThumb", "ToggleThumb", "Documentation", "Badge", "Giftbox", "VisualStudioLogo", "HomeGroup", "ExcelLogoInverse", "WordLogoInverse", "PowerPointLogoInverse",
        "Cafe", "SpeedHigh", "Commitments", "ThisPC", "MusicNote", "MicOff", "PlaybackRate1x", "EdgeLogo", "CompletedSolid", "AlbumRemove", "MessageFill", "TabletSelected",
        "MobileSelected", "LaptopSelected", "TVMonitorSelected", "DeveloperTools", "Shapes", "InsertTextBox", "LowerBrightness", "WebComponents", "OfflineStorage", "DOM", "CloudUpload",
        "ScrollUpDown", "DateTime", "Event", "Cake", "Tiles", "Org", "PartyLeader", "DRM", "CloudAdd", "AppIconDefault", "Photo2Add", "Photo2Remove", "Calories", "POI", "AddTo",
        "RadioBtnOff", "RadioBtnOn", "ExploreContent", "Embed", "Product", "ProgressLoopInner", "ProgressLoopOuter", "Blocked2", "FangBody", "Toolbox", "PageHeader", "Glimmer",
        "ChatInviteFriend", "Brush", "Shirt", "Crown", "Diamond", "ScaleUp", "QRCode", "Feedback", "SharepointLogoInverse", "YammerLogo", "Hide", "Uneditable", "ReturnToSession",
        "OpenFolderHorizontal", "CalendarMirrored", "SwayLogoInverse", "OutOfOffice", "Trophy", "ReopenPages", "EmojiTabSymbols", "AADLogo", "AccessLogo", "AdminALogoInverse32",
        "AdminCLogoInverse32", "AdminDLogoInverse32", "AdminELogoInverse32", "AdminLLogoInverse32", "AdminMLogoInverse32", "AdminOLogoInverse32", "AdminPLogoInverse32",
        "AdminSLogoInverse32", "AdminYLogoInverse32", "DelveLogoInverse", "ExchangeLogoInverse", "LyncLogo", "OfficeVideoLogoInverse", "SocialListeningLogo", "VisioLogoInverse", "Balloons",
        "Cat", "MailAlert", "MailCheck", "MailLowImportance", "MailPause", "MailRepeat", "SecurityGroup", "Table", "VoicemailForward", "VoicemailReply", "Waffle", "RemoveEvent", "EventInfo",
        "ForwardEvent", "WipePhone", "AddOnlineMeeting", "JoinOnlineMeeting", "RemoveLink", "PeopleBlock", "PeopleRepeat", "PeopleAlert", "PeoplePause", "TransferCall", "AddPhone",
        "UnknownCall", "NoteReply", "NoteForward", "NotePinned", "RemoveOccurrence", "Timeline", "EditNote", "CircleHalfFull", "Room", "Unsubscribe", "Subscribe", "HardDrive", "RecurringTask",
        "TaskManager", "TaskManagerMirrored", "Combine", "Split", "DoubleChevronUp", "DoubleChevronLeft", "DoubleChevronRight", "Ascending", "Descending", "TextBox", "TextField", "NumberField",
        "Dropdown", "PenWorkspace", "BookingsLogo", "ClassNotebookLogoInverse", "DelveAnalyticsLogo", "DocsLogoInverse", "Dynamics365Logo", "DynamicSMBLogo", "OfficeAssistantLogo",
        "OfficeStoreLogo", "OneNoteEduLogoInverse", "PlannerLogo", "PowerApps", "Suitcase", "ProjectLogoInverse", "CaretLeft8", "CaretRight8", "CaretUp8", "CaretDown8", "CaretLeftSolid8",
        "CaretRightSolid8", "CaretUpSolid8", "CaretDownSolid8", "ClearFormatting", "Superscript", "Subscript", "Strikethrough", "Export", "ExportMirrored", "SingleBookmark",
        "SingleBookmarkSolid", "DoubleChevronDown", "FollowUser", "ReplyAll", "WorkforceManagement", "RecruitmentManagement", "Questionnaire", "ManagerSelfService", "ProductionFloorManagement",
        "ProductRelease", "ProductVariant", "ReplyMirrored", "ReplyAllMirrored", "Medal", "AddGroup", "QuestionnaireMirrored", "CloudImportExport", "TemporaryUser", "CaretSolid16",
        "GroupedDescending", "GroupedAscending", "SortUp", "SortDown", "AwayStatus", "MyMoviesTV", "SyncToPC", "GenericScan", "AustralianRules", "WifiEthernet", "TrackersMirrored",
        "DateTimeMirrored", "StopSolid", "DoubleChevronUp12", "DoubleChevronDown12", "DoubleChevronLeft12", "DoubleChevronRight12", "CalendarAgenda", "ConnectVirtualMachine", "AddEvent",
        "AssetLibrary", "DataConnectionLibrary", "DocLibrary", "FormLibrary", "FormLibraryMirrored", "ReportLibrary", "ReportLibraryMirrored", "ContactCard", "CustomList", "CustomListMirrored",
        "IssueTracking", "IssueTrackingMirrored", "PictureLibrary", "OfficeAddinsLogo", "OfflineOneDriveParachute", "OfflineOneDriveParachuteDisabled", "LargeGrid", "TriangleSolidUp12",
        "TriangleSolidDown12", "TriangleSolidLeft12", "TriangleSolidRight12", "TriangleUp12", "TriangleDown12", "TriangleLeft12", "TriangleRight12", "ArrowUpRight8", "ArrowDownRight8",
        "DocumentSet", "GoToDashboard", "DelveAnalytics", "ArrowUpRightMirrored8", "ArrowDownRightMirrored8", "CompanyDirectory", "OpenEnrollment", "CompanyDirectoryMirrored", "OneDriveAdd",
        "ProfileSearch", "Header2", "Header3", "Header4", "RingerSolid", "Eyedropper", "MarketDown", "CalendarWorkWeek", "SidePanel", "GlobeFavorite", "CaretTopLeftSolid8",
        "CaretTopRightSolid8", "ViewAll2", "DocumentReply", "PlayerSettings", "ReceiptForward", "ReceiptReply", "ReceiptCheck", "Fax", "RecurringEvent", "ReplyAlt", "ReplyAllAlt", "EditStyle",
        "EditMail", "Lifesaver", "LifesaverLock", "InboxCheck", "FolderSearch", "CollapseMenu", "ExpandMenu", "Boards", "SunAdd", "SunQuestionMark", "LandscapeOrientation", "DocumentSearch",
        "PublicCalendar", "PublicContactCard", "PublicEmail", "PublicFolder", "WordDocument", "PowerPointDocument", "ExcelDocument", "GroupedList", "ClassroomLogo", "Sections", "EditPhoto",
        "Starburst", "ShareiOS", "AirTickets", "PencilReply", "Tiles2", "SkypeCircleCheck", "SkypeCircleClock", "SkypeCircleMinus", "SkypeCheck", "SkypeClock", "SkypeMinus", "SkypeMessage",
        "ClosedCaption", "ATPLogo", "OfficeFormsLogoInverse", "RecycleBin", "EmptyRecycleBin", "Hide2", "Breadcrumb", "BirthdayCake", "ClearFilter", "Flow", "TimeEntry", "CRMProcesses",
        "PageEdit", "PageArrowRight", "PageRemove", "Database", "DataManagementSettings", "CRMServices", "EditContact", "ConnectContacts", "AppIconDefaultAdd", "AppIconDefaultList",
        "ActivateOrders", "DeactivateOrders", "ProductCatalog", "ScatterChart", "AccountActivity", "DocumentManagement", "CRMReport", "KnowledgeArticle", "Relationship", "HomeVerify",
        "ZipFolder", "SurveyQuestions", "TextDocument", "TextDocumentShared", "PageCheckedOut", "SaveAndClose", "Script", "Archive", "ActivityFeed", "Compare", "EventDate", "ArrowUpRight",
        "CaretRight", "SetAction", "ChatBot", "CaretSolidLeft", "CaretSolidDown", "CaretSolidRight", "CaretSolidUp", "PowerAppsLogo", "PowerApps2Logo", "SearchIssue", "SearchIssueMirrored",
        "FabricAssetLibrary", "FabricDataConnectionLibrary", "FabricDocLibrary", "FabricFormLibrary", "FabricFormLibraryMirrored", "FabricReportLibrary", "FabricReportLibraryMirrored",
        "FabricPublicFolder", "FabricFolderSearch", "FabricMovetoFolder", "FabricUnsyncFolder", "FabricSyncFolder", "FabricOpenFolderHorizontal", "FabricFolder", "FabricFolderFill",
        "FabricNewFolder", "FabricPictureLibrary", "PhotoVideoMedia", "AddFavorite", "AddFavoriteFill", "BufferTimeBefore", "BufferTimeAfter", "BufferTimeBoth", "PublishContent",
        "ClipboardList", "ClipboardListMirrored", "CannedChat", "SkypeForBusinessLogo", "TabCenter", "PageCheckedin", "PageList", "ReadOutLoud", "CaretBottomLeftSolid8",
        "CaretBottomRightSolid8", "FolderHorizontal", "MicrosoftStaffhubLogo", "GiftboxOpen", "StatusCircleOuter", "StatusCircleInner", "StatusCircleRing", "StatusTriangleOuter",
        "StatusTriangleInner", "StatusTriangleExclamation", "StatusCircleExclamation", "StatusCircleErrorX", "StatusCircleCheckmark", "StatusCircleInfo", "StatusCircleBlock",
        "StatusCircleBlock2", "StatusCircleQuestionMark", "StatusCircleSync", "Toll", "ExploreContentSingle", "CollapseContent", "CollapseContentSingle", "InfoSolid", "GroupList",
        "ProgressRingDots", "CaloriesAdd", "BranchFork", "AddHome", "AddWork", "MobileReport", "ScaleVolume", "HardDriveGroup", "FastMode", "ToggleLeft", "ToggleRight", "TriangleShape",
        "RectangleShape", "CubeShape", "Trophy2", "BucketColor", "BucketColorFill", "Taskboard", "SingleColumn", "DoubleColumn", "TripleColumn", "ColumnLeftTwoThirds", "ColumnRightTwoThirds",
        "AccessLogoFill", "AnalyticsLogo", "AnalyticsQuery", "NewAnalyticsQuery", "AnalyticsReport", "WordLogo", "WordLogoFill", "ExcelLogo", "ExcelLogoFill", "OneNoteLogo", "OneNoteLogoFill",
        "OutlookLogo", "OutlookLogoFill", "PowerPointLogo", "PowerPointLogoFill", "PublisherLogo", "PublisherLogoFill", "ScheduleEventAction", "FlameSolid", "ServerProcesses", "Server",
        "SaveAll", "LinkedInLogo", "Decimals", "SidePanelMirrored", "ProtectRestrict", "Blog", "UnknownMirrored", "PublicContactCardMirrored", "GridViewSmall", "GridViewMedium",
        "GridViewLarge", "Step", "StepInsert", "StepShared", "StepSharedAdd", "StepSharedInsert", "ViewDashboard", "ViewList", "ViewListGroup", "ViewListTree", "TriggerAuto", "TriggerUser",
        "PivotChart", "StackedBarChart", "StackedLineChart", "BuildQueue", "BuildQueueNew", "UserFollowed", "ContactLink", "Stack", "Bullseye", "VennDiagram", "FiveTileGrid", "FocalPoint",
        "RingerRemove", "TeamsLogoInverse", "TeamsLogo", "TeamsLogoFill", "SkypeForBusinessLogoFill", "SharepointLogo", "SharepointLogoFill", "DelveLogo", "DelveLogoFill", "OfficeVideoLogo",
        "OfficeVideoLogoFill", "ExchangeLogo", "ExchangeLogoFill", "Signin", "DocumentApproval", "CloneToDesktop", "InstallToDrive", "Blur", "Build", "ProcessMetaTask", "BranchFork2",
        "BranchLocked", "BranchCommit", "BranchCompare", "BranchMerge", "BranchPullRequest", "BranchSearch", "BranchShelveset", "RawSource", "MergeDuplicate", "RowsGroup", "RowsChild",
        "Deploy", "Redeploy", "ServerEnviroment", "VisioDiagram", "HighlightMappedShapes", "TextCallout", "IconSetsFlag", "VisioLogo", "VisioLogoFill", "VisioDocument", "TimelineProgress",
        "TimelineDelivery", "Backlog", "TeamFavorite", "TaskGroup", "TaskGroupMirrored", "ScopeTemplate", "AssessmentGroupTemplate", "NewTeamProject", "CommentAdd", "CommentNext",
        "CommentPrevious", "ShopServer", "LocaleLanguage", "QueryList", "UserSync", "UserPause", "StreamingOff", "MoreVertical", "ArrowTallUpLeft", "ArrowTallUpRight", "ArrowTallDownLeft",
        "ArrowTallDownRight", "FieldEmpty", "FieldFilled", "FieldChanged", "FieldNotChanged", "RingerOff", "PlayResume", "BulletedList2", "BulletedList2Mirrored", "ImageCrosshair", "GitGraph",
        "Repo", "RepoSolid", "FolderQuery", "FolderList", "FolderListMirrored", "LocationOutline", "POISolid", "CalculatorNotEqualTo", "BoxSubtractSolid", "BoxAdditionSolid",
        "BoxMultiplySolid", "BoxPlaySolid", "BoxCheckmarkSolid", "CirclePauseSolid", "CirclePause", "MSNVideosSolid", "CircleStopSolid", "CircleStop", "NavigateBack", "NavigateBackMirrored",
        "NavigateForward", "NavigateForwardMirrored", "UnknownSolid", "UnknownMirroredSolid", "CircleAddition", "CircleAdditionSolid", "FilePDB", "FileTemplate", "FileSQL", "FileJAVA",
        "FileASPX", "FileCSS", "FileSass", "FileLess", "FileHTML", "JavaScriptLanguage", "CSharpLanguage", "CSharp", "VisualBasicLanguage", "VB", "CPlusPlusLanguage", "CPlusPlus",
        "FSharpLanguage", "FSharp", "TypeScriptLanguage", "PythonLanguage", "PY", "CoffeeScript", "MarkDownLanguage", "FullWidth", "FullWidthEdit", "Plug", "PlugSolid", "PlugConnected",
        "PlugDisconnected", "UnlockSolid", "Variable", "Parameter", "CommentUrgent", "Storyboard", "DiffInline", "DiffSideBySide", "ImageDiff", "ImagePixel", "FileBug", "FileCode",
        "FileComment", "BusinessHoursSign", "FileImage", "FileSymlink", "AutoFillTemplate", "WorkItem", "WorkItemBug", "LogRemove", "ColumnOptions", "Packages", "BuildIssue", "AssessmentGroup",
        "VariableGroup", "FullHistory", "Wheelchair", "SingleColumnEdit", "DoubleColumnEdit", "TripleColumnEdit", "ColumnLeftTwoThirdsEdit", "ColumnRightTwoThirdsEdit", "StreamLogo",
        "PassiveAuthentication", "AlertSolid", "MegaphoneSolid", "TaskSolid", "ConfigurationSolid", "BugSolid", "CrownSolid", "Trophy2Solid", "QuickNoteSolid", "ConstructionConeSolid",
        "PageListSolid", "PageListMirroredSolid", "StarburstSolid", "ReadingModeSolid", "SadSolid", "HealthSolid", "ShieldSolid", "GiftBoxSolid", "ShoppingCartSolid", "MailSolid", "ChatSolid",
        "RibbonSolid", "FinancialSolid", "FinancialMirroredSolid", "HeadsetSolid", "PermissionsSolid", "ParkingSolid", "ParkingMirroredSolid", "DiamondSolid", "AsteriskSolid",
        "OfflineStorageSolid", "BankSolid", "DecisionSolid", "Parachute", "ParachuteSolid", "FiltersSolid", "ColorSolid", "ReviewSolid", "ReviewRequestSolid", "ReviewRequestMirroredSolid",
        "ReviewResponseSolid", "FeedbackRequestSolid", "FeedbackRequestMirroredSolid", "FeedbackResponseSolid", "WorkItemBar", "WorkItemBarSolid", "Separator", "NavigateExternalInline",
        "PlanView", "TimelineMatrixView", "EngineeringGroup", "ProjectCollection", "CaretBottomRightCenter8", "CaretBottomLeftCenter8", "CaretTopRightCenter8", "CaretTopLeftCenter8",
        "DonutChart", "ChevronUnfold10", "ChevronFold10", "DoubleChevronDown8", "DoubleChevronUp8", "DoubleChevronLeft8", "DoubleChevronRight8", "ChevronDownEnd6", "ChevronUpEnd6",
        "ChevronLeftEnd6", "ChevronRightEnd6", "ContextMenu", "AzureAPIManagement", "AzureServiceEndpoint", "VSTSLogo", "VSTSAltLogo1", "VSTSAltLogo2", "FileTypeSolution",
        "WordLogoInverse16", "WordLogo16", "WordLogoFill16", "PowerPointLogoInverse16", "PowerPointLogo16", "PowerPointLogoFill16", "ExcelLogoInverse16", "ExcelLogo16", "ExcelLogoFill16",
        "OneNoteLogoInverse16", "OneNoteLogo16", "OneNoteLogoFill16", "OutlookLogoInverse16", "OutlookLogo16", "OutlookLogoFill16", "PublisherLogoInverse16", "PublisherLogo16",
        "PublisherLogoFill16", "VisioLogoInverse16", "VisioLogo16", "VisioLogoFill16", "TestBeaker", "TestBeakerSolid", "TestExploreSolid", "TestAutoSolid", "TestUserSolid",
        "TestImpactSolid", "TestPlan", "TestStep", "TestParameter", "TestSuite", "TestCase", "Sprint", "SignOut", "TriggerApproval", "Rocket", "AzureKeyVault", "Onboarding", "Transition",
        "LikeSolid", "DislikeSolid", "CRMCustomerInsightsApp", "EditCreate", "PlayReverseResume", "PlayReverse", "SearchData", "UnSetColor", "DeclineCall", "RectangularClipping",
        "TeamsLogo16", "TeamsLogoFill16", "Spacer", "SkypeLogo16", "SkypeForBusinessLogo16", "SkypeForBusinessLogoFill16", "FilterSolid", "MailUndelivered", "MailTentative",
        "MailTentativeMirrored", "MailReminder", "ReceiptUndelivered", "ReceiptTentative", "ReceiptTentativeMirrored", "Inbox", "IRMReply", "IRMReplyMirrored", "IRMForward",
        "IRMForwardMirrored", "VoicemailIRM", "EventAccepted", "EventTentative", "EventTentativeMirrored", "EventDeclined", "IDBadge", "BackgroundColor", "OfficeFormsLogoInverse16",
        "OfficeFormsLogo", "OfficeFormsLogoFill", "OfficeFormsLogo16", "OfficeFormsLogoFill16", "OfficeFormsLogoInverse24", "OfficeFormsLogo24", "OfficeFormsLogoFill24", "PageLock",
        "NotExecuted", "NotImpactedSolid", "FieldReadOnly", "FieldRequired", "BacklogBoard", "ExternalBuild", "ExternalTFVC", "ExternalXAML", "IssueSolid", "DefectSolid", "LadybugSolid",
        "NugetLogo", "TFVCLogo", "ProjectLogo32", "ProjectLogoFill32", "ProjectLogo16", "ProjectLogoFill16", "SwayLogo32", "SwayLogoFill32", "SwayLogo16", "SwayLogoFill16",
        "ClassNotebookLogo32", "ClassNotebookLogoFill32", "ClassNotebookLogo16", "ClassNotebookLogoFill16", "ClassNotebookLogoInverse32", "ClassNotebookLogoInverse16", "StaffNotebookLogo32",
        "StaffNotebookLogoFill32", "StaffNotebookLogo16", "StaffNotebookLogoFill16", "StaffNotebookLogoInverted32", "StaffNotebookLogoInverted16", "KaizalaLogo", "TaskLogo",
        "ProtectionCenterLogo32", "GallatinLogo", "Globe2", "Guitar", "Breakfast", "Brunch", "BeerMug", "Vacation", "Teeth", "Taxi", "Chopsticks", "SyncOccurence", "UnsyncOccurence",
        "GIF", "PrimaryCalendar", "SearchCalendar", "VideoOff", "MicrosoftFlowLogo", "BusinessCenterLogo", "ToDoLogoBottom", "ToDoLogoTop", "EditSolid12", "EditSolidMirrored12",
        "UneditableSolid12", "UneditableSolidMirrored12", "UneditableMirrored", "AdminALogo32", "AdminALogoFill32", "ToDoLogoInverse", "Snooze", "WaffleOffice365", "ImageSearch",
        "NewsSearch", "VideoSearch", "R", "FontColorA", "FontColorSwatch", "LightWeight", "NormalWeight", "SemiboldWeight", "GroupObject", "UngroupObject", "AlignHorizontalLeft",
        "AlignHorizontalCenter", "AlignHorizontalRight", "AlignVerticalTop", "AlignVerticalCenter", "AlignVerticalBottom", "HorizontalDistributeCenter", "VerticalDistributeCenter",
        "Ellipse", "Line", "Octagon", "Hexagon", "Pentagon", "RightTriangle", "HalfCircle", "QuarterCircle", "ThreeQuarterCircle", "SixPointStar", "TwelvePointStar", "ArrangeBringToFront",
        "ArrangeSendToBack", "ArrangeSendBackward", "ArrangeBringForward", "BorderDash", "BorderDot", "LineStyle", "LineThickness", "WindowEdit", "HintText", "MediaAdd", "AnchorLock",
        "AutoHeight", "ChartSeries", "ChartXAngle", "ChartYAngle", "Combobox", "LineSpacing", "Padding", "PaddingTop", "PaddingBottom", "PaddingLeft", "PaddingRight", "NavigationFlipper",
        "AlignJustify", "TextOverflow", "VisualsFolder", "VisualsStore", "PictureCenter", "PictureFill", "PicturePosition", "PictureStretch", "PictureTile", "Slider", "SliderHandleSize",
        "DefaultRatio", "NumberSequence", "GUID", "ReportAdd", "DashboardAdd", "MapPinSolid", "WebPublish", "PieSingleSolid", "BlockedSolid", "DrillDown", "DrillDownSolid", "DrillExpand",
        "DrillShow", "SpecialEvent", "OneDriveFolder16", "FunctionalManagerDashboard", "BIDashboard", "CodeEdit", "RenewalCurrent", "RenewalFuture", "SplitObject", "BulkUpload",
        "DownloadDocument", "GreetingCard", "Flower", "WaitlistConfirm", "WaitlistConfirmMirrored", "LaptopSecure", "DragObject", "EntryView", "EntryDecline", "ContactCardSettings",
        "ContactCardSettingsMirrored", "CalendarSettings", "CalendarSettingsMirrored", "HardDriveLock", "HardDriveUnlock", "AccountManagement", "ReportWarning", "TransitionPop",
        "TransitionPush", "TransitionEffect", "LookupEntities", "ExploreData", "AddBookmark", "SearchBookmark", "DrillThrough", "MasterDatabase", "CertifiedDatabase", "MaximumValue",
        "MinimumValue", "VisualStudioIDELogo32", "PasteAsText", "PasteAsCode", "BrowserTab", "BrowserTabScreenshot", "DesktopScreenshot", "FileYML", "ClipboardSolid", "FabricUserFolder",
        "FabricNetworkFolder", "BullseyeTarget", "AnalyticsView", "Video360Generic", "Untag", "Leave", "Trending12", "Blocked12", "Warning12", "CheckedOutByOther12", "CheckedOutByYou12",
        "CircleShapeSolid", "SquareShapeSolid", "TriangleShapeSolid", "DropShapeSolid", "RectangleShapeSolid", "ZoomToFit", "InsertColumnsLeft", "InsertColumnsRight", "InsertRowsAbove",
        "InsertRowsBelow", "DeleteColumns", "DeleteRows", "DeleteRowsMirrored", "DeleteTable", "AccountBrowser", "VersionControlPush", "StackedColumnChart2", "TripleColumnWide",
        "QuadColumn", "WhiteBoardApp16", "WhiteBoardApp32", "PinnedSolid", "InsertSignatureLine", "ArrangeByFrom", "Phishing", "CreateMailRule", "PublishCourse", "DictionaryRemove",
        "UserRemove", "UserEvent", "Encryption", "PasswordField", "OpenInNewTab", "Hide3", "VerifiedBrandSolid", "MarkAsProtected", "AuthenticatorApp", "WebTemplate", "DefenderTVM",
        "MedalSolid", "D365TalentLearn", "D365TalentInsight", "D365TalentHRCore", "BacklogList", "ButtonControl", "TableGroup", "MountainClimbing", "TagUnknown", "TagUnknownMirror",
        "TagUnknown12", "TagUnknown12Mirror", "Link12", "Presentation", "Presentation12", "Lock12", "BuildDefinition", "ReleaseDefinition", "SaveTemplate", "UserGauge",
        "BlockedSiteSolid12", "TagSolid", "OfficeChat", "OfficeChatSolid", "MailSchedule", "WarningSolid", "Blocked2Solid", "SkypeCircleArrow", "SkypeArrow", "SyncStatus",
        "SyncStatusSolid", "ProjectDocument", "ToDoLogoOutline", "VisioOnlineLogoFill32", "VisioOnlineLogo32", "VisioOnlineLogoCloud32", "VisioDiagramSync", "Event12", "EventDateMissed12",
        "UserOptional", "ResponsesMenu", "DoubleDownArrow", "DistributeDown", "BookmarkReport", "FilterSettings", "GripperDotsVertical", "MailAttached", "AddIn", "LinkedDatabase",
        "TableLink", "PromotedDatabase", "BarChartVerticalFilter", "BarChartVerticalFilterSolid", "MicrosoftTranslatorLogo", "ShowTimeAs", "FileRequest", "WorkItemAlert", "PowerBILogo16",
        "PowerBILogoBackplate16", "BulletedListText", "BulletedListBullet", "BulletedListTextMirrored", "BulletedListBulletMirrored", "NumberedListText", "NumberedListNumber",
        "NumberedListTextMirrored", "NumberedListNumberMirrored", "RemoveLinkChain", "RemoveLinkX", "FabricTextHighlight", "ClearFormattingA", "ClearFormattingEraser", "Photo2Fill",
        "IncreaseIndentText", "IncreaseIndentArrow", "DecreaseIndentText", "DecreaseIndentArrow", "IncreaseIndentTextMirrored", "IncreaseIndentArrowMirrored", "DecreaseIndentTextMirrored",
        "DecreaseIndentArrowMirrored", "CheckListText", "CheckListCheck", "CheckListTextMirrored", "CheckListCheckMirrored", "NumberSymbol", "Coupon", "VerifiedBrand", "ReleaseGate",
        "ReleaseGateCheck", "ReleaseGateError", "M365InvoicingLogo", "RemoveFromShoppingList", "ShieldAlert", "FabricTextHighlightComposite", "Dataflows", "GenericScanFilled",
        "DiagnosticDataBarTooltip", "SaveToMobile", "Orientation2", "ScreenCast", "ShowGrid", "SnapToGrid", "ContactList", "NewMail", "EyeShadow", "FabricFolderConfirm",
        "InformationBarriers", "CommentActive", "ColumnVerticalSectionEdit", "WavingHand", "ShakeDevice", "SmartGlassRemote", "Rotate90Clockwise", "Rotate90CounterClockwise",
        "CampaignTemplate", "ChartTemplate", "PageListFilter", "SecondaryNav", "ColumnVerticalSection", "SkypeCircleSlash", "SkypeSlash", "CustomizeToolbar", "DuplicateRow",
        "RemoveFromTrash", "MailOptions", "Childof", "Footer", "Header", "BarChartVerticalFill", "StackedColumnChart2Fill", "PlainText", "AccessibiltyChecker", "DatabaseSync",
        "ReservationOrders", "TabOneColumn", "TabTwoColumn", "TabThreeColumn", "MicrosoftTranslatorLogoGreen", "MicrosoftTranslatorLogoBlue", "InternalInvestigation", "AddReaction",
        "ContactHeart", "VisuallyImpaired", "EventToDoLogo", "Variable2", "ModelingView", "DisconnectVirtualMachine", "ReportLock", "Uneditable2", "Uneditable2Mirrored",
        "BarChartVerticalEdit", "GlobalNavButtonActive", "PollResults", "Rerun", "QandA", "QandAMirror", "BookAnswers", "AlertSettings", "TrimStart", "TrimEnd", "TableComputed",
        "DecreaseIndentLegacy", "IncreaseIndentLegacy", "SizeLegacy"];

}
