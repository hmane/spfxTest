import { ReactNode, CSSProperties } from 'react';

export type CardVariant = 'success' | 'error' | 'warning' | 'info' | 'default';
export type HeaderSize = 'compact' | 'regular' | 'large';

export interface ToolbarAction {
	id: string;
	label: string;
	icon?: ReactNode;
	onClick: (cardId?: string) => void;
	disabled?: boolean;
	variant?: 'default' | 'primary' | 'secondary' | 'danger';
	tooltip?: string;
	hideOnMobile?: boolean;
	mobileIcon?: ReactNode;
	ariaLabel?: string;
}

export interface CardEventData {
	cardId: string;
	isExpanded: boolean;
	timestamp: number;
	source: 'user' | 'programmatic';
}

export type CardEventType = 'expand' | 'collapse' | 'contentLoad' | 'programmaticToggle';

export interface CardContextType {
	id: string;
	isExpanded: boolean;
	allowExpand: boolean;
	disabled: boolean;
	loading: boolean;
	variant: CardVariant;
	customHeaderColor?: string;
	lazyLoad: boolean;
	hasContentLoaded: boolean;
	headerSize: HeaderSize;
	toolbarActions: ToolbarAction[];
	hideExpandButton: boolean;
	accessibility: {
		expandButtonLabel?: string;
		collapseButtonLabel?: string;
		loadingLabel?: string;
		region?: boolean;
		labelledBy?: string;
		describedBy?: string;
	};
	disableAnimation: boolean;
	onToggleExpand: (source?: 'user' | 'programmatic') => void;
	onToolbarAction: (action: ToolbarAction, event: React.MouseEvent) => void;
	onContentLoad: () => void;
}


export interface CardProps {
	/** Unique identifier for the card */
	id: string;

	/** Whether the card is expanded by default */
	defaultExpanded?: boolean;

	/** Whether the card can be collapsed/expanded */
	allowExpand?: boolean;

	/** Card header background variant */
	variant?: CardVariant;

	/** Header size - affects padding and font size */
	headerSize?: HeaderSize;

	/** Custom background color for header (overrides variant) */
	customHeaderColor?: string;

	/** Loading state */
	loading?: boolean;

	/** Loading message */
	loadingMessage?: string;

	/** Whether to show loading overlay */
	showLoadingOverlay?: boolean;

	/** Enable lazy loading - content loads only when expanded */
	lazyLoad?: boolean;

	/** Highlight border on programmatic changes */
	highlightOnProgrammaticChange?: boolean;

	/** Duration of highlight effect in milliseconds */
	highlightDuration?: number;

	/** Custom highlight color */
	highlightColor?: string;

	/** Toolbar actions that appear before expand/collapse button */
	toolbarActions?: ToolbarAction[];

	/** Hide the expand/collapse button */
	hideExpandButton?: boolean;

	/** Callback when card is expanded */
	onExpand?: (data: CardEventData) => void;

	/** Callback when card is collapsed */
	onCollapse?: (data: CardEventData) => void;

	/** Callback when card data is loaded */
	onDataLoaded?: (data: CardEventData) => void;

	/** Callback when content is loaded for first time (lazy loading) */
	onContentLoad?: (data: CardEventData) => void;

	/** Global event listener for card events */
	onCardEvent?: (type: CardEventType, data: CardEventData) => void;

	/** Custom CSS class */
	className?: string;

	/** Custom styles */
	style?: CSSProperties;

	/** Disable animations */
	disableAnimation?: boolean;

	/** Card elevation/shadow level */
	elevation?: 1 | 2 | 3 | 4 | 5;

	/** Whether card is disabled */
	disabled?: boolean;

	/** Custom theme overrides */
	theme?: {
		primaryColor?: string;
		backgroundColor?: string;
		textColor?: string;
		borderColor?: string;
	};

	/** Accessibility options */
	accessibility?: {
		expandButtonLabel?: string;
		collapseButtonLabel?: string;
		loadingLabel?: string;
		region?: boolean;
		labelledBy?: string;
		describedBy?: string;
	};

	/** Performance options */
	performance?: {
		debounceToggle?: number;
		virtualizeContent?: boolean;
		preloadThreshold?: number;
		memoizeContent?: boolean;
	};

	/** Children components */
	children: ReactNode;
}

export interface HeaderProps {
	children: ReactNode;
	className?: string;
	style?: CSSProperties;
	clickable?: boolean;
	showLoadingShimmer?: boolean;
	size?: HeaderSize;
}

export interface ToolbarButtonsProps {
	actions: ToolbarAction[];
	className?: string;
	style?: CSSProperties;
	position?: 'left' | 'right';
	showTooltips?: boolean;
	stackOnMobile?: boolean;
}

export interface ContentProps {
	children: ReactNode | (() => ReactNode);
	className?: string;
	style?: CSSProperties;
	padding?: 'none' | 'small' | 'medium' | 'large';
	loadingPlaceholder?: ReactNode;
	errorBoundary?: boolean;
}

export interface FooterProps {
	children: ReactNode;
	className?: string;
	style?: CSSProperties;
	backgroundColor?: string;
	borderTop?: boolean;
	padding?: 'none' | 'small' | 'medium' | 'large';
	textAlign?: 'left' | 'center' | 'right';
}

export interface CardState {
	id: string;
	isExpanded: boolean;
	hasContentLoaded: boolean;
}
