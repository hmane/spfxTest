import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import styles from './Card.module.scss';
import { CardContextType, CardEventData, CardProps, ToolbarAction } from './Card.types';
import { CardContext } from './CardContext';
import { cardController } from './CardController';

// Performance utilities
const useDebounce = (callback: () => void, delay: number) => {
	const timeoutRef = useRef<NodeJS.Timeout>();

	return useCallback(() => {
		if (timeoutRef.current) {
			clearTimeout(timeoutRef.current);
		}
		timeoutRef.current = setTimeout(callback, delay);
	}, [callback, delay]);
};

const useMemoizedCallback = <T extends (...args: any[]) => any>(callback: T, deps: any[]): T => {
	return useCallback(callback, deps);
};

// Main Card Component
export const Card: React.FC<CardProps> = ({
	id,
	defaultExpanded = false,
	allowExpand = true,
	variant = 'default',
	headerSize = 'regular',
	customHeaderColor,
	loading = false,
	loadingMessage = 'Loading...',
	showLoadingOverlay = false,
	lazyLoad = false,
	highlightOnProgrammaticChange = true,
	highlightDuration = 600,
	highlightColor,
	toolbarActions = [],
	hideExpandButton = false,
	onExpand,
	onCollapse,
	onDataLoaded,
	onContentLoad,
	onCardEvent,
	className = '',
	style,
	disableAnimation = false,
	elevation = 2,
	disabled = false,
	theme,
	accessibility = {},
	performance = {},
	children,
}) => {
	const [isExpanded, setIsExpanded] = useState(defaultExpanded);
	const [hasContentLoaded, setHasContentLoaded] = useState(!lazyLoad || defaultExpanded);
	const [hasDataLoaded, setHasDataLoaded] = useState(false);
	const [isHighlighted, setIsHighlighted] = useState(false);

	const previousLoadingRef = useRef(loading);
	const cardRef = useRef<HTMLDivElement>(null);
	const highlightTimeoutRef = useRef<NodeJS.Timeout>();
	const renderCountRef = useRef(0);

	// Performance tracking
	useEffect(() => {
		renderCountRef.current += 1;
		if (process.env.NODE_ENV === 'development') {
			console.debug(`Card ${id} rendered ${renderCountRef.current} times`);
		}
	});

	// Debounced toggle function
	const debouncedToggle = useDebounce(() => {
		handleToggleExpand('user');
	}, performance.debounceToggle || 0);

	// Highlight function
	const highlightCard = useCallback(() => {
		if (!highlightOnProgrammaticChange) return;

		setIsHighlighted(true);

		if (highlightTimeoutRef.current) {
			clearTimeout(highlightTimeoutRef.current);
		}

		highlightTimeoutRef.current = setTimeout(() => {
			setIsHighlighted(false);
		}, highlightDuration);
	}, [highlightOnProgrammaticChange, highlightDuration]);

	// Memoized expand/collapse functions
	const expandFn = useMemoizedCallback(
		(source: 'user' | 'programmatic' = 'programmatic') => {
			if (!isExpanded && allowExpand && !disabled) {
				setIsExpanded(true);
				if (lazyLoad && !hasContentLoaded) {
					setHasContentLoaded(true);
				}
				const eventData: CardEventData = {
					cardId: id,
					isExpanded: true,
					timestamp: Date.now(),
					source,
				};
				onExpand?.(eventData);
				onCardEvent?.('expand', eventData);
			}
		},
		[isExpanded, allowExpand, disabled, lazyLoad, hasContentLoaded, id, onExpand, onCardEvent]
	);

	const collapseFn = useMemoizedCallback(
		(source: 'user' | 'programmatic' = 'programmatic') => {
			if (isExpanded && allowExpand && !disabled) {
				setIsExpanded(false);
				const eventData: CardEventData = {
					cardId: id,
					isExpanded: false,
					timestamp: Date.now(),
					source,
				};
				onCollapse?.(eventData);
				onCardEvent?.('collapse', eventData);
			}
		},
		[isExpanded, allowExpand, disabled, id, onCollapse, onCardEvent]
	);

	const toggleFn = useMemoizedCallback(
		(source: 'user' | 'programmatic' = 'programmatic') => {
			if (isExpanded) {
				collapseFn(source);
			} else {
				expandFn(source);
			}
		},
		[isExpanded, expandFn, collapseFn]
	);

	// Register card with controller
	useEffect(() => {
		cardController.registerCard(
			id,
			isExpanded,
			hasContentLoaded,
			toggleFn,
			expandFn,
			collapseFn,
			highlightCard
		);

		return () => {
			cardController.unregisterCard(id);
			if (highlightTimeoutRef.current) {
				clearTimeout(highlightTimeoutRef.current);
			}
		};
	}, [id, isExpanded, hasContentLoaded, toggleFn, expandFn, collapseFn, highlightCard]);

	// Update controller when state changes
	useEffect(() => {
		cardController.updateCardState(id, isExpanded, hasContentLoaded);
	}, [id, isExpanded, hasContentLoaded]);

	// Handle loading state changes
	useEffect(() => {
		if (previousLoadingRef.current && !loading && !hasDataLoaded) {
			setHasDataLoaded(true);
			const eventData: CardEventData = {
				cardId: id,
				isExpanded,
				timestamp: Date.now(),
				source: 'user',
			};
			onDataLoaded?.(eventData);
			onCardEvent?.('contentLoad', eventData);
		}
		previousLoadingRef.current = loading;
	}, [loading, hasDataLoaded, onDataLoaded, onCardEvent, id, isExpanded]);

	// Handle content loading for lazy loading
	useEffect(() => {
		if (lazyLoad && isExpanded && !hasContentLoaded) {
			setHasContentLoaded(true);
			const eventData: CardEventData = {
				cardId: id,
				isExpanded,
				timestamp: Date.now(),
				source: 'user',
			};
			onContentLoad?.(eventData);
			onCardEvent?.('contentLoad', eventData);
		}
	}, [lazyLoad, isExpanded, hasContentLoaded, onContentLoad, onCardEvent, id]);

	// Handle expand/collapse
	const handleToggleExpand = useMemoizedCallback(
		(source: 'user' | 'programmatic' = 'user') => {
			if (!allowExpand || disabled) return;

			const newExpanded = !isExpanded;
			setIsExpanded(newExpanded);

			// Load content if lazy loading and expanding
			if (lazyLoad && newExpanded && !hasContentLoaded) {
				setHasContentLoaded(true);
			}

			const eventData: CardEventData = {
				cardId: id,
				isExpanded: newExpanded,
				timestamp: Date.now(),
				source,
			};

			if (newExpanded) {
				onExpand?.(eventData);
				onCardEvent?.('expand', eventData);
			} else {
				onCollapse?.(eventData);
				onCardEvent?.('collapse', eventData);
			}
		},
		[
			allowExpand,
			disabled,
			isExpanded,
			lazyLoad,
			hasContentLoaded,
			id,
			onExpand,
			onCollapse,
			onCardEvent,
		]
	);

	// Handle toolbar action click
	const handleToolbarAction = useMemoizedCallback(
		(action: ToolbarAction, event: React.MouseEvent) => {
			event.stopPropagation();
			if (!action.disabled && !disabled) {
				action.onClick(id);
			}
		},
		[disabled, id]
	);

	// Handle content load callback
	const handleContentLoad = useMemoizedCallback(() => {
		const eventData: CardEventData = {
			cardId: id,
			isExpanded,
			timestamp: Date.now(),
			source: 'user',
		};
		onContentLoad?.(eventData);
		onCardEvent?.('contentLoad', eventData);
	}, [id, isExpanded, onContentLoad, onCardEvent]);

	// Memoized styles and classes
	const cardStyle = useMemo(
		() => ({
			...style,
			...(theme?.backgroundColor && { backgroundColor: theme.backgroundColor }),
			...(theme?.borderColor && { borderColor: theme.borderColor }),
			...(theme?.textColor && { color: theme.textColor }),
			...(highlightColor &&
				isHighlighted && {
					borderColor: highlightColor,
					boxShadow: `0 0 0 2px ${highlightColor}33`,
				}),
		}),
		[style, theme, highlightColor, isHighlighted]
	);

	const cardClasses = useMemo(
		() =>
			[
				styles.card,
				styles[`elevation${elevation}`],
				disabled ? styles.disabled : '',
				isHighlighted ? styles.highlight : '',
				disableAnimation ? styles.noAnimation : '',
				className,
			]
				.filter(Boolean)
				.join(' '),
		[elevation, disabled, isHighlighted, disableAnimation, className]
	);

	// Memoized context value
	const contextValue = useMemo(
		(): CardContextType => ({
			id,
			isExpanded,
			allowExpand,
			disabled,
			loading,
			variant,
			customHeaderColor,
			lazyLoad,
			hasContentLoaded,
			headerSize,
			toolbarActions,
			hideExpandButton,
			accessibility,
			disableAnimation,
			onToggleExpand: performance.debounceToggle
				? debouncedToggle
				: () => handleToggleExpand('user'),
			onToolbarAction: handleToolbarAction,
			onContentLoad: handleContentLoad,
		}),
		[
			id,
			isExpanded,
			allowExpand,
			disabled,
			loading,
			variant,
			customHeaderColor,
			lazyLoad,
			hasContentLoaded,
			headerSize,
			toolbarActions,
			hideExpandButton,
			accessibility,
			disableAnimation,
			performance.debounceToggle,
			debouncedToggle,
			handleToggleExpand,
			handleToolbarAction,
			handleContentLoad,
		]
	);

	const cardProps = useMemo(
		() => ({
			className: cardClasses,
			style: cardStyle,
			ref: cardRef,
			...(accessibility.region && {
				role: 'region',
				'aria-labelledby': accessibility.labelledBy,
				'aria-describedby': accessibility.describedBy,
			}),
		}),
		[cardClasses, cardStyle, accessibility]
	);

	return (
		<CardContext.Provider value={contextValue}>
			<div {...cardProps}>
				{/* Loading Overlay */}
				{loading && showLoadingOverlay && (
					<div className={styles.loadingOverlay}>
						<div
							className={styles.loadingSpinner}
							aria-label={accessibility.loadingLabel || loadingMessage}
						/>
						<div className={styles.loadingText}>{loadingMessage}</div>
					</div>
				)}

				{children}
			</div>
		</CardContext.Provider>
	);
};

Card.displayName = 'Card';
