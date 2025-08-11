import React, { memo, useMemo, useEffect } from 'react';
import { ContentProps } from '../Card.types';
import { useCardContext } from '../CardContext';
import styles from '../Card.module.scss';

const Content = memo<ContentProps>(
	({
		children,
		className = '',
		style,
		padding = 'medium',
		loadingPlaceholder,
		errorBoundary = true,
	}) => {
		const {
			isExpanded,
			allowExpand,
			id,
			lazyLoad,
			hasContentLoaded,
			loading,
			onContentLoad,
			disableAnimation,
		} = useCardContext();

		const contentClasses = useMemo(
			() =>
				[
					styles.content,
					isExpanded ? styles.expanded : styles.collapsed,
					disableAnimation ? styles.noAnimation : '',
					className,
				]
					.filter(Boolean)
					.join(' '),
			[isExpanded, disableAnimation, className]
		);

		const bodyClasses = useMemo(
			() =>
				[styles.body, styles[`padding${padding.charAt(0).toUpperCase() + padding.slice(1)}` as keyof typeof styles]]
					.filter(Boolean)
					.join(' '),
			[padding]
		);

		// Handle lazy loading
		const shouldRenderContent = !lazyLoad || hasContentLoaded;
		const isContentFunction = typeof children === 'function';

		const contentToRender = useMemo(() => {
			if (!shouldRenderContent) {
				return (
					loadingPlaceholder || (
						<div className={styles.loadingShimmer} style={{ height: 100, borderRadius: 4 }} />
					)
				);
			} else if (loading && !isContentFunction) {
				return (
					loadingPlaceholder || (
						<div className={styles.loadingShimmer} style={{ height: 100, borderRadius: 4 }} />
					)
				);
			} else {
				return isContentFunction ? (children as () => React.ReactNode)() : children;
			}
		}, [shouldRenderContent, loading, isContentFunction, children, loadingPlaceholder]);

		// Notify when content loads for the first time
		useEffect(() => {
			if (shouldRenderContent && lazyLoad && isExpanded) {
				onContentLoad();
			}
		}, [shouldRenderContent, lazyLoad, isExpanded, onContentLoad]);

		return (
			<div
				className={contentClasses}
				style={style}
				id={`card-content-${id}`}
				aria-hidden={!isExpanded}
			>
				<div className={bodyClasses}>{contentToRender}</div>
			</div>
		);
	}
);

Content.displayName = 'Content';

export { Content };
