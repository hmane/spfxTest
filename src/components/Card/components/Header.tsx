import React, { memo, useMemo, useCallback } from 'react';
import { HeaderProps } from '../Card.types';
import { useCardContext } from '../CardContext';
import { ToolbarButtons } from './ToolbarButtons';
import styles from '../Card.module.scss';

// Chevron Down Icon - Memoized for performance
const ChevronDownIcon = memo<{ className?: string }>(({ className }) => (
	<svg className={className} fill="currentColor" viewBox="0 0 16 16" aria-hidden="true">
		<path
			fillRule="evenodd"
			d="M1.646 4.646a.5.5 0 0 1 .708 0L8 10.293l5.646-5.647a.5.5 0 0 1 .708.708l-6 6a.5.5 0 0 1-.708 0l-6-6a.5.5 0 0 1 0-.708z"
		/>
	</svg>
));

ChevronDownIcon.displayName = 'ChevronDownIcon';

const Header = memo<HeaderProps>(
	({ children, className = '', style, clickable = true, showLoadingShimmer = true, size }) => {
		const {
			variant,
			customHeaderColor,
			allowExpand,
			disabled,
			loading,
			onToggleExpand,
			isExpanded,
			id,
			headerSize,
			toolbarActions = [],
			hideExpandButton = false,
			accessibility = {},
		} = useCardContext();

		const effectiveSize = size || headerSize;

		const headerClasses = useMemo(
			() =>
				[
					styles.header,
					styles[variant],
					styles[`size${effectiveSize.charAt(0).toUpperCase() + effectiveSize.slice(1)}` as keyof typeof styles],
					clickable && allowExpand && !disabled ? styles.clickable : '',
					className,
				]
					.filter(Boolean)
					.join(' '),
			[variant, effectiveSize, clickable, allowExpand, disabled, className]
		);

		const headerStyle = useMemo(
			() => ({
				...(customHeaderColor ? { background: customHeaderColor } : {}),
				...style,
			}),
			[customHeaderColor, style]
		);

		const handleClick = useCallback(() => {
			if (clickable && allowExpand && !disabled) {
				onToggleExpand();
			}
		}, [clickable, allowExpand, disabled, onToggleExpand]);

		const handleKeyDown = useCallback(
			(e: React.KeyboardEvent) => {
				if (clickable && allowExpand && !disabled && (e.key === 'Enter' || e.key === ' ')) {
					e.preventDefault();
					onToggleExpand();
				}
			},
			[clickable, allowExpand, disabled, onToggleExpand]
		);

		const handleExpandClick = useCallback(
			(e: React.MouseEvent) => {
				e.stopPropagation();
				if (allowExpand && !disabled) {
					onToggleExpand();
				}
			},
			[allowExpand, disabled, onToggleExpand]
		);

		const handleExpandKeyDown = useCallback(
			(e: React.KeyboardEvent) => {
				if (e.key === 'Enter' || e.key === ' ') {
					e.preventDefault();
					e.stopPropagation();
					if (allowExpand && !disabled) {
						onToggleExpand();
					}
				}
			},
			[allowExpand, disabled, onToggleExpand]
		);

		return (
			<div
				className={headerClasses}
				style={headerStyle}
				onClick={handleClick}
				role={clickable && allowExpand ? 'button' : undefined}
				tabIndex={clickable && allowExpand && !disabled ? 0 : undefined}
				onKeyDown={handleKeyDown}
				aria-expanded={allowExpand ? isExpanded : undefined}
				aria-controls={allowExpand ? `card-content-${id}` : undefined}
			>
				<div className={styles.headerContent}>
					{loading && showLoadingShimmer && (
						<div
							className={styles.loadingShimmer}
							style={{ width: 20, height: 20, borderRadius: '50%', marginRight: 8 }}
							aria-hidden="true"
						/>
					)}
					<div className={styles.headerText}>{children}</div>
				</div>

				<div className={styles.headerActions}>
					{/* Toolbar buttons appear before expand/collapse button */}
					{toolbarActions.length > 0 && (
						<ToolbarButtons actions={toolbarActions} position="left" showTooltips={true} />
					)}

					{/* Expand/Collapse button in top right */}
					{allowExpand && !hideExpandButton && (
						<button
							type="button"
							className={styles.expandButton}
							onClick={handleExpandClick}
							onKeyDown={handleExpandKeyDown}
							aria-label={
								isExpanded
									? accessibility.collapseButtonLabel || 'Collapse card'
									: accessibility.expandButtonLabel || 'Expand card'
							}
							title={
								isExpanded
									? accessibility.collapseButtonLabel || 'Collapse card'
									: accessibility.expandButtonLabel || 'Expand card'
							}
							disabled={disabled}
						>
							<ChevronDownIcon
								className={`${styles.expandIcon} ${isExpanded ? styles.expanded : ''}`}
							/>
						</button>
					)}
				</div>
			</div>
		);
	}
);

Header.displayName = 'Header';

export { Header };
