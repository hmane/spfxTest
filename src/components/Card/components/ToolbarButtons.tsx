import { memo, useCallback, useMemo } from 'react';
import styles from '../Card.module.scss';
import { ToolbarAction, ToolbarButtonsProps } from '../Card.types';
import { useCardContext } from '../CardContext';

const ToolbarButtons = memo<ToolbarButtonsProps>(
	({
		actions = [],
		className = '',
		style,
		position = 'left',
		showTooltips = true,
		stackOnMobile = false,
	}) => {
		const { onToolbarAction, disabled } = useCardContext();

		const toolbarClasses = useMemo(
			() =>
				[
					styles.toolbarButtons,
					styles[
						`position${position.charAt(0).toUpperCase() + position.slice(1)}` as keyof typeof styles
					],
					stackOnMobile ? styles.stackMobile : '',
					className,
				]
					.filter(Boolean)
					.join(' '),
			[position, stackOnMobile, className]
		);

		const renderToolbarButton = useCallback(
			(action: ToolbarAction) => {
				const buttonClasses = [
					styles.toolbarButton,
					action.variant ? styles[action.variant] : styles.default,
					action.hideOnMobile ? (styles as any).hideOnMobile : '',
				]
					.filter(Boolean)
					.join(' ');

				const isMobile = typeof window !== 'undefined' && window.innerWidth <= 768;

				return (
					<button
						key={action.id}
						type="button"
						className={buttonClasses}
						onClick={(e) => onToolbarAction(action, e)}
						disabled={action.disabled || disabled}
						aria-label={action.ariaLabel || action.label}
						title={showTooltips ? action.tooltip : undefined}
					>
						{action.icon && (
							<span className="buttonIcon" aria-hidden="true">
								{isMobile && action.mobileIcon ? action.mobileIcon : action.icon}
							</span>
						)}
						<span className={styles.buttonText}>{action.label}</span>
					</button>
				);
			},
			[onToolbarAction, disabled, showTooltips]
		);

		if (actions.length === 0) {
			return null;
		}

		return (
			<div className={toolbarClasses} style={style}>
				{actions.map(renderToolbarButton)}
			</div>
		);
	}
);

ToolbarButtons.displayName = 'ToolbarButtons';

export { ToolbarButtons };
