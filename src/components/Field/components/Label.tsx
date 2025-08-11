import React, { memo, useMemo, useCallback } from 'react';
import { LabelProps } from '../Field.types';
import { useFieldContext } from '../Field';
import styles from '../Field.module.scss';

export const Label = memo<LabelProps>(
	({
		children,
		required = false,
		info,
		htmlFor,
		wrap = 'normal',
		maxWidth,
		className = '',
		style,
	}) => {
		const { fieldId, fieldName } = useFieldContext();

		const labelClasses = useMemo(
			() =>
				[
					styles.fieldLabel,
					styles[`wrap${wrap.charAt(0).toUpperCase() + wrap.slice(1)}` as keyof typeof styles],
					className,
				]
					.filter(Boolean)
					.join(' '),
			[wrap, className]
		);

		const labelStyle = useMemo(
			() => ({
				...(maxWidth && { maxWidth }),
				...style,
			}),
			[maxWidth, style]
		);

		const handleInfoClick = useCallback(
			(e: React.MouseEvent) => {
				e.preventDefault();
				e.stopPropagation();
				if (typeof info === 'string') {
					alert(info); // Simple implementation, could be enhanced
				}
			},
			[info]
		);

		return (
			<label className={labelClasses} style={labelStyle} htmlFor={htmlFor || fieldId || fieldName}>
				<span className={styles.labelText}>{children}</span>
				{required && (
					<span
						className={styles.requiredIndicator}
						aria-label="required"
						title="This field is required"
					>
						*
					</span>
				)}
				{info && (
					<button
						type="button"
						className={styles.infoButton}
						onClick={handleInfoClick}
						aria-label="Additional information"
						title={typeof info === 'string' ? info : 'Additional information'}
					>
						?
					</button>
				)}
			</label>
		);
	}
);

Label.displayName = 'Label';
