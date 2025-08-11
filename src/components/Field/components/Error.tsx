import React, { memo, useMemo } from 'react';
import { ErrorProps } from '../Field.types';
import { useFieldContext } from '../Field';
import styles from '../Field.module.scss';

export const Error = memo<ErrorProps>(
	({ children, animation = 'slide', position = 'below', className = '', style }) => {
		const { validationState } = useFieldContext();

		const errorClasses = useMemo(
			() =>
				[
					styles.fieldError,
					styles[`animation${animation.charAt(0).toUpperCase() + animation.slice(1)}` as keyof typeof styles],
					styles[`position${position.charAt(0).toUpperCase() + position.slice(1)}` as keyof typeof styles],
					className,
				]
					.filter(Boolean)
					.join(' '),
			[animation, position, className]
		);

		const errorMessage = children || validationState.error?.message;
		const showError = !validationState.isValid && validationState.isTouched;

		if (!showError || !errorMessage) {
			return null;
		}

		return (
			<div className={errorClasses} style={style} role="alert" aria-live="polite">
				<span className={styles.errorIcon} aria-hidden="true">
					⚠️
				</span>
				<span>{errorMessage}</span>
			</div>
		);
	}
);

Error.displayName = 'Error';
