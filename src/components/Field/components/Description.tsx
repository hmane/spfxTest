import React, { memo, useMemo } from 'react';
import { DescriptionProps } from '../Field.types';
import styles from '../Field.module.scss';

export const Description = memo<DescriptionProps>(
	({ children, variant = 'help', className = '', style }) => {
		const descriptionClasses = useMemo(
			() =>
				[
					styles.fieldDescription,
					styles[`variant${variant.charAt(0).toUpperCase() + variant.slice(1)}` as keyof typeof styles],
					className,
				]
					.filter(Boolean)
					.join(' '),
			[variant, className]
		);

		return (
			<div className={descriptionClasses} style={style} role="note">
				{children}
			</div>
		);
	}
);

Description.displayName = 'Description';
