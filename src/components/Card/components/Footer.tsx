import React, { memo, useMemo } from 'react';
import { FooterProps } from '../Card.types';
import styles from '../Card.module.scss';

const Footer = memo<FooterProps>(
	({
		children,
		className = '',
		style,
		backgroundColor,
		borderTop = true,
		padding = 'medium',
		textAlign = 'left',
	}) => {
		const footerClasses = useMemo(
			() =>
				[
					styles.footer,
					styles[`padding${padding.charAt(0).toUpperCase() + padding.slice(1)}` as keyof typeof styles],
					styles[`text${textAlign.charAt(0).toUpperCase() + textAlign.slice(1)}` as keyof typeof styles],
					!borderTop ? styles.noBorder : '',
					className,
				]
					.filter(Boolean)
					.join(' '),
			[padding, textAlign, borderTop, className]
		);

		const footerStyle = useMemo(
			() => ({
				...(backgroundColor && { backgroundColor }),
				...style,
			}),
			[backgroundColor, style]
		);

		return (
			<div className={footerClasses} style={footerStyle}>
				{children}
			</div>
		);
	}
);

Footer.displayName = 'Footer';

export { Footer };
