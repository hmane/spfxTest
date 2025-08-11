import React, {
	memo,
	useMemo,
	useEffect,
	useState,
	createContext,
	useContext,
	useRef,
} from 'react';
import { FieldGroupProps, FieldGroupContextType } from '../Field.types';
import styles from '../Field.module.scss';

// FieldGroup Context
const FieldGroupContext = createContext<FieldGroupContextType | null>(null);

export const useFieldGroupContext = () => {
	const context = useContext(FieldGroupContext);
	return context; // Can be null, it's optional
};

export const FieldGroup = memo<FieldGroupProps>(
	({
		id,
		children,
		labelWidth = 'auto',
		className = '',
		style,
		spacing = 'normal',
		layout = 'horizontal',
		disabled = false,
	}) => {
		const [calculatedLabelWidth, setCalculatedLabelWidth] = useState<string>('150px');
		const groupRef = useRef<HTMLDivElement>(null);

		// Auto-calculate label width if set to 'auto'
		useEffect(() => {
			if (labelWidth === 'auto' && groupRef.current) {
				const calculateOptimalWidth = () => {
					const labels = groupRef.current?.querySelectorAll(`.${styles.fieldLabel}`);
					if (!labels || labels.length === 0) return;

					let maxWidth = 0;
					labels.forEach((label) => {
						const labelElement = label as HTMLElement;
						const textElement = labelElement.querySelector(`.${styles.labelText}`) as HTMLElement;
						if (textElement) {
							// Create temporary element to measure text width
							const tempDiv = document.createElement('div');
							tempDiv.style.position = 'absolute';
							tempDiv.style.visibility = 'hidden';
							tempDiv.style.height = 'auto';
							tempDiv.style.width = 'auto';
							tempDiv.style.whiteSpace = 'nowrap';
							tempDiv.style.fontSize = window.getComputedStyle(textElement).fontSize;
							tempDiv.style.fontFamily = window.getComputedStyle(textElement).fontFamily;
							tempDiv.style.fontWeight = window.getComputedStyle(textElement).fontWeight;
							tempDiv.textContent = textElement.textContent;

							document.body.appendChild(tempDiv);
							const width = tempDiv.offsetWidth;
							document.body.removeChild(tempDiv);

							maxWidth = Math.max(maxWidth, width);
						}
					});

					if (maxWidth > 0) {
						// Add some padding and set as CSS custom property
						const optimalWidth = `${maxWidth + 20}px`;
						setCalculatedLabelWidth(optimalWidth);

						// Set CSS custom property for all fields in this group
						if (groupRef.current) {
							groupRef.current.style.setProperty('--field-label-width', optimalWidth);
						}
					}
				};

				// Calculate on mount and when children change
				const timer = setTimeout(calculateOptimalWidth, 100);
				return () => clearTimeout(timer);
			} else if (typeof labelWidth === 'string' || typeof labelWidth === 'number') {
				// Set explicit label width
				const widthValue = typeof labelWidth === 'number' ? `${labelWidth}px` : labelWidth;
				if (groupRef.current) {
					groupRef.current.style.setProperty('--field-label-width', widthValue);
				}
			}
		}, [labelWidth, children]);

		const groupClasses = useMemo(
			() =>
				[
					styles.fieldGroup,
					styles[
						`spacing${spacing.charAt(0).toUpperCase() + spacing.slice(1)}` as keyof typeof styles
					],
					styles[
						`layout${layout.charAt(0).toUpperCase() + layout.slice(1)}` as keyof typeof styles
					],
					disabled ? styles.disabled : '',
					className,
				]
					.filter(Boolean)
					.join(' '),
			[spacing, layout, labelWidth, disabled, className]
		);

		const contextValue = useMemo(
			(): FieldGroupContextType => ({
				labelWidth,
				spacing,
				layout,
				disabled,
			}),
			[labelWidth, spacing, layout, disabled]
		);

		return (
			<FieldGroupContext.Provider value={contextValue}>
				<div ref={groupRef} className={groupClasses} style={style} data-field-group-id={id}>
					{children}
				</div>
			</FieldGroupContext.Provider>
		);
	}
);

FieldGroup.displayName = 'FieldGroup';
