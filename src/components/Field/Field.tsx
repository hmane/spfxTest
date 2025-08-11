import React, { useEffect, useRef, useMemo, useCallback, createContext, useContext } from 'react';
import { useController } from 'react-hook-form';
import { FieldProps, FieldContextType, ValidationState } from './Field.types';
import { fieldController } from './FieldController';
import styles from './Field.module.scss';

// Field Context
const FieldContext = createContext<FieldContextType | null>(null);

export const useFieldContext = () => {
	const context = useContext(FieldContext);
	if (!context) {
		throw new Error('Field child components must be used within a Field component');
	}
	return context;
};

// Main Field Component
export const Field: React.FC<FieldProps> = ({
	id,
	name,
	control,
	rules,
	layout = 'horizontal',
	disabled = false,
	cardId,
	accordionId,
	accordionItemId,
	className = '',
	style,
	onValidationChange,
	onFocus,
	autoFocus = false,
	scrollOnFocus = true,
	children,
}) => {
	const fieldRef = useRef<HTMLDivElement>(null);
	const fieldId = id || name;

	// React Hook Form integration
	const {
		field,
		fieldState: { invalid, error, isDirty, isTouched },
		formState,
	} = useController({
		name,
		control,
		rules,
	});

	// Auto-detect validation state from React Hook Form
	const validationState: ValidationState = useMemo(
		() => ({
			isValid: !invalid,
			error,
			isDirty,
			isTouched,
		}),
		[invalid, error, isDirty, isTouched]
	);

	// Focus function for navigation controller
	const focusField = useCallback(() => {
		if (fieldRef.current) {
			const input = fieldRef.current.querySelector(
				'input, select, textarea, button, [tabindex]'
			) as HTMLElement;
			if (input) {
				input.focus();

				// Scroll into view if configured
				if (scrollOnFocus) {
					input.scrollIntoView({
						behavior: 'smooth',
						block: 'center',
						inline: 'nearest',
					});
				}

				onFocus?.();
				return true;
			}
		}
		return false;
	}, [scrollOnFocus, onFocus]);

	// Scroll function for navigation controller
	const scrollToField = useCallback(() => {
		if (fieldRef.current) {
			fieldRef.current.scrollIntoView({
				behavior: 'smooth',
				block: 'center',
				inline: 'nearest',
			});
			return true;
		}
		return false;
	}, []);

	// Auto-register with field controller and navigation
	useEffect(() => {
		fieldController.registerField(fieldId, {
			element: fieldRef.current || undefined,
			focusFn: focusField,
			scrollFn: scrollToField,
			hierarchy: {
				cardId,
				accordionId,
				accordionItemId,
				fieldName: name,
			},
		});

		return () => {
			fieldController.unregisterField(fieldId);
		};
	}, [fieldId, name, cardId, accordionId, accordionItemId, focusField, scrollToField]);

	// Notify validation changes
	useEffect(() => {
		if (onValidationChange) {
			const errors = error ? [error.message || 'Validation error'] : [];
			onValidationChange(validationState.isValid, errors);
		}
	}, [validationState.isValid, error, onValidationChange]);

	// Auto focus on mount
	useEffect(() => {
		if (autoFocus) {
			const timer = setTimeout(() => {
				focusField();
			}, 100); // Small delay to ensure component is rendered

			return () => clearTimeout(timer);
		}
	}, [autoFocus, focusField]);

	// Detect if field has label
	const hasLabel = useMemo(() => {
		return React.Children.toArray(children).some(
			(child) => React.isValidElement(child) && (child.type as any)?.displayName === 'Label'
		);
	}, [children]);

	// Field classes
	const fieldClasses = useMemo(
		() =>
			[
				styles.field,
				styles[`layout${layout.charAt(0).toUpperCase() + layout.slice(1)}` as keyof typeof styles],
				!hasLabel ? styles.noLabel : '',
				disabled ? styles.disabled : '',
				className,
			]
				.filter(Boolean)
				.join(' '),
		[layout, hasLabel, disabled, className]
	);

	// Context value for child components
	const fieldContext = useMemo(
		(): FieldContextType => ({
			fieldName: name,
			fieldId,
			validationState,
			value: field.value,
			onChange: field.onChange,
			onBlur: field.onBlur,
			disabled,
			cardId,
			accordionId,
			accordionItemId,
		}),
		[name, fieldId, validationState, field, disabled, cardId, accordionId, accordionItemId]
	);

	return (
		<FieldContext.Provider value={fieldContext}>
			<div
				ref={fieldRef}
				className={fieldClasses}
				style={style}
				data-field-name={name}
				data-field-id={fieldId}
			>
				{children}
			</div>
		</FieldContext.Provider>
	);
};

Field.displayName = 'Field';

export { FieldContext };
