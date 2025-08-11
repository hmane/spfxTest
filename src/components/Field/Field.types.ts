import { ReactNode, CSSProperties } from 'react';
import { Control, RegisterOptions, FieldError } from 'react-hook-form';

export interface ValidationState {
	isValid: boolean;
	error?: FieldError;
	isDirty: boolean;
	isTouched: boolean;
}

export interface FieldContextType {
	fieldName: string;
	fieldId: string;
	validationState: ValidationState;
	value: any;
	onChange: (value: any) => void;
	onBlur: () => void;
	disabled?: boolean;
	cardId?: string;
	accordionId?: string;
	accordionItemId?: string;
}

export interface FieldProps {
	/** Field identifier for navigation */
	id?: string;

	/** React Hook Form field name */
	name: string;

	/** React Hook Form control */
	control?: Control<any>;

	/** React Hook Form validation rules */
	rules?: RegisterOptions;

	/** Layout direction */
	layout?: 'horizontal' | 'vertical';

	/** Whether field is disabled */
	disabled?: boolean;

	/** Component hierarchy for navigation */
	cardId?: string;
	accordionId?: string;
	accordionItemId?: string;

	/** Custom CSS class */
	className?: string;

	/** Custom styles */
	style?: CSSProperties;

	/** Validation change callback */
	onValidationChange?: (isValid: boolean, errors: string[]) => void;

	/** Focus callback */
	onFocus?: () => void;

	/** Auto focus on mount */
	autoFocus?: boolean;

	/** Scroll when focused externally */
	scrollOnFocus?: boolean;

	/** Children components */
	children: ReactNode;
}

export interface LabelProps {
	children: ReactNode;
	required?: boolean;
	info?: string | React.ReactNode;
	htmlFor?: string;
	wrap?: 'normal' | 'break-word' | 'nowrap';
	maxWidth?: string | number;
	className?: string;
	style?: CSSProperties;
}

export interface DescriptionProps {
	children: ReactNode;
	variant?: 'help' | 'info' | 'warning';
	className?: string;
	style?: CSSProperties;
}

export interface ErrorProps {
	children?: ReactNode;
	animation?: 'slide' | 'fade' | 'none';
	position?: 'below' | 'inline';
	className?: string;
	style?: CSSProperties;
}

export interface FieldGroupProps {
	id?: string;
	children: ReactNode;
	labelWidth?: string | number | 'auto';
	className?: string;
	style?: CSSProperties;
	spacing?: 'compact' | 'normal' | 'relaxed';
	layout?: 'horizontal' | 'vertical';
	disabled?: boolean;
}

export interface FieldGroupContextType {
	labelWidth: string | number | 'auto';
	spacing: 'compact' | 'normal' | 'relaxed';
	layout: 'horizontal' | 'vertical';
	disabled?: boolean;
}

// Navigation integration types
export interface FieldHierarchy {
	cardId?: string;
	accordionId?: string;
	accordionItemId?: string;
	fieldName: string;
}

export interface FieldRegistration {
	element?: HTMLElement;
	focusFn: () => boolean; // Changed from void to boolean
	scrollFn: () => boolean; // Changed from void to boolean
	hierarchy: FieldHierarchy;
}
