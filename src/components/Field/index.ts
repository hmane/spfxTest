import React, { useMemo } from 'react';

export { Field } from './Field';
export { Label } from './components/Label';
export { Description } from './components/Description';
export { Error } from './components/Error';
export { FieldGroup } from './components/FieldGroup';

// Context and hooks
export { FieldContext, useFieldContext } from './Field';
export { useFieldGroupContext } from './components/FieldGroup';

// Controller and utilities
export { fieldController, withFieldController, FieldControllerComponent } from './FieldController';

// Types
export type {
	FieldProps,
	LabelProps,
	DescriptionProps,
	ErrorProps,
	FieldGroupProps,
	FieldContextType,
	FieldGroupContextType,
	ValidationState,
	FieldHierarchy,
	FieldRegistration,
} from './Field.types';

// Import WithFieldControllerProps from the controller file
export type { WithFieldControllerProps } from './FieldController';

// Hook for functional components
import { fieldController } from './FieldController';

export const useFieldController = () => {
	return useMemo(
		() => ({
			focusField: (id: string) => fieldController.focusField(id),
			scrollToField: (id: string) => fieldController.scrollToField(id),
			getAllFields: () => fieldController.getAllFields(),
			getFieldHierarchy: (fieldId: string) => fieldController.getFieldHierarchy(fieldId),
			getFieldsByCard: (cardId: string) => fieldController.getFieldsByCard(cardId),
			getFieldsByAccordion: (accordionId: string) =>
				fieldController.getFieldsByAccordion(accordionId),
			getFieldsByAccordionItem: (accordionItemId: string) =>
				fieldController.getFieldsByAccordionItem(accordionItemId),
			focusFirstFieldInCard: (cardId: string) => fieldController.focusFirstFieldInCard(cardId),
			focusFirstFieldInAccordion: (accordionId: string) =>
				fieldController.focusFirstFieldInAccordion(accordionId),
			focusFirstFieldInAccordionItem: (accordionItemId: string) =>
				fieldController.focusFirstFieldInAccordionItem(accordionItemId),
			focusFieldsInSequence: (fieldIds: string[], delay?: number) =>
				fieldController.focusFieldsInSequence(fieldIds, delay),
			scrollToFieldsInSequence: (fieldIds: string[], delay?: number) =>
				fieldController.scrollToFieldsInSequence(fieldIds, delay),
			isFieldRegistered: (fieldId: string) => fieldController.isFieldRegistered(fieldId),
			getRegisteredFieldCount: () => fieldController.getRegisteredFieldCount(),
			getFieldStats: () => fieldController.getFieldStats(),
			subscribe: (fieldId: string, callback: (action: string, data?: any) => void) =>
				fieldController.subscribe(fieldId, callback),
			subscribeGlobal: (callback: (action: string, fieldId: string, data?: any) => void) =>
				fieldController.subscribeGlobal(callback),
		}),
		[]
	);
};
