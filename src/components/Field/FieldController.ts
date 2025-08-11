import React from 'react';
import { FieldRegistration, FieldHierarchy } from './Field.types';

interface FieldControllerSubscription {
	fieldId: string;
	callback: (action: string, data?: any) => void;
}

class FieldController {
	private static instance: FieldController;
	private fields = new Map<string, FieldRegistration>();
	private subscriptions = new Map<string, FieldControllerSubscription[]>();
	private globalSubscriptions: ((action: string, fieldId: string, data?: any) => void)[] = [];

	private constructor() {}

	static getInstance(): FieldController {
		if (!FieldController.instance) {
			FieldController.instance = new FieldController();
		}
		return FieldController.instance;
	}

	// Core registration methods
	registerField(id: string, registration: FieldRegistration): void {
		this.fields.set(id, registration);
		this.notifyGlobalSubscribers('register', id, registration.hierarchy);
	}

	unregisterField(id: string): void {
		this.fields.delete(id);
		this.subscriptions.delete(id);
		this.notifyGlobalSubscribers('unregister', id);
	}

	// Subscription methods for class components
	subscribe(fieldId: string, callback: (action: string, data?: any) => void): () => void {
		if (!this.subscriptions.has(fieldId)) {
			this.subscriptions.set(fieldId, []);
		}

		const subscription: FieldControllerSubscription = { fieldId, callback };
		this.subscriptions.get(fieldId)!.push(subscription);

		// Return unsubscribe function
		return () => {
			const subs = this.subscriptions.get(fieldId);
			if (subs) {
				const index = subs.indexOf(subscription);
				if (index > -1) {
					subs.splice(index, 1);
				}
			}
		};
	}

	subscribeGlobal(callback: (action: string, fieldId: string, data?: any) => void): () => void {
		this.globalSubscriptions.push(callback);

		// Return unsubscribe function
		return () => {
			const index = this.globalSubscriptions.indexOf(callback);
			if (index > -1) {
				this.globalSubscriptions.splice(index, 1);
			}
		};
	}

	private notifySubscribers(fieldId: string, action: string, data?: any): void {
		// Notify field-specific subscribers
		const fieldSubs = this.subscriptions.get(fieldId);
		if (fieldSubs) {
			fieldSubs.forEach((sub) => sub.callback(action, data));
		}
	}

	private notifyGlobalSubscribers(action: string, fieldId: string, data?: any): void {
		// Notify global subscribers
		this.globalSubscriptions.forEach((callback) => callback(action, fieldId, data));
	}

	// Focus management methods
	focusField(id: string): boolean {
		const field = this.fields.get(id);
		if (field && field.focusFn) {
			try {
				const success = field.focusFn();
				if (success) {
					this.notifySubscribers(id, 'focus');
					this.notifyGlobalSubscribers('focus', id);
				}
				return success;
			} catch (error) {
				console.warn(`Failed to focus field ${id}:`, error);
				return false;
			}
		}
		return false;
	}

	scrollToField(id: string): boolean {
		const field = this.fields.get(id);
		if (field && field.scrollFn) {
			try {
				const success = field.scrollFn();
				if (success) {
					this.notifySubscribers(id, 'scroll');
					this.notifyGlobalSubscribers('scroll', id);
				}
				return success;
			} catch (error) {
				console.warn(`Failed to scroll to field ${id}:`, error);
				return false;
			}
		}
		return false;
	}

	// Validation and field state methods
	getAllFields(): string[] {
		return Array.from(this.fields.keys());
	}

	getFieldHierarchy(fieldId: string): FieldHierarchy | null {
		const field = this.fields.get(fieldId);
		return field ? field.hierarchy : null;
	}

	getFieldsByCard(cardId: string): string[] {
		return Array.from(this.fields.entries())
			.filter(([_, field]) => field.hierarchy.cardId === cardId)
			.map(([id, _]) => id);
	}

	getFieldsByAccordion(accordionId: string): string[] {
		return Array.from(this.fields.entries())
			.filter(([_, field]) => field.hierarchy.accordionId === accordionId)
			.map(([id, _]) => id);
	}

	getFieldsByAccordionItem(accordionItemId: string): string[] {
		return Array.from(this.fields.entries())
			.filter(([_, field]) => field.hierarchy.accordionItemId === accordionItemId)
			.map(([id, _]) => id);
	}

	// Focus management for groups
	focusFirstFieldInCard(cardId: string): boolean {
		const fields = this.getFieldsByCard(cardId);
		if (fields.length > 0) {
			return this.focusField(fields[0]);
		}
		return false;
	}

	focusFirstFieldInAccordion(accordionId: string): boolean {
		const fields = this.getFieldsByAccordion(accordionId);
		if (fields.length > 0) {
			return this.focusField(fields[0]);
		}
		return false;
	}

	focusFirstFieldInAccordionItem(accordionItemId: string): boolean {
		const fields = this.getFieldsByAccordionItem(accordionItemId);
		if (fields.length > 0) {
			return this.focusField(fields[0]);
		}
		return false;
	}

	// Batch operations
	focusFieldsInSequence(fieldIds: string[], delay: number = 100): Promise<boolean[]> {
		return new Promise((resolve) => {
			const results: boolean[] = [];
			let index = 0;

			const focusNext = () => {
				if (index < fieldIds.length) {
					const result = this.focusField(fieldIds[index]);
					results.push(result);
					index++;
					setTimeout(focusNext, delay);
				} else {
					resolve(results);
				}
			};

			focusNext();
		});
	}

	scrollToFieldsInSequence(fieldIds: string[], delay: number = 500): Promise<boolean[]> {
		return new Promise((resolve) => {
			const results: boolean[] = [];
			let index = 0;

			const scrollNext = () => {
				if (index < fieldIds.length) {
					const result = this.scrollToField(fieldIds[index]);
					results.push(result);
					index++;
					setTimeout(scrollNext, delay);
				} else {
					resolve(results);
				}
			};

			scrollNext();
		});
	}

	// Utility methods
	isFieldRegistered(fieldId: string): boolean {
		return this.fields.has(fieldId);
	}

	getRegisteredFieldCount(): number {
		return this.fields.size;
	}

	clearAllFields(): void {
		const fieldIds = Array.from(this.fields.keys());
		fieldIds.forEach((id) => this.unregisterField(id));
		this.notifyGlobalSubscribers('clearAll', 'all');
	}

	// Get field statistics
	getFieldStats(): {
		totalFields: number;
		fieldsByCard: { [cardId: string]: number };
		fieldsByAccordion: { [accordionId: string]: number };
		fieldsWithoutCard: number;
	} {
		const stats = {
			totalFields: this.fields.size,
			fieldsByCard: {} as { [cardId: string]: number },
			fieldsByAccordion: {} as { [accordionId: string]: number },
			fieldsWithoutCard: 0,
		};

		this.fields.forEach((field, fieldId) => {
			const { cardId, accordionId } = field.hierarchy;

			if (cardId) {
				stats.fieldsByCard[cardId] = (stats.fieldsByCard[cardId] || 0) + 1;
			} else {
				stats.fieldsWithoutCard++;
			}

			if (accordionId) {
				stats.fieldsByAccordion[accordionId] = (stats.fieldsByAccordion[accordionId] || 0) + 1;
			}
		});

		return stats;
	}
}

// Export singleton instance
export const fieldController = FieldController.getInstance();

// Class Component Base for easy integration
export class FieldControllerComponent extends React.Component {
	protected fieldController = fieldController;
	private unsubscribers: (() => void)[] = [];

	// Subscribe to field events
	protected subscribeToField(
		fieldId: string,
		callback: (action: string, data?: any) => void
	): void {
		const unsubscribe = this.fieldController.subscribe(fieldId, callback);
		this.unsubscribers.push(unsubscribe);
	}

	// Subscribe to all field events
	protected subscribeToAllFields(
		callback: (action: string, fieldId: string, data?: any) => void
	): void {
		const unsubscribe = this.fieldController.subscribeGlobal(callback);
		this.unsubscribers.push(unsubscribe);
	}

	componentWillUnmount() {
		// Clean up subscriptions
		this.unsubscribers.forEach((unsubscribe) => unsubscribe());
	}
}

// HOC for easy integration
export interface WithFieldControllerProps {
	fieldController: typeof fieldController;
}

export function withFieldController<P extends WithFieldControllerProps>(
	WrappedComponent: React.ComponentType<P>
): React.ComponentType<Omit<P, keyof WithFieldControllerProps>> {
	return class extends React.Component<Omit<P, keyof WithFieldControllerProps>> {
		render() {
			return React.createElement(WrappedComponent, {
				...(this.props as P),
				fieldController: fieldController,
			});
		}
	};
}
