import React from 'react';
import { CardState } from './Card.types';

interface CardControllerSubscription {
	cardId: string;
	callback: (action: string, data?: any) => void;
}

interface CardRegistration {
	isExpanded: boolean;
	hasContentLoaded: boolean;
	toggleFn: (source?: 'user' | 'programmatic') => void;
	expandFn: (source?: 'user' | 'programmatic') => void;
	collapseFn: (source?: 'user' | 'programmatic') => void;
	highlightFn?: () => void;
}

class CardController {
	private static instance: CardController;
	private cards = new Map<string, CardRegistration>();
	private subscriptions = new Map<string, CardControllerSubscription[]>();
	private globalSubscriptions: ((action: string, cardId: string, data?: any) => void)[] = [];

	private constructor() {}

	static getInstance(): CardController {
		if (!CardController.instance) {
			CardController.instance = new CardController();
		}
		return CardController.instance;
	}

	// Core registration methods
	registerCard(
		id: string,
		isExpanded: boolean,
		hasContentLoaded: boolean,
		toggleFn: (source?: 'user' | 'programmatic') => void,
		expandFn: (source?: 'user' | 'programmatic') => void,
		collapseFn: (source?: 'user' | 'programmatic') => void,
		highlightFn?: () => void
	): void {
		this.cards.set(id, {
			isExpanded,
			hasContentLoaded,
			toggleFn,
			expandFn,
			collapseFn,
			highlightFn,
		});

		this.notifyGlobalSubscribers('register', id, { isExpanded, hasContentLoaded });
	}

	unregisterCard(id: string): void {
		this.cards.delete(id);
		this.subscriptions.delete(id);
		this.notifyGlobalSubscribers('unregister', id);
	}

	updateCardState(id: string, isExpanded: boolean, hasContentLoaded: boolean): void {
		const card = this.cards.get(id);
		if (card) {
			card.isExpanded = isExpanded;
			card.hasContentLoaded = hasContentLoaded;
			this.notifyGlobalSubscribers('stateUpdate', id, { isExpanded, hasContentLoaded });
		}
	}

	// Subscription methods for class components
	subscribe(cardId: string, callback: (action: string, data?: any) => void): () => void {
		if (!this.subscriptions.has(cardId)) {
			this.subscriptions.set(cardId, []);
		}

		const subscription: CardControllerSubscription = { cardId, callback };
		this.subscriptions.get(cardId)!.push(subscription);

		// Return unsubscribe function
		return () => {
			const subs = this.subscriptions.get(cardId);
			if (subs) {
				const index = subs.indexOf(subscription);
				if (index > -1) {
					subs.splice(index, 1);
				}
			}
		};
	}

	subscribeGlobal(callback: (action: string, cardId: string, data?: any) => void): () => void {
		this.globalSubscriptions.push(callback);

		// Return unsubscribe function
		return () => {
			const index = this.globalSubscriptions.indexOf(callback);
			if (index > -1) {
				this.globalSubscriptions.splice(index, 1);
			}
		};
	}

	private notifySubscribers(cardId: string, action: string, data?: any): void {
		// Notify card-specific subscribers
		const cardSubs = this.subscriptions.get(cardId);
		if (cardSubs) {
			cardSubs.forEach((sub) => sub.callback(action, data));
		}
	}

	private notifyGlobalSubscribers(action: string, cardId: string, data?: any): void {
		// Notify global subscribers
		this.globalSubscriptions.forEach((callback) => callback(action, cardId, data));
	}

	// Public API methods with highlighting support
	expandAll(highlight: boolean = true): void {
		this.cards.forEach((card, id) => {
			if (!card.isExpanded) {
				card.expandFn('programmatic');
				if (highlight && card.highlightFn) {
					card.highlightFn();
				}
				this.notifySubscribers(id, 'expand', { source: 'programmatic' });
			}
		});
		this.notifyGlobalSubscribers('expandAll', 'all', { highlight });
	}

	collapseAll(highlight: boolean = true): void {
		this.cards.forEach((card, id) => {
			if (card.isExpanded) {
				card.collapseFn('programmatic');
				if (highlight && card.highlightFn) {
					card.highlightFn();
				}
				this.notifySubscribers(id, 'collapse', { source: 'programmatic' });
			}
		});
		this.notifyGlobalSubscribers('collapseAll', 'all', { highlight });
	}

	toggleCard(id: string, highlight: boolean = true): boolean {
		const card = this.cards.get(id);
		if (card) {
			card.toggleFn('programmatic');
			if (highlight && card.highlightFn) {
				card.highlightFn();
			}
			this.notifySubscribers(id, 'toggle', {
				source: 'programmatic',
				newState: !card.isExpanded,
			});
			this.notifyGlobalSubscribers('toggle', id, { highlight, newState: !card.isExpanded });
			return true;
		}
		return false;
	}

	expandCard(id: string, highlight: boolean = true): boolean {
		const card = this.cards.get(id);
		if (card && !card.isExpanded) {
			card.expandFn('programmatic');
			if (highlight && card.highlightFn) {
				card.highlightFn();
			}
			this.notifySubscribers(id, 'expand', { source: 'programmatic' });
			this.notifyGlobalSubscribers('expand', id, { highlight });
			return true;
		}
		return false;
	}

	collapseCard(id: string, highlight: boolean = true): boolean {
		const card = this.cards.get(id);
		if (card && card.isExpanded) {
			card.collapseFn('programmatic');
			if (highlight && card.highlightFn) {
				card.highlightFn();
			}
			this.notifySubscribers(id, 'collapse', { source: 'programmatic' });
			this.notifyGlobalSubscribers('collapse', id, { highlight });
			return true;
		}
		return false;
	}

	highlightCard(id: string): boolean {
		const card = this.cards.get(id);
		if (card && card.highlightFn) {
			card.highlightFn();
			this.notifySubscribers(id, 'highlight');
			this.notifyGlobalSubscribers('highlight', id);
			return true;
		}
		return false;
	}

	isCardExpanded(id: string): boolean {
		const card = this.cards.get(id);
		return card ? card.isExpanded : false;
	}

	isCardRegistered(id: string): boolean {
		return this.cards.has(id);
	}

	getCardStates(): CardState[] {
		return Array.from(this.cards.entries()).map(([id, card]) => ({
			id,
			isExpanded: card.isExpanded,
			hasContentLoaded: card.hasContentLoaded,
		}));
	}

	getCardState(id: string): CardState | null {
		const card = this.cards.get(id);
		if (card) {
			return {
				id,
				isExpanded: card.isExpanded,
				hasContentLoaded: card.hasContentLoaded,
			};
		}
		return null;
	}

	getExpandedCards(): string[] {
		return Array.from(this.cards.entries())
			.filter(([_, card]) => card.isExpanded)
			.map(([id, _]) => id);
	}

	getCollapsedCards(): string[] {
		return Array.from(this.cards.entries())
			.filter(([_, card]) => !card.isExpanded)
			.map(([id, _]) => id);
	}

	// Batch operations for performance
	batchOperation(
		operations: Array<{ cardId: string; action: 'expand' | 'collapse' | 'toggle' }>,
		highlight: boolean = true
	): void {
		operations.forEach(({ cardId, action }) => {
			switch (action) {
				case 'expand':
					this.expandCard(cardId, highlight);
					break;
				case 'collapse':
					this.collapseCard(cardId, highlight);
					break;
				case 'toggle':
					this.toggleCard(cardId, highlight);
					break;
			}
		});
		this.notifyGlobalSubscribers('batchOperation', 'multiple', { operations, highlight });
	}

	// Utility methods
	getRegisteredCardIds(): string[] {
		return Array.from(this.cards.keys());
	}

	getCardCount(): number {
		return this.cards.size;
	}

	getExpandedCardCount(): number {
		return Array.from(this.cards.values()).filter((card) => card.isExpanded).length;
	}

	clearAllCards(): void {
		const cardIds = Array.from(this.cards.keys());
		cardIds.forEach((id) => this.unregisterCard(id));
		this.notifyGlobalSubscribers('clearAll', 'all');
	}
}

// Export singleton instance and base class
export const cardController = CardController.getInstance();

// Class Component Base for easy integration
export class CardControllerComponent extends React.Component {
	protected cardController = cardController;
	private unsubscribers: (() => void)[] = [];

	// Subscribe to card events
	protected subscribeToCard(cardId: string, callback: (action: string, data?: any) => void): void {
		const unsubscribe = this.cardController.subscribe(cardId, callback);
		this.unsubscribers.push(unsubscribe);
	}

	// Subscribe to all card events
	protected subscribeToAllCards(
		callback: (action: string, cardId: string, data?: any) => void
	): void {
		const unsubscribe = this.cardController.subscribeGlobal(callback);
		this.unsubscribers.push(unsubscribe);
	}

	componentWillUnmount() {
		// Clean up subscriptions
		this.unsubscribers.forEach((unsubscribe) => unsubscribe());
	}
}

// Class Component Helper - HOC for easy integration
export interface WithCardControllerProps {
	cardController: typeof cardController;
}

export function withCardController<P extends WithCardControllerProps>(
	WrappedComponent: React.ComponentType<P>
): React.ComponentType<Omit<P, keyof WithCardControllerProps>> {
	return class extends React.Component<Omit<P, keyof WithCardControllerProps>> {
		render() {
			return React.createElement(WrappedComponent, {
				...(this.props as P),
				cardController: cardController,
			});
		}
	};
}
