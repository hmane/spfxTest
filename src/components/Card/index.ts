import { useMemo } from 'react';

export { Card } from './Card';
export { Content } from './components/Content';
export { Footer } from './components/Footer';
export { Header } from './components/Header';
export { ToolbarButtons } from './components/ToolbarButtons';

// Context and hooks
export { CardContext, useCardContext } from './CardContext';

// Controller and utilities
export { cardController, CardControllerComponent, withCardController } from './CardController';

// Types
export type {
	CardContextType, CardEventData, CardEventType, CardProps, CardState,
	CardVariant, ContentProps,
	FooterProps, HeaderProps, HeaderSize, ToolbarAction, ToolbarButtonsProps
} from './Card.types';

// Import WithCardControllerProps from the controller file, not types
export type { WithCardControllerProps } from './CardController';

// Hooks for functional components
import { cardController } from './CardController';

export const useCardController = () => {
	return useMemo(
		() => ({
			expandAll: (highlight?: boolean) => cardController.expandAll(highlight),
			collapseAll: (highlight?: boolean) => cardController.collapseAll(highlight),
			toggleCard: (id: string, highlight?: boolean) => cardController.toggleCard(id, highlight),
			expandCard: (id: string, highlight?: boolean) => cardController.expandCard(id, highlight),
			collapseCard: (id: string, highlight?: boolean) => cardController.collapseCard(id, highlight),
			highlightCard: (id: string) => cardController.highlightCard(id),
			getCardStates: () => cardController.getCardStates(),
			getCardState: (id: string) => cardController.getCardState(id),
			isCardExpanded: (id: string) => cardController.isCardExpanded(id),
			isCardRegistered: (id: string) => cardController.isCardRegistered(id),
			getExpandedCards: () => cardController.getExpandedCards(),
			getCollapsedCards: () => cardController.getCollapsedCards(),
			getRegisteredCardIds: () => cardController.getRegisteredCardIds(),
			getCardCount: () => cardController.getCardCount(),
			getExpandedCardCount: () => cardController.getExpandedCardCount(),
			subscribe: (cardId: string, callback: (action: string, data?: any) => void) =>
				cardController.subscribe(cardId, callback),
			subscribeGlobal: (callback: (action: string, cardId: string, data?: any) => void) =>
				cardController.subscribeGlobal(callback),
			batchOperation: (
				operations: Array<{ cardId: string; action: 'expand' | 'collapse' | 'toggle' }>,
				highlight?: boolean
			) => cardController.batchOperation(operations, highlight),
		}),
		[]
	);
};
