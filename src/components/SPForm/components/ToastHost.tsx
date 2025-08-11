import * as React from 'react';
import { MessageBar, MessageBarType, Layer, Stack } from '@fluentui/react';

export type ToastKind = 'success' | 'error' | 'warning' | 'info';

export interface Toast {
	id: string;
	kind: ToastKind;
	text: string;
	timeoutMs?: number;
}

type Ctx = {
	push: (t: Omit<Toast, 'id'>) => void;
};

const ToastCtx = React.createContext<Ctx | null>(null);

export const useToasts = () => {
	const ctx = React.useContext(ToastCtx);
	if (!ctx) throw new Error('useToasts must be used within <ToastHost>');
	return ctx;
};

export const ToastHost: React.FC<{ children: React.ReactNode }> = ({ children }) => {
	const [toasts, setToasts] = React.useState<Toast[]>([]);

	const push = React.useCallback((t: Omit<Toast, 'id'>) => {
		const id = crypto.randomUUID();
		const toast: Toast = { id, ...t };
		setToasts((prev) => [...prev, toast]);
		if (toast.timeoutMs !== 0) {
			const ms = toast.timeoutMs ?? 4000;
			window.setTimeout(() => {
				setToasts((prev) => prev.filter((x) => x.id !== id));
			}, ms);
		}
	}, []);

	const remove = (id: string) => setToasts((prev) => prev.filter((x) => x.id !== id));

	const kindMap: Record<ToastKind, MessageBarType> = {
		success: MessageBarType.success,
		error: MessageBarType.error,
		warning: MessageBarType.warning,
		info: MessageBarType.info,
	};

	return (
		<ToastCtx.Provider value={{ push }}>
			{children}
			<Layer>
				<Stack
					tokens={{ childrenGap: 8 }}
					styles={{
						root: {
							position: 'fixed',
							right: 16,
							top: 16,
							zIndex: 100000,
							width: 420,
							maxWidth: 'calc(100vw - 32px)',
						},
					}}
				>
					{toasts.map((t) => (
						<MessageBar
							key={t.id}
							messageBarType={kindMap[t.kind]}
							isMultiline={false}
							onDismiss={() => remove(t.id)}
						>
							{t.text}
						</MessageBar>
					))}
				</Stack>
			</Layer>
		</ToastCtx.Provider>
	);
};
