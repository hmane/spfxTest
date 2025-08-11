// utils/theme.ts
import { ITheme, loadTheme } from '@fluentui/style-utilities';

type ThemeListener = (t: ITheme) => void;

export class ThemeBus {
	private listeners: ThemeListener[] = [];
	private current?: ITheme;

	snapshot() {
		return this.current;
	}

	set(t: ITheme) {
		this.current = t;
		loadTheme(t as any);
		for (const l of this.listeners) l(t);
	}

	onChange(cb: ThemeListener) {
		this.listeners.push(cb);
		return () => {
			this.listeners = this.listeners.filter((x) => x !== cb);
		};
	}
}
