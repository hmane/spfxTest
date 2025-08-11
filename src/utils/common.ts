/**
 * Deep clone an object (simple implementation)
 */
export const deepClone = <T>(obj: T): T => {
	if (obj === null || typeof obj !== 'object') {
		return obj;
	}

	if (obj instanceof Date) {
		return new Date(obj.getTime()) as unknown as T;
	}

	if (obj instanceof Array) {
		return obj.map((item) => deepClone(item)) as unknown as T;
	}

	if (typeof obj === 'object') {
		const clonedObj = {} as T;
		for (const key in obj) {
			if (obj.hasOwnProperty(key)) {
				clonedObj[key] = deepClone(obj[key]);
			}
		}
		return clonedObj;
	}

	return obj;
};

/**
 * Debounce function to limit function calls
 */
export const debounce = <T extends (...args: any[]) => any>(
	func: T,
	wait: number
): ((...args: Parameters<T>) => void) => {
	let timeout: number | undefined;

	return (...args: Parameters<T>) => {
		if (timeout !== undefined) {
			clearTimeout(timeout);
		}
		timeout = window.setTimeout(() => func(...args), wait);
	};
};

/**
 * Throttle function to limit function calls
 */
export const throttle = <T extends (...args: any[]) => any>(
	func: T,
	limit: number
): ((...args: Parameters<T>) => void) => {
	let inThrottle: boolean;

	return (...args: Parameters<T>) => {
		if (!inThrottle) {
			func(...args);
			inThrottle = true;
			setTimeout(() => (inThrottle = false), limit);
		}
	};
};

/**
 * Check if value is null or undefined
 */
export const isNullOrUndefined = (value: any): value is null | undefined => {
	return value === null || value === undefined;
};

/**
 * Check if string is null, undefined, or empty
 */
export const isNullOrEmpty = (value: string | null | undefined): value is null | undefined | '' => {
	return isNullOrUndefined(value) || value === '';
};

/**
 * Check if string is null, undefined, empty, or whitespace
 */
export const isNullOrWhiteSpace = (
	value: string | null | undefined
): value is null | undefined | '' => {
	return isNullOrEmpty(value) || value!.trim() === '';
};

/**
 * Safe string trimming
 */
export const safeTrim = (value: string | null | undefined): string => {
	return isNullOrUndefined(value) ? '' : value!.toString().trim();
};

/**
 * Convert string to title case
 */
export const toTitleCase = (str: string): string => {
	return str.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
};

/**
 * Truncate string with ellipsis
 */
export const truncateString = (str: string, maxLength: number, suffix: string = '...'): string => {
	if (str.length <= maxLength) {
		return str;
	}
	return str.substr(0, maxLength - suffix.length) + suffix;
};

/**
 * Generate random string
 */
export const generateRandomString = (length: number = 8): string => {
	const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
	let result = '';
	for (let i = 0; i < length; i++) {
		result += chars.charAt(Math.floor(Math.random() * chars.length));
	}
	return result;
};

/**
 * Format number with commas
 */
export const formatNumber = (num: number): string => {
	return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
};

/**
 * Safe JSON parse with default value
 */
export const safeJsonParse = <T>(jsonString: string, defaultValue: T): T => {
	try {
		return JSON.parse(jsonString);
	} catch {
		return defaultValue;
	}
};

/**
 * Get nested property value safely
 */
export const getNestedProperty = (obj: any, path: string): any => {
	return path.split('.').reduce((current, prop) => {
		return current && current[prop] !== undefined ? current[prop] : undefined;
	}, obj);
};

/**
 * Set nested property value safely
 */
export const setNestedProperty = (obj: any, path: string, value: any): void => {
	const keys = path.split('.');
	const lastKey = keys.pop()!;
	const target = keys.reduce((current, key) => {
		if (!current[key] || typeof current[key] !== 'object') {
			current[key] = {};
		}
		return current[key];
	}, obj);
	target[lastKey] = value;
};

/**
 * Compare two arrays for equality (shallow)
 */
export const arrayEquals = <T>(a: T[], b: T[]): boolean => {
	if (a.length !== b.length) return false;
	return a.every((val, index) => val === b[index]);
};

/**
 * Remove duplicates from array
 */
export const uniqueArray = <T>(array: T[]): T[] => {
	return [...new Set(array)];
};

/**
 * Chunk array into smaller arrays
 */
export const chunkArray = <T>(array: T[], size: number): T[][] => {
	const chunks: T[][] = [];
	for (let i = 0; i < array.length; i += size) {
		chunks.push(array.slice(i, i + size));
	}
	return chunks;
};
