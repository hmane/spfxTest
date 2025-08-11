/**
 * common/index.ts
 * -------------------------------------------------
 * Public surface for the SPFx "common" helpers.
 * Keep this small & intentional so we can evolve internals freely.
 */

// ---- Context hub (primary entry point) ----
export { Context, getSp, getGraph } from './context/pnpjs-config';

// ---- Types commonly needed by consumers ----
export type { SPContext, ContextOptions, CacheTTL, HttpResult } from './utils/types';

// ---- Link helpers (advanced/edge cases) ----
// Most callers should use: Context.getContext().links
export { linksAsync, LinksUnbound } from './utils/links';
export type { LinksBound } from './utils/links';
