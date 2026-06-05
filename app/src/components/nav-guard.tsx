'use client';

import { createContext, useContext } from 'react';

/**
 * A navigation guard lets a screen intercept attempts to navigate away (e.g.
 * sidebar clicks) so it can warn about unsaved changes first.
 *
 * The guard receives a `proceed` callback that performs the actual navigation.
 * Return `true` to intercept — the guard then takes responsibility for calling
 * `proceed()` later (e.g. after the user confirms in a modal). Return `false`
 * to let navigation happen immediately.
 */
export type NavGuardFn = (proceed: () => void) => boolean;

type RegisterNavGuard = (fn: NavGuardFn | null) => void;

const NavGuardContext = createContext<RegisterNavGuard>(() => {});

export const NavGuardProvider = NavGuardContext.Provider;

/** Register (or clear, with `null`) the active navigation guard. */
export function useNavGuard(): RegisterNavGuard {
  return useContext(NavGuardContext);
}
