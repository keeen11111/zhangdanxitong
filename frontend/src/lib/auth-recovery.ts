export type UnauthorizedRecoveryActions = {
  clear: () => void;
  redirect: (target: string) => void;
};

/** Clears an invalid local session and restores the sign-in entry point. */
export function recoverFromUnauthorized(
  status: number,
  currentPath: string,
  actions: UnauthorizedRecoveryActions,
): boolean {
  if (status !== 401) return false;

  actions.clear();
  if (currentPath !== "/login") actions.redirect("/login");
  return true;
}
