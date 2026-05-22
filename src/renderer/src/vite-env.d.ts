/// <reference types="vite/client" />

import type { PwdSafeAgentApi } from "../../shared/types";

declare global {
  interface Window {
    pwdSafeAgent?: PwdSafeAgentApi;
  }
}
