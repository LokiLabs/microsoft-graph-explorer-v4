import { GraphAuthenticationProvider } from './GraphAuthenticationProvider';
import { GraphApplicationAuthenticationProvider } from './GraphApplicationAuthenticationProvider';

export { GraphClient } from './graph-client';
export const authProvider = new GraphAuthenticationProvider();
export const appAuthProvider = new GraphApplicationAuthenticationProvider();
