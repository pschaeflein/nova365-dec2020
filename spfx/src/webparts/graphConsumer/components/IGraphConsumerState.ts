import { User } from '@microsoft/microsoft-graph-types';

export interface IGraphConsumerState {
  users: Array<User>;
  searchFor: string;
}
