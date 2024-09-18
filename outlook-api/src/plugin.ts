import {
  createPlugin,
  createRoutableExtension,
} from '@backstage/core-plugin-api';

import { rootRouteRef } from './routes';

export const outlookApiPlugin = createPlugin({
  id: 'outlook-api',
  routes: {
    root: rootRouteRef,
  },
});

export const OutlookApiPage = outlookApiPlugin.provide(
  createRoutableExtension({
    name: 'OutlookApiPage',
    component: () =>
      import('./components/OutlookMailComponents').then(m => m.OutlookComponent),
    mountPoint: rootRouteRef,
  }),
);

export const OutlookCalendarPage = outlookApiPlugin.provide(
  createRoutableExtension({
    name: 'OutlookCalendarPage',
    component: () =>
      import('./components/OutlookCalendarComponent').then(m => m.OutlookCalendarComponent),
    mountPoint: rootRouteRef,
  }),
);
