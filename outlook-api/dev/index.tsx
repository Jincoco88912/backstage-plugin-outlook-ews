import React from 'react';
import { createDevApp } from '@backstage/dev-utils';
import { outlookApiPlugin, OutlookApiPage } from '../src/plugin';

createDevApp()
  .registerPlugin(outlookApiPlugin)
  .addPage({
    element: <OutlookApiPage />,
    title: 'Root Page',
    path: '/outlook-api',
  })
  .render();
