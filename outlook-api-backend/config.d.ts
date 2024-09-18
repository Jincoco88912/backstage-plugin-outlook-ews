import { Config as BackstageConfig } from '@backstage/config';

export interface Config extends BackstageConfig {
  exampleOutlook: {
    /**
     * @visibility backend
     */
    sessionSecret: string;

    /**
     * @visibility backend
     */
    aesSecretKey: string;

    /**
     * @visibility backend
     */
    redis: {
      /**
       * @visibility backend
       */
      host: string;

      /**
       * @visibility backend
       */
      port: number;
    };
  };
}