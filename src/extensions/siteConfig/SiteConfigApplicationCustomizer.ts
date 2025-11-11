import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'SiteConfigApplicationCustomizerStrings';
import { createPnpSpfx } from './Initialization';
import deployWebParts from '../schema/WebPart Deployment/Deployment';
import { provisionRequiredLists } from '../schema/List Provision/RequiredListProvision';

const LOG_SOURCE: string = 'SiteConfigApplicationCustomizer';

export interface ISiteConfigApplicationCustomizerProperties {
  testMessage: string;
}

export default class SiteConfigApplicationCustomizer
  extends BaseApplicationCustomizer<ISiteConfigApplicationCustomizerProperties> {

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    const sp = createPnpSpfx(this.context as any);
    const spAny = sp as any;

    //--------------------Page Web Part Deployment--------------------//
    const webInfo: { IsRootWeb?: boolean; Title?: string; ServerRelativeUrl?: string; IsSubWeb?: boolean } = await spAny.web
      .select('IsRootWeb', 'IsSubWeb', 'Title', 'ServerRelativeUrl')();

    if (webInfo?.IsRootWeb || webInfo?.IsSubWeb === false) {
      return;
    }

    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }

    try {
      await deployWebParts(sp as any);
    } catch (e) {
      Log.error(LOG_SOURCE, e as any);
    }

    //---------------Required Lists Provisioning--------------------//
    try {
      const siteUrl = this.context.pageContext?.web?.absoluteUrl || 'unknown-site';
      const storageKey = `requiredListsProvisioned:${siteUrl}`;
      const storageAvailable = typeof window !== 'undefined' && !!window.sessionStorage;

      const alreadyProvisioned = storageAvailable ? window.sessionStorage.getItem(storageKey) : null;

      if (!alreadyProvisioned) {
        await provisionRequiredLists(sp);
        if (storageAvailable) {
          try {
            window.sessionStorage.setItem(storageKey, new Date().toISOString());
          } catch (e) {
            console.warn('Could not write provisioning flag to sessionStorage', e);
          }
        }
      }
    } catch (err) {
      console.error('Error while provisioning required lists:', err);
    }
    
    return Promise.resolve();
  }
}
