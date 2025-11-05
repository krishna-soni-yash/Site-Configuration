import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'SiteConfigApplicationCustomizerStrings';
import { createPnpSpfx } from '../schema/Initialization';
import deployWebParts from '../schema/WebPart Deployment/Deployment';

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

    return Promise.resolve();
  }
}
