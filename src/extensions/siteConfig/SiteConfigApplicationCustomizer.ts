/*eslint-disable*/
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { createPnpSpfx } from './Initialization';
import deployWebParts from '../schema/WebPart Deployment/Deployment';
//import { provisionRequiredLists } from '../schema/List Provision/RequiredListProvision';
import {
  ProjectDocumentsLibraryName,
  provisionProjectDocumentsLibrary
} from '../schema/Library Provsion/libraries/ProjectDocuments';

const LOG_SOURCE: string = 'SiteConfigApplicationCustomizer';

export interface ISiteConfigApplicationCustomizerProperties {
  testMessage: string;
}

export default class SiteConfigApplicationCustomizer
  extends BaseApplicationCustomizer<ISiteConfigApplicationCustomizerProperties> {

  public async onInit(): Promise<void> {

    const sp = createPnpSpfx(this.context as any);

    const PARENT_SITE_NAME = "SMARTIQ Plus";
    //to check current site name
    const currentSiteName = this.context.pageContext.web.title;
    
    if (currentSiteName !== PARENT_SITE_NAME) {
      
      //---------------Required Lists Provisioning--------------------//
      // try {
      //   const siteUrl = this.context.pageContext?.web?.absoluteUrl || 'unknown-site';
      //   const storageKey = `requiredListsProvisioned:${siteUrl}`;
      //   const storageAvailable = typeof window !== 'undefined' && !!window.sessionStorage;

      //   const alreadyProvisioned = storageAvailable ? window.sessionStorage.getItem(storageKey) : null;

      //   if (!alreadyProvisioned) {
      //     await provisionRequiredLists(sp);
      //     if (storageAvailable) {
      //       try {
      //         window.sessionStorage.setItem(storageKey, new Date().toISOString());
      //       } catch (e) {
      //         console.warn('Could not write provisioning flag to sessionStorage', e);
      //       }
      //     }
      //   }
      // } catch (err) {
      //   console.error('Error while provisioning required lists:', err);
      // }

      //--------------------Page Web Part Deployment--------------------//
      let message: string = this.properties.testMessage;
      if (!message) {
        message = '(No properties were provided.)';
      }

      try {
        await deployWebParts(sp as any);
      } catch (e) {
        Log.error(LOG_SOURCE, e as any);
      }

      //--------------------Provision Document Libraries--------------------//
      try {
        let projectDocumentsExists = false;
        try {
          await sp.web.lists.getByTitle(ProjectDocumentsLibraryName).select('Id')();
          projectDocumentsExists = true;
        } catch (error) {
          projectDocumentsExists = false;
        }

        if (!projectDocumentsExists) {
          await provisionProjectDocumentsLibrary(sp);
        }
      } catch (err) {
        console.error('Error while provisioning document libraries:', err);
      }
    }

    return Promise.resolve();
  }
}
