import {
  IAdaptiveCardAction,
  IComponentDefinition,
  IDataSourceDefinition,
  IExtensibilityLibrary,
  ILayoutDefinition,
  IQueryModifierDefinition,
  ISuggestionProviderDefinition,
  LayoutType,
  LayoutRenderType,
  BaseLayout
} from "@pnp/modern-search-extensibility";
import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { IPropertyPaneField } from "@microsoft/sp-property-pane";
import { DigestCheckboxWebComponent } from "./digestCheckboxComponent";
import { DigestComposeTriggerWebComponent } from "./digestComposeTriggerComponent";

// News Digest layout: a stacked, full-width list of News articles, each
// with a checkbox that adds/removes it from a shared in-memory selection
// store (digestSelectionStore.ts), and a persistent floating trigger that
// opens a compose dialog for whatever's currently selected. Deliberately
// does not use PnP Modern Search's built-in itemSelectionProps — that
// mechanism is for connecting two web parts' filters together, not for
// collecting full item records for an external action like sending email.
export interface INewsDigestLayoutProperties {
  // Reserved for future toggles if any part of this layout needs to
  // become independently switchable per web part instance.
}

export class NewsDigestLayout extends BaseLayout<INewsDigestLayoutProperties> {

  // Static field, not created inline inside getCustomLayouts() — a fresh
  // ServiceKey object minted on every call registers as a distinct entry
  // each time, which manifests as duplicate layout options and a property
  // pane that loses configuration pages partway through. (Root cause of
  // the very first bug hit in this series — not repeating it.)
  public static readonly serviceKey: ServiceKey<NewsDigestLayout> =
    ServiceKey.create<NewsDigestLayout>('NewsDigestLayout', NewsDigestLayout);

  public getPropertyPaneFieldsConfiguration(): IPropertyPaneField<unknown>[] {
    return [];
  }
}

export class PnpNewsDigestLibrary implements IExtensibilityLibrary {
  public static readonly serviceKey: ServiceKey<PnpNewsDigestLibrary> =
    ServiceKey.create<PnpNewsDigestLibrary>('SPFx:PnpNewsDigestLibrary', PnpNewsDigestLibrary);

  constructor(serviceScope: ServiceScope) {
    // Constructor intentionally left minimal
  }

  getCustomLayouts(): ILayoutDefinition[] {
    return [
      {
        name: 'News Digest',
        iconName: 'Mail',
        key: 'NewsDigestLayout',
        type: LayoutType.Results,
        renderType: LayoutRenderType.Handlebars,
        // eslint-disable-next-line @typescript-eslint/no-var-requires
        templateContent: require('./news-digest-layout.html').default.toString(),
        serviceKey: NewsDigestLayout.serviceKey
      }
    ];
  }

  public getCustomWebComponents(): IComponentDefinition<unknown>[] {
    return [
      {
        componentName: 'digest-checkbox-component',
        componentClass: DigestCheckboxWebComponent
      },
      {
        componentName: 'digest-compose-trigger-component',
        componentClass: DigestComposeTriggerWebComponent
      }
    ];
  }

  getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [];
  }

  registerHandlebarsCustomizations?(handlebarsNamespace: typeof Handlebars): void {
    // Reused verbatim from pnp-featured-news-layout and
    // pnp-multilingual-search-layouts — builds the encoded filter-state
    // query string PnP Search Filters expects for a taxonomy-backed
    // RefinableString, so the Department/OrgClass pills can act as
    // clickable filter links, same as on the other two layouts.
    handlebarsNamespace.registerHelper('taxonomyFilterQuery', (
      rawRefinableString: string,
      label: string,
      filterName: string
    ) => {
      if (!rawRefinableString) return '';

      const parts = rawRefinableString.split(';');
      const gp0Part = parts.find((p) => p.startsWith('GP0|#'));
      const l0PartRaw = parts.find((p) => p.startsWith('L0|#'));
      if (!gp0Part || !l0PartRaw) return '';

      const l0Part = l0PartRaw.split('|').slice(0, 2).join('|');

      const filterState = [{
        filterName,
        values: [{
          name: label,
          value: `or(${gp0Part},${l0Part})`,
          operator: 0
        }],
        operator: 'or',
        hideNodesNotInDataSet: true,
        expandAllNodesByDefault: false
      }];

      return encodeURIComponent(JSON.stringify(filterState));
    });

    // Formats a raw ISO date string (as LastModifiedTime comes back from
    // search) into a short, readable date for display on cards. Written
    // as our own helper rather than assuming a built-in date helper
    // exists in this Handlebars runtime — safer given earlier surprises
    // with assumed-available helpers (e.g. gt) turning out not to exist.
    handlebarsNamespace.registerHelper('formatDate', (isoString: string) => {
      if (!isoString) return '';
      const date = new Date(isoString);
      if (isNaN(date.getTime())) return isoString;
      return date.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
    });
  }

  invokeCardAction(action: IAdaptiveCardAction): void {
  }

  getCustomQueryModifiers?(): IQueryModifierDefinition[] {
    return [];
  }

  getCustomDataSources?(): IDataSourceDefinition[] {
    return [];
  }

  public name(): string {
    return 'PnpNewsDigestLibrary';
  }
}