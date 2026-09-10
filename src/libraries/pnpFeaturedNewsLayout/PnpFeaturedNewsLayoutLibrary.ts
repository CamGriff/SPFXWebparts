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
import { PropertyPaneCheckbox, IPropertyPaneField } from "@microsoft/sp-property-pane";
import { ReportContentWebComponent } from "./ReportContentComponent";
import { ShareArticleWebComponent } from "./ShareArticleComponent";

// Featured News layout: renders one "hero" item (the most recently
// published item flagged FeaturedNews, falling back to the single most
// recent item overall if none are flagged) alongside a small number of
// supporting tiles for the rest. Hero selection happens entirely in the
// Handlebars template via a custom block helper — the underlying search
// query just needs to return items sorted newest-first; no extra data
// source or REST call is needed for this part, since FeaturedNews only
// needs to be a *retrievable* field (Selected Properties), not a refiner.
export interface IFeaturedNewsLayoutProperties {
  // Reserved for future toggles (e.g. showReportButton, showShareButton)
  // if these actions should become independently switchable per web part
  // instance, mirroring the pattern used in the multilingual search layout.
}

export class FeaturedNewsLayout extends BaseLayout<IFeaturedNewsLayoutProperties> {

  // Static field, not created inline inside getCustomLayouts() — a fresh
  // ServiceKey object minted on every call registers as a distinct entry
  // each time, which manifests as duplicate layout options and a property
  // pane that loses configuration pages partway through.
  public static readonly serviceKey: ServiceKey<FeaturedNewsLayout> =
    ServiceKey.create<FeaturedNewsLayout>('FeaturedNewsLayout', FeaturedNewsLayout);

  public getPropertyPaneFieldsConfiguration(): IPropertyPaneField<unknown>[] {
    return [];
  }
}

export class PnpFeaturedNewsLayoutLibrary implements IExtensibilityLibrary {
  public static readonly serviceKey: ServiceKey<PnpFeaturedNewsLayoutLibrary> =
    ServiceKey.create<PnpFeaturedNewsLayoutLibrary>('SPFx:PnpFeaturedNewsLayoutLibrary', PnpFeaturedNewsLayoutLibrary);

  constructor(serviceScope: ServiceScope) {
    // Constructor intentionally left minimal
  }

  getCustomLayouts(): ILayoutDefinition[] {
    return [
      {
        name: 'Featured News',
        iconName: 'NewsSearch',
        key: 'FeaturedNewsLayout',
        type: LayoutType.Results,
        renderType: LayoutRenderType.Handlebars,
        // eslint-disable-next-line @typescript-eslint/no-var-requires
        templateContent: require('./featured-news-layout.html').default.toString(),
        serviceKey: FeaturedNewsLayout.serviceKey
      }
    ];
  }

  public getCustomWebComponents(): IComponentDefinition<unknown>[] {
    return [
      {
        componentName: 'report-content-component',
        componentClass: ReportContentWebComponent
      },
      {
        componentName: 'share-article-component',
        componentClass: ShareArticleWebComponent
      }
    ];
  }

  getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [];
  }

  registerHandlebarsCustomizations?(handlebarsNamespace: typeof Handlebars): void {
    // Block helper: {{#withFeaturedHero data.items}}...{{/withFeaturedHero}}
    // Walks the (already newest-first sorted) result set, returns the
    // first item where FeaturedNews is truthy as the hero. If none are
    // flagged, falls back to the very first item in the set (the single
    // most recent post overall), so the hero slot is never empty.
    //
    // NOTE: the exact truthy check for item.FeaturedNews (e.g. "1" vs
    // "true" vs boolean true) needs confirming once the field is actually
    // retrievable — Yes/No site columns can come back from search as
    // different literal values depending on tenant/version. Treat this
    // as a placeholder to verify, not a confirmed-correct check.
    handlebarsNamespace.registerHelper('withFeaturedHero', function (
      this: unknown,
      items: any[],
      options: Handlebars.HelperOptions
    ) {
      if (!items || items.length === 0) {
        return '';
      }
      const hero = items.find((item) => isFeatured(item)) || items[0];
      // Pass blockParams so the template can bind an explicit named
      // parameter (as |item|), matching the pattern used in
      // pnp-multilingual-search-layouts, rather than relying on implicit
      // `this` context switching inside the block.
      return options.fn(hero, { blockParams: [hero] });
    });

    // Block helper: {{#eachNonHero data.items 3}}...{{/eachNonHero}}
    // Renders up to `limit` items, excluding whichever one withFeaturedHero
    // selected — re-running the same selection logic so the two helpers
    // never disagree about which item is the hero. limit defaults to 3
    // (matching the 3-tiles-down-the-side design) if omitted.
    handlebarsNamespace.registerHelper('eachNonHero', function (
      this: unknown,
      items: any[],
      limitOrOptions: number | Handlebars.HelperOptions,
      maybeOptions?: Handlebars.HelperOptions
    ) {
      const options = maybeOptions || (limitOrOptions as Handlebars.HelperOptions);
      const limit = maybeOptions ? (limitOrOptions as number) : 3;

      if (!items || items.length === 0) {
        return '';
      }
      const hero = items.find((item) => isFeatured(item)) || items[0];
      let result = '';
      let count = 0;
      for (const item of items) {
        if (item === hero) {
          continue;
        }
        if (count >= limit) {
          break;
        }
        result += options.fn(item, { blockParams: [item] });
        count++;
      }
      return result;
    });

    // Builds the encoded filter-state query string PnP Search Filters
    // expects for a taxonomy-backed RefinableString, so a pill can act as
    // a clickable filter link. Copied verbatim from the working
    // implementation in pnp-multilingual-search-layouts — same mechanism,
    // reused across both the Directorate (RefinableString113) and
    // OrgClassification (RefinableString111) pills on this card.
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
    return 'PnpFeaturedNewsLayoutLibrary';
  }
}

// Placeholder truthiness check for the FeaturedNews Yes/No column's
// search-returned value — needs confirming against real data once the
// field is retrievable (see selectedProperties check below).
function isFeatured(item: any): boolean {
  // Confirmed against real search results: the Yes/No column FeaturedNews
  // is retrievable as FeaturedNewsOWSBOOL, and comes back as the string
  // "1" when checked, or null when unchecked.
  const value = item?.FeaturedNewsOWSBOOL;
  return value === '1' || value === true;
}