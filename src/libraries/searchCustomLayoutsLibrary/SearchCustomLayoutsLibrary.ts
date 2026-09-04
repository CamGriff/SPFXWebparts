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
import { ContactOwnerWebComponent } from "../ContactOwnerComponent";
import { loadTranslationDrift, getDriftRecord } from './translationDriftStore';

export interface ISearchCardsLayoutProperties {
  showTranslationDriftFlag: boolean;
  showContactOwnerButton: boolean;
}

export class SearchCardsLayout extends BaseLayout<ISearchCardsLayoutProperties> {

  // Hoisted to a static field so the same ServiceKey object is reused on
  // every call, rather than a new key object being minted each time
  // getCustomLayouts() runs (this was the root cause of the duplicate
  // layout tile and the property pane losing pages 2/3).
  public static readonly serviceKey: ServiceKey<SearchCardsLayout> =
    ServiceKey.create<SearchCardsLayout>('SearchCardsLayout', SearchCardsLayout);

  // Bulk-fetches the TranslationDrift list once, before PnP renders the
  // template, so per-card Handlebars helpers can look values up
  // synchronously instead of each needing to await a call themselves.
  public async onInit(): Promise<void> {
    console.log('[SearchCustomLayoutsLibrary][DEBUG] onInit start, this.serviceScope:', this.serviceScope);
    await loadTranslationDrift(this.serviceScope, 'TranslationDrift');
  }

  public getPropertyPaneFieldsConfiguration(): IPropertyPaneField<unknown>[] {
    return [
      PropertyPaneCheckbox('layoutProperties.showTranslationDriftFlag', {
        text: 'Show translation drift flag',
        checked: true
      }),
      PropertyPaneCheckbox('layoutProperties.showContactOwnerButton', {
        text: 'Show contact owner action for outdated content',
        checked: true
      })
    ];
  }
}

export class SearchCustomLayoutsLibrary implements IExtensibilityLibrary {
  public static readonly serviceKey: ServiceKey<SearchCustomLayoutsLibrary> =
    ServiceKey.create<SearchCustomLayoutsLibrary>('SPFx:SearchCustomLayoutsLibrary', SearchCustomLayoutsLibrary);

  constructor(serviceScope: ServiceScope) {
    // Constructor intentionally left minimal
  }

  getCustomLayouts(): ILayoutDefinition[] {
    return [
      {
        name: 'Multilingual Cards',
        iconName: 'DocumentManagement',
        key: 'MultilingualCardsLayout',
        type: LayoutType.Results,
        renderType: LayoutRenderType.Handlebars,
        // eslint-disable-next-line @typescript-eslint/no-var-requires
        templateContent: require('./multilingual-cards-layout.html').default.toString(),
        serviceKey: SearchCardsLayout.serviceKey
      }
    ];
  }

  public getCustomWebComponents(): IComponentDefinition<unknown>[] {
    return [
      {
        componentName: 'contact-owner-component',
        componentClass: ContactOwnerWebComponent
      }
    ];
  }

  getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [];
  }

  registerHandlebarsCustomizations?(handlebarsNamespace: typeof Handlebars): void {
    // Block helper: {{#ifDrift item.UniqueID}}...{{else}}...{{/ifDrift}}
    // Avoids Handlebars subexpressions ({{#if (hasDrift ...)}}), which may
    // not be supported by PnP Modern Search's bundled Handlebars runtime.
    handlebarsNamespace.registerHelper('ifDrift', function (
      this: unknown,
      pageGuid: string,
      options: Handlebars.HelperOptions
    ) {
      const record = getDriftRecord(pageGuid);
      if (record && record.driftStatus !== 'In Sync') {
        return options.fn(this);
      }
      return options.inverse(this);
    });

    handlebarsNamespace.registerHelper('driftStatusClass', (pageGuid: string) => {
      const record = getDriftRecord(pageGuid);
      if (!record) return '';
      return record.driftStatus.toLowerCase().replace(/\s+/g, '');
    });

    handlebarsNamespace.registerHelper('driftBadgeLabel', (pageGuid: string) => {
      const record = getDriftRecord(pageGuid);
      if (!record) return '';
      return record.daysDrift > 0
        ? `${record.driftStatus} \u00b7 ${record.daysDrift}d`
        : record.driftStatus;
    });

    handlebarsNamespace.registerHelper('driftTooltip', (pageGuid: string) => {
      const record = getDriftRecord(pageGuid);
      if (!record) return '';
      if (record.driftStatus === 'Missing') return 'No translated version exists yet';
      if (record.driftStatus === 'Orphaned') return 'Translated page has no matching source page';
      if (record.driftStatus === 'Abandoned') return 'Translation exists but was never approved/published';
      if (record.daysDrift > 0) return `Translation is ${record.daysDrift} days behind the source page`;
      return record.driftStatus;
    });

    // Builds the encoded filter-state query string PnP Search Filters
    // expects for a taxonomy-backed RefinableString, so a pill/tag can
    // act as a clickable filter link (per the official "Clickable column
    // filter link" scenario, adapted for hierarchical/taxonomy fields
    // rather than plain text — taxonomy filters use an "or(GP0|#..,L0|#..)"
    // value shape, not a hex-encoded plain string).
    //
    // rawRefinableString: the raw field value, e.g.
    //   "GP0|#f7aebcce-...;L0|#0f7aebcce-...|Data protection;GTSet|#..."
    // label: the display label to show as the selected filter's name
    //   (use the Auto-resolved, language-correct field here)
    // filterName: the managed property name the Filters web part is
    //   configured to filter on, e.g. "RefinableString111"
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

      // L0 parts look like "L0|#0<guid>|Label" — keep only "L0|#0<guid>",
      // dropping the trailing "|Label" segment.
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
    return 'SearchCustomLayoutsLibrary';
  }
}