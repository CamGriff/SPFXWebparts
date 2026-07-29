import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users";
import { ITool } from "./models/ITool";
import { IUserPreference } from "./models/IUserPreference";

const TOOLS_LIST = "OECDIntranet_Tools";
const PREFS_LIST = "OECDIntranet_ToolsFavorites";

export class ToolsService {

  private sp;
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
    this.sp = spfi().using(SPFx(context));
  }

  // ─── LANGUAGE DETECTION ────────────────────────────────────────

  public getLanguage(): string {
    const lcid = this.context.pageContext.cultureInfo.currentUICultureName;
    // Returns "fr-FR", "fr-BE" etc for French, otherwise default to EN
    return lcid.startsWith("fr") ? "FR" : "EN";
  }

  // ─── TOOLS LIST ────────────────────────────────────────────────

  public async getAllTools(): Promise<ITool[]> {
    const items = await this.sp.web.lists
      .getByTitle(TOOLS_LIST)
      .items
      .select(
        "ID",
        "TitleEN",
        "TitleFR",
        "ToolDescriptionEN",
        "ToolDescriptionFR",
        "ToolURLEN",
        "ToolURLFR",
        "ToolIcon"
      )();

    const lang = this.getLanguage();

    return items.map((item) => ({
      id: item.ID.toString(),
      titleEN: item.TitleEN,
      titleFR: item.TitleFR,
      descriptionEN: item.ToolDescriptionEN,
      descriptionFR: item.ToolDescriptionFR,
      urlEN: item.ToolURLEN,
      urlFR: item.ToolURLFR,
      toolIcon: item.ToolIcon,
      // Resolved fields — components just use tool.title, tool.url
      title: lang === "FR" ? item.TitleFR : item.TitleEN,
      description: lang === "FR" ? item.ToolDescriptionFR : item.ToolDescriptionEN,
      url: lang === "FR" ? item.ToolURLFR : item.ToolURLEN,
    }));
  }

  // ─── USER PREFERENCES ──────────────────────────────────────────

  public async getUserPreference(): Promise<IUserPreference | undefined> {
    const userId = this.context.pageContext.user.loginName;

    const items = await this.sp.web.lists
      .getByTitle(PREFS_LIST)
      .items
      .filter(`Title eq '${userId}'`)
      .select("ID", "Title", "SelectedTools", "LastUpdated")
      .top(1)();

    if (items.length === 0) return undefined;

    const item = items[0];
    return {
      id: item.ID.toString(),
      userId: item.Title,
      tools: item.SelectedTools ? JSON.parse(item.SelectedTools) : [],
      lastUpdated: item.LastUpdated,
    };
  }

  public async saveUserPreference(tools: string[]): Promise<void> {
    const userId = this.context.pageContext.user.loginName;
    const existing = await this.getUserPreference();

    const payload = {
      Title: userId,
      SelectedTools: JSON.stringify(tools),
      LastUpdated: new Date().toISOString(),
    };

    if (existing?.id) {
      // Update existing preference
      await this.sp.web.lists
        .getByTitle(PREFS_LIST)
        .items
        .getById(parseInt(existing.id))
        .update(payload);
    } else {
      // Create new preference
      await this.sp.web.lists
        .getByTitle(PREFS_LIST)
        .items
        .add(payload);
    }
  }
}