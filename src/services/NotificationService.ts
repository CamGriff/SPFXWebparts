import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import { IDriftItem } from '../models/IDriftItem';

const TRANSLATION_DRIFT_LIST_TITLE = 'TranslationDrift';
const MAX_ITEMS_PER_REQUEST = 5000;
const SEND_MAIL_ENDPOINT = '/me/sendMail';

interface IExistingDriftRow {
  Id: number;
  PageGuid: string;
}

export async function sendNudgeEmail(
  graphClient: MSGraphClientV3,
  translatorEmail: string,
  translatorName: string,
  stalePages: IDriftItem[],
  siteUrl: string
): Promise<void> {
  if (stalePages.length === 0) {
    return;
  }

  const message = {
    subject: `Translation pages need attention — ${siteUrl}`,
    body: {
      contentType: 'HTML',
      content: buildEmailBody(translatorName, stalePages, siteUrl)
    },
    toRecipients: [
      {
        emailAddress: {
          address: translatorEmail,
          name: translatorName
        }
      }
    ]
  };

  await graphClient.api(SEND_MAIL_ENDPOINT).post({ message, saveToSentItems: 'false' });
}

export async function markNudgeSent(sp: SPFI, pageGuids: string[], language: string): Promise<void> {
  if (pageGuids.length === 0) {
    return;
  }

  const pageGuidSet = new Set(pageGuids.map((guid) => guid.toLowerCase()));
  const list = sp.web.lists.getByTitle(TRANSLATION_DRIFT_LIST_TITLE);

  const rows: IExistingDriftRow[] = await list.items
    .select('Id', 'PageGuid')
    .filter(`TranslationLanguage eq '${escapeODataString(language)}'`)
    .top(MAX_ITEMS_PER_REQUEST)();

  const matchingRows = rows.filter((row) => pageGuidSet.has(row.PageGuid.toLowerCase()));
  const nudgeDate = new Date().toISOString();

  for (const row of matchingRows) {
    await list.items.getById(row.Id).update({ NudgeSent: true, NudgeDate: nudgeDate });
  }
}

function buildEmailBody(translatorName: string, stalePages: IDriftItem[], siteUrl: string): string {
  const rows = stalePages.map((page) => buildRowHtml(page)).join('');

  return `
    <html>
      <body style="font-family: 'Segoe UI', Arial, sans-serif; color: #323130;">
        <p>Hi ${escapeHtml(translatorName)},</p>
        <p>The following pages on <strong>${escapeHtml(siteUrl)}</strong> need a translation update:</p>
        <table style="border-collapse: collapse; width: 100%; font-size: 14px;">
          <thead>
            <tr style="background-color: #f3f2f1; text-align: left;">
              <th style="padding: 8px; border: 1px solid #e1dfdd;">Page</th>
              <th style="padding: 8px; border: 1px solid #e1dfdd;">Status</th>
              <th style="padding: 8px; border: 1px solid #e1dfdd;">Days Drift</th>
              <th style="padding: 8px; border: 1px solid #e1dfdd;">English Page</th>
              <th style="padding: 8px; border: 1px solid #e1dfdd;">French Page</th>
            </tr>
          </thead>
          <tbody>
            ${rows}
          </tbody>
        </table>
        <p>Please review and update the translations at your earliest convenience.</p>
      </body>
    </html>
  `;
}

function buildRowHtml(page: IDriftItem): string {
  return `
    <tr>
      <td style="padding: 8px; border: 1px solid #e1dfdd;">${escapeHtml(page.DefaultPageTitle)}</td>
      <td style="padding: 8px; border: 1px solid #e1dfdd;">${escapeHtml(page.DriftStatus)}</td>
      <td style="padding: 8px; border: 1px solid #e1dfdd;">${page.DaysDrift}</td>
      <td style="padding: 8px; border: 1px solid #e1dfdd;">${buildLinkCell(page.DefaultPageUrl)}</td>
      <td style="padding: 8px; border: 1px solid #e1dfdd;">${buildLinkCell(page.TranslationPageUrl)}</td>
    </tr>
  `;
}

function buildLinkCell(url: string): string {
  if (!url) {
    return '&mdash;';
  }
  const escapedUrl = escapeHtml(url);
  return `<a href="${escapedUrl}">${escapedUrl}</a>`;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}
