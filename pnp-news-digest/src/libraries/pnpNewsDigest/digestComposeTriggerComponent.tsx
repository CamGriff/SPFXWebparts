import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import { ServiceScope } from '@microsoft/sp-core-library';
import { PageContext } from '@microsoft/sp-page-context';
import { MSGraphClientFactory, SPHttpClient } from '@microsoft/sp-http';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { TextField } from '@fluentui/react/lib/TextField';
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { getSelectedItems, getSelectedCount, clearSelection, onSelectionChange, IDigestItem } from './digestSelectionStore';

export interface IDigestComposeTriggerComponentProps {
    peoplePickerContext: IPeoplePickerContext;
    graphClientFactory: MSGraphClientFactory;
}

interface ISelectedPerson {
    email: string;
    displayName: string;
}

export interface IDigestComposeTriggerComponentState {
    selectedCount: number;
    showDialog: boolean;
    recipients: ISelectedPerson[];
    introMessage: string;
    isSending: boolean;
    error: string | undefined;
    sent: boolean;
}

export class DigestComposeTriggerComponent extends React.Component<IDigestComposeTriggerComponentProps, IDigestComposeTriggerComponentState> {
    private unsubscribe: () => void;

    public constructor(props: IDigestComposeTriggerComponentProps) {
        super(props);
        this.state = {
            selectedCount: getSelectedCount(),
            showDialog: false,
            recipients: [],
            introMessage: '',
            isSending: false,
            error: undefined,
            sent: false
        };
    }

    public componentDidMount(): void {
        // Keeps the floating button's live count in sync with checkbox
        // clicks happening elsewhere on the page, without needing the
        // whole card grid to re-render.
        this.unsubscribe = onSelectionChange((items) => {
            this.setState({ selectedCount: items.length });
        });
    }

    public componentWillUnmount(): void {
        if (this.unsubscribe) {
            this.unsubscribe();
        }
    }

    private handleOpenClick = (): void => {
        this.setState({ showDialog: true, error: undefined, sent: false });
    }

    private onPeopleChange = (items: any[]): void => {
        const recipients: ISelectedPerson[] = (items || []).map((person) => {
            const email = person.secondaryText || person.loginName || person.id;
            return {
                email,
                displayName: person.text || person.displayName || email
            };
        });
        this.setState({ recipients });
    }

    private onIntroChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        this.setState({ introMessage: newValue || '' });
    }

    private buildDigestHtml = (items: IDigestItem[], introMessage: string): string => {
        const introHtml = introMessage
            ? `<p>${escapeHtml(introMessage)}</p>`
            : '';

        const articlesHtml = items.map((item) => `
            <tr>
                <td style="padding: 12px; border-bottom: 1px solid #edebe9; vertical-align: top;">
                    ${item.thumbnailUrl
                        ? `<img src="${escapeHtml(item.thumbnailUrl)}" alt="" width="120" style="border-radius: 4px; display: block;" />`
                        : ''}
                </td>
                <td style="padding: 12px; border-bottom: 1px solid #edebe9; vertical-align: top;">
                    <a href="${escapeHtml(item.link)}" style="font-size: 16px; font-weight: 600; color: #0078d4; text-decoration: none;">
                        ${escapeHtml(item.title)}
                    </a>
                    <p style="margin: 6px 0 0; font-size: 13px; color: #605e5c;">
                        ${escapeHtml(item.summary || '')}
                    </p>
                </td>
            </tr>
        `).join('');

        return `
            <html>
                <body style="font-family: 'Segoe UI', Arial, sans-serif; color: #323130;">
                    ${introHtml}
                    <table style="border-collapse: collapse; width: 100%;">
                        ${articlesHtml}
                    </table>
                </body>
            </html>
        `;
    }

    private handleSend = async (): Promise<void> => {
        const { recipients, introMessage } = this.state;
        const items = getSelectedItems();

        if (recipients.length === 0) {
            this.setState({ error: 'Please choose at least one recipient.' });
            return;
        }
        if (items.length === 0) {
            this.setState({ error: 'No articles are currently selected.' });
            return;
        }

        this.setState({ isSending: true, error: undefined });

        try {
            const client = await this.props.graphClientFactory.getClient('3');

            await client.api('/me/sendMail').post({
                message: {
                    subject: `News digest: ${items.length} article${items.length === 1 ? '' : 's'}`,
                    body: {
                        contentType: 'HTML',
                        content: this.buildDigestHtml(items, introMessage)
                    },
                    toRecipients: recipients.map((r) => ({ emailAddress: { address: r.email } }))
                }
            });

            this.setState({ isSending: false, sent: true });
            // Clears the selection once the digest has actually been sent,
            // so the next round starts fresh rather than carrying over
            // whatever was previously picked.
            clearSelection();
        } catch (error) {
            console.error('Error sending digest:', error);
            this.setState({ error: 'Failed to send the digest.', isSending: false });
        }
    }

    private closeDialog = (): void => {
        if (!this.state.isSending) {
            this.setState({ showDialog: false, error: undefined, introMessage: '', recipients: [] });
        }
    }

    public render(): React.ReactElement<IDigestComposeTriggerComponentProps> {
        const { selectedCount, showDialog, isSending, error, sent, recipients } = this.state;
        const items = getSelectedItems();

        return (
            <div>
                {selectedCount > 0 && (
                    <PrimaryButton
                        onClick={this.handleOpenClick}
                        styles={{
                            root: { padding: '0 24px', height: 40, fontWeight: 600 },
                            label: { fontSize: 14 }
                        }}
                    >
                        Compose digest ({selectedCount})
                    </PrimaryButton>
                )}

                <Dialog
                    hidden={!showDialog}
                    onDismiss={this.closeDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Compose news digest',
                        subText: `Sending ${items.length} article${items.length === 1 ? '' : 's'}.`
                    }}
                    modalProps={{ isBlocking: isSending }}
                    minWidth={480}
                >
                    {error && (
                        <MessageBar messageBarType={MessageBarType.error}>
                            {error}
                        </MessageBar>
                    )}

                    {sent ? (
                        <MessageBar messageBarType={MessageBarType.success}>
                            Digest sent to {recipients.map((r) => r.displayName).join(', ')}.
                        </MessageBar>
                    ) : (
                        <>
                            <ul style={{ maxHeight: 160, overflowY: 'auto', paddingLeft: 20, margin: '0 0 12px' }}>
                                {items.map((item) => (
                                    <li key={item.id} style={{ fontSize: 13 }}>{item.title}</li>
                                ))}
                            </ul>
                            <PeoplePicker
                                context={this.props.peoplePickerContext}
                                titleText="Send to"
                                personSelectionLimit={10}
                                showtooltip={true}
                                required={true}
                                principalTypes={[PrincipalType.User]}
                                resolveDelay={300}
                                onChange={this.onPeopleChange}
                            />
                            <TextField
                                label="Add an intro note (optional)"
                                multiline
                                rows={2}
                                value={this.state.introMessage}
                                onChange={this.onIntroChange}
                            />
                        </>
                    )}

                    <DialogFooter>
                        {!sent && (
                            <PrimaryButton
                                onClick={this.handleSend}
                                text={isSending ? 'Sending...' : 'Send digest'}
                                disabled={isSending || recipients.length === 0 || items.length === 0}
                            />
                        )}
                        <DefaultButton
                            onClick={this.closeDialog}
                            text={sent ? 'Close' : 'Cancel'}
                            disabled={isSending}
                        />
                    </DialogFooter>
                    {isSending && <Spinner label="Sending..." size={SpinnerSize.large} />}
                </Dialog>
            </div>
        );
    }
}

function escapeHtml(value: string): string {
    return (value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

export class DigestComposeTriggerWebComponent extends BaseWebComponent {
    private _peoplePickerContext: IPeoplePickerContext | undefined;
    private _graphClientFactory: MSGraphClientFactory | undefined;

    public async connectedCallback(): Promise<void> {
        const serviceScope: ServiceScope = this._serviceScope;
        serviceScope.whenFinished(() => {
            const pageContext = serviceScope.consume(PageContext.serviceKey);
            const spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
            this._graphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);

            // Same manually-assembled minimal context as ShareArticleComponent
            // — confirmed working there, reused verbatim here.
            this._peoplePickerContext = {
                absoluteUrl: pageContext.web.absoluteUrl,
                msGraphClientFactory: this._graphClientFactory,
                spHttpClient: spHttpClient
            } as unknown as IPeoplePickerContext;

            const component = <DigestComposeTriggerComponent
                peoplePickerContext={this._peoplePickerContext as IPeoplePickerContext}
                graphClientFactory={this._graphClientFactory as MSGraphClientFactory}
            />;
            ReactDOM.render(component, this);
        });
    }

    protected onDispose(): void {
        (ReactDOM.unmountComponentAtNode as unknown as (container: Element) => boolean)(this);
    }
}