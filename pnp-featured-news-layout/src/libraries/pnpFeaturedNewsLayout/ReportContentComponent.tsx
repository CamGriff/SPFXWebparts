import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import { ServiceScope } from '@microsoft/sp-core-library';
import { PageContext } from '@microsoft/sp-page-context';
import { MSGraphClientFactory } from '@microsoft/sp-http';
import { SPFx, spfi } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import { Web } from "@pnp/sp/webs";
import { IconButton, PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { TextField } from '@fluentui/react/lib/TextField';

export interface IReportContentComponentProps {
    context: PageContext;
    graphClientFactory: MSGraphClientFactory;
    itemId?: string;
    pageUrl?: string;
    siteUrl?: string;
    webUrl?: string;
    title?: string;
}

export interface IReportContentComponentState {
    showDialog: boolean;
    isLoadingOwner: boolean;
    isSending: boolean;
    ownerEmail: string | undefined;
    ownerName: string | undefined;
    customMessage: string;
    error: string | undefined;
    sent: boolean;
}

export class ReportContentComponent extends React.Component<IReportContentComponentProps, IReportContentComponentState> {
    private sp: ReturnType<typeof spfi>;

    public constructor(props: IReportContentComponentProps) {
        super(props);
        this.state = {
            showDialog: false,
            isLoadingOwner: false,
            isSending: false,
            ownerEmail: undefined,
            ownerName: undefined,
            customMessage: '',
            error: undefined,
            sent: false
        };
        // This SPFx() call — given a real PageContext under { pageContext } —
        // is the proven-working pattern from ContactOwnerComponent. Do not
        // change this to a bare serviceScope; that variant never resolved
        // correctly when tried elsewhere in this codebase.
        this.sp = spfi().using(SPFx({ pageContext: this.props.context }));
    }

    private handleOpenClick = async (): Promise<void> => {
        this.setState({ showDialog: true, isLoadingOwner: true, error: undefined, sent: false });
        try {
            await this.loadOwner();
        } catch (error) {
            console.error('Error loading content owner:', error);
            this.setState({ error: 'Could not determine the content owner.', isLoadingOwner: false });
        }
    }

    private loadOwner = async (): Promise<void> => {
        const { webUrl, itemId } = this.props;
        if (!webUrl || !itemId) {
            throw new Error('Missing required parameters');
        }

        const web = Web([this.sp.web, webUrl]);
        const list = web.lists.getByTitle("Site Pages");

        // ContentOwner is a real Person/Group field on Site Pages, maintained
        // by editors — same lookup already confirmed working end-to-end on
        // the multilingual search cards.
        const item = await list.items.getById(parseInt(itemId, 10))
            .select("ContentOwner/Title", "ContentOwner/EMail")
            .expand("ContentOwner")();

        this.setState({
            ownerName: item.ContentOwner?.Title || undefined,
            ownerEmail: item.ContentOwner?.EMail || undefined,
            isLoadingOwner: false
        });
    }

    private handleSend = async (): Promise<void> => {
        const { ownerEmail, ownerName, customMessage } = this.state;
        const { title, pageUrl } = this.props;

        if (!ownerEmail) {
            this.setState({ error: 'No email address found for the content owner.' });
            return;
        }

        this.setState({ isSending: true, error: undefined });

        try {
            const client = await this.props.graphClientFactory.getClient('3');

            const messageLines = [
                `Hi ${ownerName || ''},`,
                '',
                `A reader has flagged "${title}" as possibly needing review.`
            ];
            if (customMessage) {
                messageLines.push('', customMessage);
            }
            messageLines.push('', `Article: ${pageUrl}`);

            await client.api('/me/sendMail').post({
                message: {
                    subject: `Content flagged for review: ${title}`,
                    body: {
                        contentType: 'Text',
                        content: messageLines.join('\n')
                    },
                    toRecipients: [
                        { emailAddress: { address: ownerEmail } }
                    ]
                }
            });

            this.setState({ isSending: false, sent: true });
        } catch (error) {
            console.error('Error sending report notification:', error);
            this.setState({ error: 'Failed to send the notification email.', isSending: false });
        }
    }

    private closeDialog = (): void => {
        if (!this.state.isSending) {
            this.setState({ showDialog: false, error: undefined, customMessage: '' });
        }
    }

    private onMessageChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        this.setState({ customMessage: newValue || '' });
    }

    public render(): React.ReactElement<IReportContentComponentProps> {
        const { showDialog, isLoadingOwner, isSending, ownerEmail, ownerName, error, sent } = this.state;
        const { title } = this.props;

        return (
            <div>
                <IconButton
                    iconProps={{ iconName: 'Flag' }}
                    title="Report an issue with this article"
                    ariaLabel="Report an issue with this article"
                    onClick={this.handleOpenClick}
                />

                <Dialog
                    hidden={!showDialog}
                    onDismiss={this.closeDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Report content issue',
                        subText: `Let the content owner know something looks wrong with "${title}".`
                    }}
                    modalProps={{ isBlocking: isSending }}
                    minWidth={420}
                >
                    {error && (
                        <MessageBar messageBarType={MessageBarType.error}>
                            {error}
                        </MessageBar>
                    )}

                    {isLoadingOwner ? (
                        <Spinner label="Looking up content owner..." size={SpinnerSize.medium} />
                    ) : sent ? (
                        <MessageBar messageBarType={MessageBarType.success}>
                            Notification sent to {ownerName || ownerEmail}.
                        </MessageBar>
                    ) : ownerEmail ? (
                        <>
                            <p>This will email <strong>{ownerName || ownerEmail}</strong> ({ownerEmail}) about this article.</p>
                            <TextField
                                label="Add a note (optional)"
                                multiline
                                rows={3}
                                value={this.state.customMessage}
                                onChange={this.onMessageChange}
                            />
                        </>
                    ) : (
                        <MessageBar messageBarType={MessageBarType.warning}>
                            No content owner email could be found for this article.
                        </MessageBar>
                    )}

                    <DialogFooter>
                        {!sent && (
                            <PrimaryButton
                                onClick={this.handleSend}
                                text={isSending ? 'Sending...' : 'Send notification'}
                                disabled={isSending || isLoadingOwner || !ownerEmail}
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

export class ReportContentWebComponent extends BaseWebComponent {
    private _pageContext: PageContext | undefined;
    private _graphClientFactory: MSGraphClientFactory | undefined;

    public async connectedCallback(): Promise<void> {
        const props = this.resolveAttributes();
        const serviceScope: ServiceScope = this._serviceScope;
        serviceScope.whenFinished(() => {
            this._pageContext = serviceScope.consume(PageContext.serviceKey);
            this._graphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);
        });
        const component = <ReportContentComponent
            context={this._pageContext as PageContext}
            graphClientFactory={this._graphClientFactory as MSGraphClientFactory}
            {...props} />;
        ReactDOM.render(component, this);
    }

    protected onDispose(): void {
        ReactDOM.unmountComponentAtNode(this);
    }
}