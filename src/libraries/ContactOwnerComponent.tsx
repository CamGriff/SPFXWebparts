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

export interface IContactOwnerComponentProps {
    context: PageContext;
    graphClientFactory: MSGraphClientFactory;
    itemId?: string;
    pageUrl?: string;
    siteUrl?: string;
    webUrl?: string;
    title?: string;
    driftStatus?: string;
}

export interface IContactOwnerComponentState {
    showDialog: boolean;
    isLoadingOwner: boolean;
    isSending: boolean;
    ownerEmail: string | undefined;
    ownerName: string | undefined;
    customMessage: string;
    error: string | undefined;
    sent: boolean;
}

export class ContactOwnerComponent extends React.Component<IContactOwnerComponentProps, IContactOwnerComponentState> {
    private sp: ReturnType<typeof spfi>;

    public constructor(props: IContactOwnerComponentProps) {
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
        this.sp = spfi().using(SPFx({ pageContext: this.props.context }));
    }

    private handleOpenClick = async (): Promise<void> => {
        this.setState({ showDialog: true, isLoadingOwner: true, error: undefined, sent: false });
        try {
            await this.loadOwner();
        } catch (error) {
            console.error('Error loading page owner:', error);
            this.setState({ error: 'Could not determine the page owner.', isLoadingOwner: false });
        }
    }

    private loadOwner = async (): Promise<void> => {
        const { webUrl, itemId } = this.props;
        if (!webUrl || !itemId) {
            throw new Error('Missing required parameters');
        }

        const web = Web([this.sp.web, webUrl]);
        const list = web.lists.getByTitle("Site Pages");

        // ContentOwner is a real Person/Group field on Site Pages,
        // maintained by editors — a better signal than Author (who created
        // the page) for who should be notified about outdated content.
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
        const { title, pageUrl, driftStatus } = this.props;

        if (!ownerEmail) {
            this.setState({ error: 'No email address found for the page owner.' });
            return;
        }

        this.setState({ isSending: true, error: undefined });

        try {
            const client = await this.props.graphClientFactory.getClient('3');

            const messageLines = [
                `Hi ${ownerName || ''},`,
                '',
                `This is a note about "${title}", which is currently flagged: ${driftStatus}.`
            ];
            if (customMessage) {
                messageLines.push('', customMessage);
            }
            messageLines.push('', `Page: ${pageUrl}`);

            await client.api('/me/sendMail').post({
                message: {
                    subject: `Content needs attention: ${title}`,
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
            console.error('Error sending owner notification:', error);
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

    public render(): React.ReactElement<IContactOwnerComponentProps> {
        const { showDialog, isLoadingOwner, isSending, ownerEmail, ownerName, error, sent } = this.state;
        const { title, driftStatus } = this.props;

        return (
            <div>
                <IconButton
                    iconProps={{ iconName: 'Mail' }}
                    title="Contact page owner"
                    ariaLabel="Contact page owner about outdated content"
                    onClick={this.handleOpenClick}
                />

                <Dialog
                    hidden={!showDialog}
                    onDismiss={this.closeDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Contact page owner',
                        subText: `"${title}" is flagged: ${driftStatus}`
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
                        <Spinner label="Looking up page owner..." size={SpinnerSize.medium} />
                    ) : sent ? (
                        <MessageBar messageBarType={MessageBarType.success}>
                            Notification sent to {ownerName || ownerEmail}.
                        </MessageBar>
                    ) : ownerEmail ? (
                        <>
                            <p>This will email <strong>{ownerName || ownerEmail}</strong> ({ownerEmail}) about this page.</p>
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
                            No owner email could be found for this page.
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

export class ContactOwnerWebComponent extends BaseWebComponent {
    private _pageContext: PageContext | undefined;
    private _graphClientFactory: MSGraphClientFactory | undefined;

    public async connectedCallback(): Promise<void> {
        const props = this.resolveAttributes();
        const serviceScope: ServiceScope = this._serviceScope;
        serviceScope.whenFinished(() => {
            this._pageContext = serviceScope.consume(PageContext.serviceKey);
            this._graphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);
        });
        const component = <ContactOwnerComponent context={this._pageContext as PageContext} graphClientFactory={this._graphClientFactory as MSGraphClientFactory} {...props} />;
        ReactDOM.render(component, this);
    }

    protected onDispose(): void {
        ReactDOM.unmountComponentAtNode(this);
    }
}