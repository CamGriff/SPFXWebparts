import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import { ServiceScope } from '@microsoft/sp-core-library';
import { PageContext } from '@microsoft/sp-page-context';
import { MSGraphClientFactory, SPHttpClient, AadHttpClientFactory } from '@microsoft/sp-http';
import { IconButton, PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { TextField } from '@fluentui/react/lib/TextField';
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from '@pnp/spfx-controls-react/lib/PeoplePicker';

export interface IShareArticleComponentProps {
    // PeoplePicker from @pnp/spfx-controls-react expects a full
    // WebPartContext-shaped object (SPHttpClient, MSGraphClientFactory,
    // pageContext, etc.) — not just PageContext. Since a web component only
    // has access to individual services consumed from serviceScope, this is
    // assembled manually in ShareArticleWebComponent.connectedCallback()
    // below. This assembly is the one part of this component genuinely
    // unverified against a real build — if PeoplePicker throws on missing
    // properties, this object is the first place to check.
    peoplePickerContext: IPeoplePickerContext;
    graphClientFactory: MSGraphClientFactory;
    pageUrl?: string;
    title?: string;
}

interface ISelectedPerson {
    email: string;
    displayName: string;
}

export interface IShareArticleComponentState {
    showDialog: boolean;
    selectedPerson: ISelectedPerson | undefined;
    customMessage: string;
    isSending: boolean;
    error: string | undefined;
    sent: boolean;
}

export class ShareArticleComponent extends React.Component<IShareArticleComponentProps, IShareArticleComponentState> {

    public constructor(props: IShareArticleComponentProps) {
        super(props);
        this.state = {
            showDialog: false,
            selectedPerson: undefined,
            customMessage: '',
            isSending: false,
            error: undefined,
            sent: false
        };
    }

    private handleOpenClick = (): void => {
        this.setState({ showDialog: true, error: undefined, sent: false });
    }

    private onPeopleChange = (items: any[]): void => {
        if (!items || items.length === 0) {
            this.setState({ selectedPerson: undefined });
            return;
        }
        const person = items[0];
        // PeoplePicker resolves email/login differently depending on the
        // source (Search vs SharePoint) — secondaryText is the most
        // reliable field for an email-formatted UPN in most tenants, with
        // a couple of fallbacks. Worth confirming against a real picked
        // person during testing rather than assuming this is exhaustive.
        const email = person.secondaryText || person.loginName || person.id;
        this.setState({
            selectedPerson: {
                email,
                displayName: person.text || person.displayName || email
            }
        });
    }

    private handleSend = async (): Promise<void> => {
        const { selectedPerson, customMessage } = this.state;
        const { title, pageUrl } = this.props;

        if (!selectedPerson?.email) {
            this.setState({ error: 'Please choose someone to share this with.' });
            return;
        }

        this.setState({ isSending: true, error: undefined });

        try {
            const client = await this.props.graphClientFactory.getClient('3');

            // This only emails a link — it does not grant the recipient any
            // access to the page. If they don't already have permission to
            // view it, the link in the email simply won't work for them.
            const messageLines = [
                `Hi ${selectedPerson.displayName},`,
                '',
                `Thought you might find this article worth a read: "${title}".`
            ];
            if (customMessage) {
                messageLines.push('', customMessage);
            }
            messageLines.push('', `Article: ${pageUrl}`);

            await client.api('/me/sendMail').post({
                message: {
                    subject: `Sharing: ${title}`,
                    body: {
                        contentType: 'Text',
                        content: messageLines.join('\n')
                    },
                    toRecipients: [
                        { emailAddress: { address: selectedPerson.email } }
                    ]
                }
            });

            this.setState({ isSending: false, sent: true });
        } catch (error) {
            console.error('Error sending share email:', error);
            this.setState({ error: 'Failed to send the email.', isSending: false });
        }
    }

    private closeDialog = (): void => {
        if (!this.state.isSending) {
            this.setState({ showDialog: false, error: undefined, customMessage: '', selectedPerson: undefined });
        }
    }

    private onMessageChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        this.setState({ customMessage: newValue || '' });
    }

    public render(): React.ReactElement<IShareArticleComponentProps> {
        const { showDialog, isSending, selectedPerson, error, sent } = this.state;
        const { title, peoplePickerContext } = this.props;

        return (
            <div>
                <IconButton
                    iconProps={{ iconName: 'Share' }}
                    title="Share this article"
                    ariaLabel="Share this article by email"
                    onClick={this.handleOpenClick}
                />

                <Dialog
                    hidden={!showDialog}
                    onDismiss={this.closeDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Share article',
                        subText: `Email a link to "${title}" to a colleague.`
                    }}
                    modalProps={{ isBlocking: isSending }}
                    minWidth={420}
                >
                    {error && (
                        <MessageBar messageBarType={MessageBarType.error}>
                            {error}
                        </MessageBar>
                    )}

                    {sent ? (
                        <MessageBar messageBarType={MessageBarType.success}>
                            Sent to {selectedPerson?.displayName}.
                        </MessageBar>
                    ) : (
                        <>
                            <PeoplePicker
                                context={peoplePickerContext}
                                titleText="Share with"
                                personSelectionLimit={1}
                                showtooltip={true}
                                required={true}
                                principalTypes={[PrincipalType.User]}
                                resolveDelay={300}
                                onChange={this.onPeopleChange}
                            />
                            <TextField
                                label="Add a note (optional)"
                                multiline
                                rows={3}
                                value={this.state.customMessage}
                                onChange={this.onMessageChange}
                            />
                            <p style={{ fontSize: 12, color: '#605e5c' }}>
                                This only sends a link by email — it does not change who can access the page.
                            </p>
                        </>
                    )}

                    <DialogFooter>
                        {!sent && (
                            <PrimaryButton
                                onClick={this.handleSend}
                                text={isSending ? 'Sending...' : 'Send'}
                                disabled={isSending || !selectedPerson}
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

export class ShareArticleWebComponent extends BaseWebComponent {
    private _peoplePickerContext: IPeoplePickerContext | undefined;
    private _graphClientFactory: MSGraphClientFactory | undefined;

    public async connectedCallback(): Promise<void> {
        const props = this.resolveAttributes();
        const serviceScope: ServiceScope = this._serviceScope;
        serviceScope.whenFinished(() => {
            const pageContext = serviceScope.consume(PageContext.serviceKey);
            const spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
            this._graphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);

            // Manually assembled minimal context for PeoplePicker, since a
            // web component doesn't have a real WebPartContext to hand it —
            // only whatever individual services get consumed from
            // serviceScope here. This is the one piece of this component
            // not yet confirmed against a real build; if PeoplePicker fails
            // to resolve people, this shape is the first thing to check
            // against @pnp/spfx-controls-react's actual IPeoplePickerContext
            // type for the installed version.
            this._peoplePickerContext = {
                absoluteUrl: pageContext.web.absoluteUrl,
                msGraphClientFactory: this._graphClientFactory,
                spHttpClient: spHttpClient
            } as unknown as IPeoplePickerContext;

            const component = <ShareArticleComponent
                peoplePickerContext={this._peoplePickerContext as IPeoplePickerContext}
                graphClientFactory={this._graphClientFactory as MSGraphClientFactory}
                {...props} />;
            ReactDOM.render(component, this);
        });
    }

    protected onDispose(): void {
        ReactDOM.unmountComponentAtNode(this);
    }
}