import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { selectItem, deselectItem, isSelected, IDigestItem } from './digestSelectionStore';

export interface IDigestCheckboxComponentProps {
    itemId: string;
    title: string;
    link: string;
    summary: string;
    thumbnailUrl?: string;
}

export interface IDigestCheckboxComponentState {
    checked: boolean;
}

export class DigestCheckboxComponent extends React.Component<IDigestCheckboxComponentProps, IDigestCheckboxComponentState> {

    public constructor(props: IDigestCheckboxComponentProps) {
        super(props);
        // Reflects the store's current state on mount, not just false —
        // relevant if this exact item is already selected (e.g. the same
        // article legitimately reappears after a filter change or a
        // re-render), so the checkbox doesn't silently desync from the
        // store it's supposed to represent.
        this.state = {
            checked: isSelected(this.props.itemId)
        };
    }

    private handleChange = (_event?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
        const isNowChecked = !!checked;
        this.setState({ checked: isNowChecked });

        if (isNowChecked) {
            const item: IDigestItem = {
                id: this.props.itemId,
                title: this.props.title,
                link: this.props.link,
                summary: this.props.summary,
                thumbnailUrl: this.props.thumbnailUrl
            };
            selectItem(item);
        } else {
            // Deselecting on the card removes it from the digest
            // immediately — no separate "remove" step inside the compose
            // dialog, by design.
            deselectItem(this.props.itemId);
        }
    }

    public render(): React.ReactElement<IDigestCheckboxComponentProps> {
        return (
            <Checkbox
                checked={this.state.checked}
                onChange={this.handleChange}
                ariaLabel={`Add "${this.props.title}" to digest`}
            />
        );
    }
}

export class DigestCheckboxWebComponent extends BaseWebComponent {

    public async connectedCallback(): Promise<void> {
        const props = this.resolveAttributes();
        const component = <DigestCheckboxComponent
            itemId={props.itemId as string}
            title={props.title as string}
            link={props.link as string}
            summary={props.summary as string}
            thumbnailUrl={props.thumbnailUrl as string | undefined}
        />;
        ReactDOM.render(component, this);
    }

    protected onDispose(): void {
        ReactDOM.unmountComponentAtNode(this);
    }
}