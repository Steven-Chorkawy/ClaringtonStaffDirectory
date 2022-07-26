import * as React from 'react';
import { CommandButton, IContextualMenuItem, IContextualMenuProps, IIconProps } from '@fluentui/react';

export interface ICommandButtonsProps {
    menuItems: IContextualMenuItem[];
}

export default class CommandButtons extends React.Component<ICommandButtonsProps, any> {

    constructor(props) {
        super(props);
    }

    public render(): React.ReactElement<any> {
        const menuProps: IContextualMenuProps = {
            items: this.props.menuItems,
            // By default, the menu will be focused when it opens. Uncomment the next line to prevent this.
            // shouldFocusOnMount: false
        };

        const moreOptionsButtonProps: IIconProps = { iconName: 'Add' };

        return <CommandButton title={'More Options'} iconProps={moreOptionsButtonProps} menuProps={menuProps} ariaLabel={'More Options'} />
    }
} 