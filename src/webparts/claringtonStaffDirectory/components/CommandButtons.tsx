import * as React from 'react';
import { CommandButton, IContextualMenuProps, IIconProps } from '@fluentui/react';


export default class CommandButtons extends React.Component<any, any> {

    constructor(props) {
        super(props);
    }

    public render(): React.ReactElement<any> {
        const menuProps: IContextualMenuProps = {
            items: [
                {
                    key: 'emailMessage',
                    text: 'Email message',
                    iconProps: { iconName: 'Mail' },
                },
                {
                    key: 'calendarEvent',
                    text: 'Calendar event',
                    iconProps: { iconName: 'Calendar' },
                },
            ],
            // By default, the menu will be focused when it opens. Uncomment the next line to prevent this.
            // shouldFocusOnMount: false
        };

        const verticalMenuIcon: IIconProps = { iconName: 'MoreVertical' };
        
        return <CommandButton iconProps={verticalMenuIcon} text="Options" menuProps={menuProps} />
    }
} 