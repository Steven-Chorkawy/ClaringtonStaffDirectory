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
                    key: 'excelExport',
                    text: 'Export to Excel',
                    title: 'Download Staff list as excel document.',
                    iconProps: { iconName: 'ExcelLogo' },
                },
                {
                    key: 'reloadStaffList',
                    text: 'Refresh Staff List',
                    title: 'Get most up-to-date list of staff members.',
                    iconProps: { iconName: 'Refresh' },
                },
            ],
            // By default, the menu will be focused when it opens. Uncomment the next line to prevent this.
            // shouldFocusOnMount: false
        };

        const moreOptionsButtonProps: IIconProps = { iconName: 'Add' };
        
        return <CommandButton title={'More Options'} iconProps={moreOptionsButtonProps} menuProps={menuProps} ariaLabel={'More Options'} />
    }
} 