"use client";

import React from 'react';
import { Layout, Menu } from 'antd';
import type { MenuProps } from 'antd';
import {
    AppstoreOutlined,
    FileTextOutlined,
    HistoryOutlined,
    SettingOutlined
} from '@ant-design/icons';
import { usePathname, useRouter } from 'next/navigation';

const { Sider } = Layout;

type MenuItem = Required<MenuProps>['items'][number];

const menuItems: MenuItem[] = [
    {
        key: '/dashboard',
        icon: <AppstoreOutlined />,
        label: 'Dashboard',
        disabled: false,
    },
    {
        key: '/generate',
        icon: <FileTextOutlined />,
        label: 'Generate',
        disabled: false,
        children: [
            {
                key: '/generate/design-document',
                label: 'Design Document',
                disabled: false,
            },
            {
                key: '/generate/final-config',
                label: 'Final Configuration',
                disabled: false,
            },
        ],
    },
    {
        key: '/history',
        icon: <HistoryOutlined />,
        label: 'History',
        disabled: false,
    },
    {
        key: '/settings',
        icon: <SettingOutlined />,
        label: 'Settings',
        disabled: false,
    },
];

export default function Sidebar() {
    const pathname = usePathname();
    const router = useRouter();

    // Determine the selected key based on current pathname
    // For task workflow pages like /generate/design-document/tasks/xxx/setup,
    // we still highlight the /generate/design-document menu item
    const getSelectedKey = () => {
        if (pathname.startsWith('/generate/final-config')) return '/generate/final-config';
        if (pathname.startsWith('/generate/design-document')) return '/generate/design-document';
        if (pathname.startsWith('/history')) return '/history';
        if (pathname.startsWith('/settings')) return '/settings';
        if (pathname.startsWith('/dashboard')) return '/dashboard';
        return pathname;
    };

    // Determine which submenus should be open
    const getOpenKeys = () => {
        if (pathname.startsWith('/generate')) return ['/generate'];
        return [];
    };

    const handleMenuClick: MenuProps['onClick'] = ({ key }) => {
        router.push(key);
    };

    return (
        <Sider
            collapsible
            width={240}
            theme="light"
            style={{
                overflow: 'auto',
                height: '100%',
                borderRight: '1px solid #f0f0f0',
            }}
        >
            <Menu
                mode="inline"
                selectedKeys={[getSelectedKey()]}
                defaultOpenKeys={getOpenKeys()}
                style={{ height: '100%', borderRight: 0, paddingTop: '8px' }}
                items={menuItems}
                onClick={handleMenuClick}
            />
        </Sider>
    );
}
