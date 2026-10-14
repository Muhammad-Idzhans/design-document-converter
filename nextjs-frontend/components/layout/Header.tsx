"use client";

import { FileTextOutlined, UserOutlined, LogoutOutlined, ExclamationCircleOutlined } from "@ant-design/icons";
import { useRouter } from "next/navigation";
import { Dropdown, Avatar, Modal } from "antd";
import type { MenuProps } from "antd";
import { useSession, signOut } from "next-auth/react";

const isAuthDisabled = process.env.NEXT_PUBLIC_DISABLE_AUTH === "true";

export default function Header() {
    const router = useRouter();
    const { data: session } = useSession();

    // Dynamically retrieve user info from Microsoft Entra ID session
    const userName = session?.user?.name || "Loading...";
    const userEmail = session?.user?.email || "Loading...";

    const handleLogoClick = () => {
        router.push("/generate/design-document");
    };

    const confirmSignOut = () => {
        Modal.confirm({
            title: "Log out of your account?",
            icon: <ExclamationCircleOutlined />,
            content: "Are you sure you want to log out? Any unsaved progress may be lost.",
            okText: "Log Out",
            okType: "danger",
            cancelText: "Cancel",
            onOk: async () => {
                await signOut({ callbackUrl: "/" });
            },
        });
    };

    const userMenuItems: MenuProps['items'] = [
        {
            key: 'user-info',
            label: (
                <div style={{ padding: '4px 0' }}>
                    <div style={{ fontWeight: 600, fontSize: '14px', color: '#1c2b36' }}>
                        {userName}
                    </div>
                    <div style={{ fontSize: '12px', color: '#6c757d' }}>
                        {userEmail}
                    </div>
                </div>
            ),
            disabled: true,
            style: { cursor: 'default' },
        },
        {
            type: 'divider',
        },
        {
            key: 'sign-out',
            icon: <LogoutOutlined />,
            label: 'Sign Out',
            danger: true,
            onClick: confirmSignOut,
        },
    ];

    return (
        <header
            style={{
                backgroundColor: '#1c2b36',
                padding: '0 24px',
                height: '56px',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                borderBottom: '1px solid #2a3f4d',
                position: 'sticky',
                top: 0,
                zIndex: 100,
            }}
        >
            {/* Brand Logo & Name */}
            <div
                onClick={handleLogoClick}
                style={{
                    display: 'flex',
                    alignItems: 'center',
                    cursor: 'pointer',
                }}
            >
                <div
                    style={{
                        width: '36px',
                        height: '36px',
                        fontSize: '18px',
                        borderRadius: '8px',
                        backgroundColor: '#1677ff',
                        color: '#fff',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        marginRight: '12px',
                    }}
                >
                    <FileTextOutlined />
                </div>
                <span
                    style={{
                        fontSize: '1.05rem',
                        fontWeight: 600,
                        color: '#ffffff',
                        letterSpacing: '-0.3px',
                    }}
                >
                    Enfrasys Document Generator
                </span>
            </div>

            {/* User Profile Dropdown (Only rendered when auth is active and user is signed in) */}
            {!isAuthDisabled && session?.user && (
                <Dropdown
                    menu={{ items: userMenuItems }}
                    trigger={['click']}
                    placement="bottomRight"
                >
                    <div
                        style={{
                            display: 'flex',
                            alignItems: 'center',
                            gap: '10px',
                            cursor: 'pointer',
                            padding: '4px 8px',
                            borderRadius: '8px',
                            transition: 'background-color 0.2s ease',
                        }}
                        onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'rgba(255,255,255,0.1)')}
                        onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
                    >
                        <div style={{ textAlign: 'right' }}>
                            <div style={{ fontSize: '13px', fontWeight: 600, color: '#ffffff', lineHeight: 1.3 }}>
                                {userName}
                            </div>
                            <div style={{ fontSize: '11px', color: 'rgba(255,255,255,0.6)', lineHeight: 1.3 }}>
                                {userEmail}
                            </div>
                        </div>
                        <Avatar
                            size={36}
                            icon={<UserOutlined />}
                            style={{ backgroundColor: '#1677ff', flexShrink: 0 }}
                        />
                    </div>
                </Dropdown>
            )}
        </header>
    );
}