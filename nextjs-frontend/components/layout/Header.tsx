"use client";

import { FileTextOutlined, ExclamationCircleOutlined, HistoryOutlined, SettingOutlined } from "@ant-design/icons";
import { usePathname, useRouter } from "next/navigation";
import { Modal } from "antd";

export default function Header() {
    const pathname = usePathname();
    const router = useRouter();

    // Simple logic to highlight the correct step in the progress indicator
    // Supports both old flat routes and new /tasks/[taskId]/* routes
    let step = 1;
    if (pathname === "/preview" || pathname.endsWith("/setup")) step = 2;
    else if (pathname === "/processing" || pathname.endsWith("/processing")) step = 3;
    else if (pathname === "/edit" || pathname.endsWith("/review")) step = 4;

    const tabItems = [
        {
            path: '/',
            label: 'Generate',
            icon: <FileTextOutlined />,
            disabled: false
        },
        {
            path: '/history',
            label: 'History',
            icon: <HistoryOutlined />,
            disabled: false
        },
        {
            path: '/settings',
            label: 'Settings',
            icon: <SettingOutlined />,
            disabled: false
        },
    ];

    // Check if user is in an active workflow (setup or processing)
    const isInActiveWorkflow = pathname.endsWith("/setup") || pathname === "/preview"
        || pathname.endsWith("/processing") || pathname === "/processing";

    const navigateWithGuard = (targetPath: string) => {
        // If already on the target page, do nothing
        if (pathname === targetPath) return;

        if (isInActiveWorkflow) {
            Modal.confirm({
                title: "Leave Current Process?",
                icon: <ExclamationCircleOutlined />,
                content: "You are currently in the middle of a document generation process. Are you sure you want to leave?",
                okText: "Yes, leave",
                cancelText: "Cancel",
                okButtonProps: { danger: true },
                onOk() {
                    sessionStorage.clear();
                    router.push(targetPath);
                },
            });
        } else {
            router.push(targetPath);
        }
    };

    const handleLogoClick = () => {
        if (pathname === "/") return;
        navigateWithGuard("/");
    };

    return (
        <header className="bg-white border-bottom px-4 py-3 d-flex align-items-center justify-content-between sticky-top">
            {/* Brand Logo & Name */}
            <div
                className="d-flex align-items-center"
                onClick={handleLogoClick}
                style={{ cursor: "pointer" }} // Added pointer to indicate it is clickable
            >
                <div
                    className="bg-primary text-white d-flex align-items-center justify-content-center me-3"
                    style={{ width: "40px", height: "40px", fontSize: "20px", borderRadius: "8px" }}
                >
                    <FileTextOutlined />
                </div>
                <h5 className="mb-0 fw-semibold text-dark d-flex align-items-center" style={{ fontSize: "1.15rem" }}>
                    Enfrasys Document Generator
                </h5>
            </div>

            <div>
                <div className="d-flex align-items-center" style={{ gap: "8px" }}>
                    {tabItems.map((item) => {
                        // Check if the current path matches the item's path
                        // For Dashboard ('/'), we want an exact match.
                        // For others, we check if it starts with the path.
                        const isActive = item.path === '/' ? pathname === '/' : pathname.startsWith(item.path);
                        return (
                            <button
                                key={item.path}
                                onClick={() => {
                                    if (!item.disabled) navigateWithGuard(item.path);
                                }}
                                disabled={item.disabled}
                                className="btn d-flex align-items-center gap-2 border-0 shadow-none"
                                style={{
                                    padding: "8px 16px",
                                    borderRadius: "8px",
                                    fontSize: "15px",
                                    fontWeight: isActive ? "600" : "500",
                                    color: item.disabled ? "#adb5bd" : (isActive ? "#1c2b36" : "#6c757d"),
                                    backgroundColor: isActive ? "#f0f4f8" : "transparent",
                                    transition: "all 0.2s ease",
                                    cursor: item.disabled ? "not-allowed" : "pointer",
                                    opacity: item.disabled ? 0.6 : 1
                                }}
                            >
                                <span style={{ fontSize: "16px" }}>{item.icon}</span>
                                <span>{item.label}</span>
                            </button>
                        );
                    })}
                </div>
            </div>

            {/* Stepper & Help */}
            {/* Dynamic Progress Indicator */}
            {/* <div className="d-flex align-items-center">
                <div className="d-flex align-items-center gap-2">
                    {[1, 2, 3, 4].map((s) => (
                        <div
                            key={s}
                            className={`rounded-pill ${s <= step ? "bg-primary" : "bg-secondary bg-opacity-25"}`}
                            style={{ height: "4px", width: "40px", transition: "all 0.3s ease" }}
                        ></div>
                    ))}
                </div>
            </div> */}

        </header>
    );
}