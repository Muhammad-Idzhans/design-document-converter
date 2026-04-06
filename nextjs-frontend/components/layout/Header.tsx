"use client";

import { FileTextOutlined } from "@ant-design/icons";
import Link from "next/link";
import { usePathname } from "next/navigation";

export default function Header() {
    const pathname = usePathname();

    // Simple logic to highlight the correct step in the progress indicator
    let step = 1;
    if (pathname === "/preview") step = 2;
    else if (pathname === "/processing") step = 3;
    else if (pathname === "/edit") step = 4;

    return (
        <header className="bg-white border-bottom px-4 py-3 d-flex align-items-center justify-content-between sticky-top">
            {/* Brand Logo & Name */}
            <div className="d-flex align-items-center">
                <div
                    className="bg-primary text-white d-flex align-items-center justify-content-center me-3"
                    style={{ width: "40px", height: "40px", fontSize: "20px", borderRadius: "8px" }}
                >
                    <FileTextOutlined />
                </div>
                <h5 className="mb-0 fw-semibold text-dark d-flex align-items-center" style={{ fontSize: "1.15rem" }}>
                    Design Document Generator
                </h5>
            </div>

            {/* Stepper & Help */}
            <div className="d-flex align-items-center">
                {/* Dynamic Progress Indicator */}
                <div className="d-flex align-items-center gap-2">
                    {[1, 2, 3, 4].map((s) => (
                        <div
                            key={s}
                            className={`rounded-pill ${s <= step ? "bg-primary" : "bg-secondary bg-opacity-25"}`}
                            style={{ height: "4px", width: "40px", transition: "all 0.3s ease" }}
                        ></div>
                    ))}
                </div>

                {/* <Link href="/help" className="text-secondary fw-bold text-decoration-none" style={{ fontSize: "0.85rem", letterSpacing: "0.5px" }}>
                    HELP
                </Link> */}
            </div>
        </header>
    );
}
