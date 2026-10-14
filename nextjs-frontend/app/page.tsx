"use client";

import { Button } from "antd";
import { FileTextOutlined, ArrowRightOutlined } from "@ant-design/icons";
import { useRouter } from "next/navigation";
import { signIn, useSession } from "next-auth/react";
import { useEffect, useState } from "react";

const isAuthDisabled = process.env.NEXT_PUBLIC_DISABLE_AUTH === "true";

export default function LoginPage() {
    const router = useRouter();
    const { data: session, status } = useSession();
    const [isRedirecting, setIsRedirecting] = useState(false);

    useEffect(() => {
        if (!isAuthDisabled && status === "authenticated") {
            setIsRedirecting(true);
            router.push("/generate/design-document");
        }
    }, [status, router]);

    const handleLogin = () => {
        if (isAuthDisabled) {
            router.push("/generate/design-document");
        } else {
            signIn("azure-ad", { callbackUrl: "/generate/design-document" });
        }
    };

    if (!isAuthDisabled && (status === "loading" || isRedirecting)) {
        return (
            <div className="d-flex align-items-center justify-content-center min-vh-100 bg-light">
                <div className="spinner-border text-primary" role="status">
                    <span className="visually-hidden">Loading...</span>
                </div>
            </div>
        );
    }

    return (
        <div className="d-flex align-items-center justify-content-center min-vh-100 bg-light">
            <div className="bg-white p-5 rounded shadow-sm text-center w-100" style={{ maxWidth: '500px' }}>
                <div
                    className="bg-primary text-white d-flex align-items-center justify-content-center mx-auto mb-4"
                    style={{ width: "64px", height: "64px", fontSize: "32px", borderRadius: "12px" }}
                >
                    <FileTextOutlined />
                </div>

                <h3 className="fw-bold mb-2">
                    {isAuthDisabled ? "Welcome to Document Generator" : "Welcome Back"}
                </h3>
                <p className="text-muted mb-4">
                    {isAuthDisabled
                        ? "Design & Configuration Document Converter"
                        : "Sign in to Enfrasys Document Generator"}
                </p>

                {isAuthDisabled ? (
                    <Button
                        type="primary"
                        size="large"
                        className="w-100 d-flex align-items-center justify-content-center gap-2"
                        style={{ height: '48px', fontSize: '15px', fontWeight: 500 }}
                        onClick={handleLogin}
                    >
                        <span>Let&apos;s get started</span>
                        <ArrowRightOutlined />
                    </Button>
                ) : (
                    <Button
                        type="default"
                        size="large"
                        className="w-100 d-flex align-items-center justify-content-center gap-2"
                        style={{ height: '48px', fontSize: '15px', fontWeight: 500 }}
                        onClick={handleLogin}
                    >
                        {/* Minimal SVG for Microsoft Logo */}
                        <svg xmlns="http://www.w3.org/2000/svg" width="21" height="21" viewBox="0 0 21 21">
                            <path fill="#f25022" d="M1 1h9v9H1z" />
                            <path fill="#00a4ef" d="M1 11h9v9H1z" />
                            <path fill="#7fba00" d="M11 1h9v9h-9z" />
                            <path fill="#ffb900" d="M11 11h9v9h-9z" />
                        </svg>
                        Sign in with Microsoft
                    </Button>
                )}
            </div>
        </div>
    );
}