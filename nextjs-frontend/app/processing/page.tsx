"use client";

import { useEffect, useState } from "react";
import { Steps, Button, Spin } from "antd";
import { CheckCircleFilled, LoadingOutlined } from "@ant-design/icons";
import { useRouter } from "next/navigation";

export default function ProcessingPage() {
  const router = useRouter();
  const [taskId, setTaskId] = useState<string | null>(null);

  // UI Loading Steps
  const [currentStep, setCurrentStep] = useState(0);
  const [isCompleted, setIsCompleted] = useState(false);
  const [hasError, setHasError] = useState(false);
  const [errorMessage, setErrorMessage] = useState("");

  const stepsLabels = [
    "Extracting Content & Images",
    "Analyzing Architecture Diagrams via Vision AI",
    "Orchestrating Document Structure",
    "Drafting Markdown Sections",
    "Finalizing Word Document Format"
  ];

  // 1. On Mount: Get Task ID
  useEffect(() => {
    const storedTaskId = sessionStorage.getItem("documentTaskId");
    if (!storedTaskId) {
      router.push("/");
    } else {
      setTaskId(storedTaskId);
    }
  }, [router]);

  // 2. Fake UI Progression (since the backend script blocks synchronously right now)
  useEffect(() => {
    if (isCompleted || hasError) return;

    // Slowly increment steps visually up to step 3, wait at step 3 for completion
    const timers = [
      setTimeout(() => setCurrentStep(1), 8000),   // After 8s: Move to Vision AI
      setTimeout(() => setCurrentStep(2), 25000),  // After 25s: Move to Orchestration
      setTimeout(() => setCurrentStep(3), 60000),  // After 60s: Move to Drafting
      setTimeout(() => setCurrentStep(4), 110000)  // After 110s: Move to Finalizing
    ];

    return () => timers.forEach((t) => clearTimeout(t));
  }, [isCompleted, hasError]);

  // 3. Polling the FastAPI Backend
  useEffect(() => {
    if (!taskId || isCompleted || hasError) return;

    const pollInterval = setInterval(async () => {
      try {
        const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";
        const res = await fetch(`${API_BASE_URL}/api/status/${taskId}`);
        if (!res.ok) return;

        const data = await res.json();

        if (data.status === "completed") {
          setIsCompleted(true);
          setCurrentStep(5); // All done!
          clearInterval(pollInterval);

          // Save generated content to session storage for the Review Page Phase 4
          if (data.markdown_draft) {
            sessionStorage.setItem("documentMarkdown", data.markdown_draft);
          }
          if (data.asset_library) {
            sessionStorage.setItem("documentAssets", JSON.stringify(data.asset_library));
          }
        }
        else if (data.status === "failed") {
          setHasError(true);
          setErrorMessage(data.error || "Backend pipeline threw an exception.");
          clearInterval(pollInterval);
        }

      } catch (err) {
        console.warn("Polling error:", err);
      }
    }, 5000); // Check every 5 seconds

    return () => clearInterval(pollInterval);
  }, [taskId, isCompleted, hasError]);


  // RENDER HELPERS
  const StepItems = stepsLabels.map((title, index) => {
    let icon, status: "wait" | "process" | "finish" | "error" = "wait";

    if (hasError && index === currentStep) {
      status = "error";
    }
    else if (index < currentStep || isCompleted) {
      status = "finish";
      icon = <CheckCircleFilled className="text-success" style={{ fontSize: '24px' }} />;
    } else if (index === currentStep) {
      status = "process";
      icon = <Spin indicator={<LoadingOutlined style={{ fontSize: 24 }} spin />} />;
    } else {
      status = "wait";
    }

    return {
      title: <span className={index <= currentStep ? "fw-bold" : "text-muted"}>{title}</span>,
      status,
      icon,
    };
  });

  return (
    <div className="container-fluid bg-white d-flex flex-grow-1 align-items-center justify-content-center py-5">
      <div
        className="card shadow-lg border-0"
        style={{ width: "100%", maxWidth: "700px", borderRadius: "16px", padding: "50px" }}
      >

        {!hasError ? (
          <>
            <div className="text-center mb-5">
              <h2 className="fw-bold fs-3 text-dark mb-2">AI is generating your document ...</h2>
              <p className="text-muted">This may take 3-5 minutes depending on the length of your presentation.</p>
            </div>

            <div className="px-md-5 mx-md-4 custom-stepper-wrapper">
              <Steps
                direction="vertical"
                current={currentStep}
                items={StepItems}
                size="default"
              />
            </div>

            {isCompleted && (
              <div className="text-center mt-5 animation-fade-in">
                <Button
                  type="primary"
                  size="large"
                  onClick={() => router.push("/edit")}
                  style={{
                    height: "50px",
                    padding: "0 40px",
                    fontSize: "16px",
                    fontWeight: "600",
                    borderRadius: "8px",
                    backgroundColor: "#2b5aee"
                  }}
                >
                  Preview Generated Document
                </Button>
              </div>
            )}
          </>
        ) : (
          <div className="text-center py-5">
            <div className="mb-4 text-danger">
              {/* Optional: using a generic failure icon from Ant Design if imported, or just generic styles */}
              <div style={{ fontSize: "64px", lineHeight: 1, marginBottom: "16px" }}>⚠️</div>
            </div>
            <h3 className="fw-bold text-dark mb-2">Processing Failed</h3>
            <p className="text-muted mb-5">{errorMessage}</p>
            <Button
              type="primary"
              danger
              size="large"
              onClick={() => router.push("/")}
              style={{ borderRadius: "8px", padding: "0 30px" }}
            >
              Return to Upload
            </Button>
          </div>
        )}

      </div>

      <style jsx global>{`
        /* Overriding some default Ant Design Stepper styles for larger appearance */
        .custom-stepper-wrapper .ant-steps-item-title {
          font-size: 16px !important;
          line-height: 28px !important;
          margin-left: 10px;
        }
        .custom-stepper-wrapper .ant-steps-item-tail {
           margin-left: 12px;
        }
        .animation-fade-in {
          animation: fadeIn 0.5s ease-in-out;
        }
        @keyframes fadeIn {
          from { opacity: 0; transform: translateY(10px); }
          to { opacity: 1; transform: translateY(0); }
        }
      `}</style>
    </div>
  );
}
