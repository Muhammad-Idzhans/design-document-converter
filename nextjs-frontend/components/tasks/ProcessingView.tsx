"use client";

import { useEffect, useRef, useState } from "react";
import { Steps, Button, Spin } from "antd";
import { CheckCircleFilled, LoadingOutlined } from "@ant-design/icons";
import { useRouter, useParams } from "next/navigation";

interface ProcessingViewProps {
  generatorType: string; // e.g. "design-document"
}

export default function ProcessingView({ generatorType }: ProcessingViewProps) {
  const router = useRouter();
  const params = useParams();
  const taskId = params.taskId as string;

  const basePath = `/generate/${generatorType}`;

  // UI Loading Steps
  const [currentStep, setCurrentStep] = useState(0);
  const [isCompleted, setIsCompleted] = useState(false);
  const [hasError, setHasError] = useState(false);
  const [errorMessage, setErrorMessage] = useState("");
  const [stepName, setStepName] = useState("");
  const notFoundRef = useRef(0);

  // const designDocumentSteps = [
  //   "Extracting Content & Images",
  //   "Analyzing Architecture Diagrams via Vision AI",
  //   "Orchestrating Document Structure",
  //   "Drafting Markdown Sections",
  //   "Finalizing Word Document Format"
  // ];

  // const finalConfigSteps = [
  //   "Connecting to Azure Tenant",
  //   "Fetching Azure Resources & Configurations",
  //   "Orchestrating Document Structure",
  //   "Drafting Markdown Sections",
  //   "Finalizing Word Document Format"
  // ];

  // A step is complete once the backend's progress reaches its doneAt value
  const designDocumentSteps = [
    { label: "Extracting Content & Images", doneAt: 50 },
    { label: "Analyzing Architecture Diagrams via Vision AI", doneAt: 70 },
    { label: "Orchestrating Document Structure", doneAt: 85 },
    { label: "Drafting Markdown Sections", doneAt: 95 },
    { label: "Finalizing Word Document Format", doneAt: 100 },
  ];

  const finalConfigSteps = [
    { label: "Fetching Azure Resources & Configurations", doneAt: 50 },
    { label: "Writing Service Overviews with AI", doneAt: 85 },
    { label: "Building Word Document", doneAt: 100 },
  ];

  const steps = generatorType === "final-config" ? finalConfigSteps : designDocumentSteps;
  const stepsLabels = steps.map((s) => s.label);

  // 1. On Mount: Validate Task ID
  useEffect(() => {
    if (!taskId) {
      router.push(basePath);
    }
  }, [taskId, router, basePath]);

  // // 2. Fake UI Progression (since the backend script blocks synchronously right now)
  // useEffect(() => {
  //   if (isCompleted || hasError) return;

  //   // Slowly increment steps visually up to step 3, wait at step 3 for completion
  //   const timers = generatorType === "final-config" ? [
  //     setTimeout(() => setCurrentStep(1), 3000),
  //     setTimeout(() => setCurrentStep(2), 15000),
  //     setTimeout(() => setCurrentStep(3), 30000),
  //     setTimeout(() => setCurrentStep(4), 70000)
  //   ] : [
  //     setTimeout(() => setCurrentStep(1), 8000),
  //     setTimeout(() => setCurrentStep(2), 25000),
  //     setTimeout(() => setCurrentStep(3), 60000),
  //     setTimeout(() => setCurrentStep(4), 110000)
  //   ];

  //   return () => timers.forEach((t) => clearTimeout(t));
  // }, [isCompleted, hasError]);

  // 3. Polling the FastAPI Backend
  useEffect(() => {
    if (!taskId || isCompleted || hasError) return;

    const pollInterval = setInterval(async () => {
      try {
        const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";
        // const res = await fetch(`${API_BASE_URL}/api/status/${taskId}`);
        // if (!res.ok) return;

        // const data = await res.json();

        const res = await fetch(`${API_BASE_URL}/api/status/${taskId}`);
        if (!res.ok) {
          // 6.4: stop polling if the task keeps returning 404 (about 30 seconds)
          if (res.status === 404) {
            notFoundRef.current += 1;
            if (notFoundRef.current >= 6) {
              setHasError(true);
              setErrorMessage("This task could not be found. It may have been deleted or the server was restarted.");
              clearInterval(pollInterval);
            }
          }
          return;
        }
        notFoundRef.current = 0;

        const data = await res.json();

        // 6.3: drive the stepper from the backend's real progress
        if (data.status === "processing") {
          const progress = Number(data.progress) || 0;
          setCurrentStep(steps.filter((s) => progress >= s.doneAt).length);
          if (data.step_name) setStepName(data.step_name);
        }

        if (data.status === "completed") {
          setIsCompleted(true);
          setCurrentStep(steps.length); // All done!
          clearInterval(pollInterval);
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
              <p className="text-muted">
                {generatorType === "final-config"
                  ? "This may take 1-2 minutes depending on the number of resources selected."
                  : "This may take 3-5 minutes depending on the length of your presentation."}
              </p>
              {stepName && !isCompleted && (
                <p className="text-primary small fw-medium mt-2 mb-0">{stepName}</p>
              )}
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
                  onClick={() => router.push(`${basePath}/tasks/${taskId}/review`)}
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
              <div style={{ fontSize: "64px", lineHeight: 1, marginBottom: "16px" }}>⚠️</div>
            </div>
            <h3 className="fw-bold text-dark mb-2">Processing Failed</h3>
            <p className="text-muted mb-5">{errorMessage}</p>
            <Button
              type="primary"
              danger
              size="large"
              onClick={() => router.push(basePath)}
              style={{ borderRadius: "8px", padding: "0 30px" }}
            >
              {generatorType === "final-config" ? "Back to Final Config" : "Return to Upload"}
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
