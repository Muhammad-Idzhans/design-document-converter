"use client";

import { useEffect, useState, useRef, useCallback, useMemo } from "react";
import { Button, Input, Modal, Spin, FloatButton, Typography } from "antd";
import {
  FileWordOutlined,
  EditOutlined,
  ExclamationCircleOutlined,
  LoadingOutlined,
  ReloadOutlined,
  WarningOutlined,
  ArrowLeftOutlined,
  UndoOutlined
} from "@ant-design/icons";
import { useRouter, useParams, useSearchParams } from "next/navigation";
import { renderAsync } from "docx-preview";

const { Text, Title } = Typography;

export default function ReviewPage() {
  const router = useRouter();
  const params = useParams();
  const searchParams = useSearchParams();
  const taskId = params.taskId as string;
  const isFromHistory = searchParams.get('from') === 'history';

  const [docName, setDocName] = useState("Generated_Design_Document");
  const [costMetrics, setCostMetrics] = useState<any>(null);
  const [pricingSettings, setPricingSettings] = useState<any>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [loadError, setLoadError] = useState("");
  const [isClient, setIsClient] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);

  // NEW: State to hold the downloaded document before rendering
  const [docBlob, setDocBlob] = useState<Blob | null>(null);

  const docxContainerRef = useRef<HTMLDivElement>(null);

  // 1. We ONLY fetch the data here, we do not render it yet.
  const fetchDocxData = useCallback(async (tid: string) => {
    setIsLoading(true);
    setLoadError("");
    setDocBlob(null);

    try {
      const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";

      // Fetch Cost Metrics (raw token/page counts from backend)
      const statusRes = await fetch(`${API_BASE_URL}/api/status/${tid}`);
      if (statusRes.ok) {
        const statusData = await statusRes.json();
        if (statusData.cost_metrics) setCostMetrics(statusData.cost_metrics);
        if (statusData.generated_filename) setDocName(statusData.generated_filename);
      }

      // Fetch user-configured pricing from settings.json
      const settingsRes = await fetch('/api/settings');
      if (settingsRes.ok) {
        const settingsData = await settingsRes.json();
        if (!settingsData.error) setPricingSettings(settingsData);
      }

      // Fetch the actual DOCX file
      const docxRes = await fetch(`${API_BASE_URL}/api/download/${tid}?filename=preview`);
      if (!docxRes.ok) {
        throw new Error("Failed to fetch the DOCX file from the server.");
      }

      const blob = await docxRes.blob();

      // Save the file to state and turn off the loading spinner
      setDocBlob(blob);
      setIsLoading(false);

    } catch (err: any) {
      console.error("DOCX Fetch Error:", err);
      setLoadError(err.message || "Failed to load the document.");
      setIsLoading(false);
    }
  }, []);

  // 2. NEW: This watches for the DOM to be ready. 
  // Once isLoading is false, the div is on the screen, and we safely render the document!
  useEffect(() => {
    if (!isLoading && !loadError && docBlob && docxContainerRef.current) {
      // Clear container to prevent duplicate renders during React Strict Mode
      docxContainerRef.current.innerHTML = "";

      renderAsync(docBlob, docxContainerRef.current, undefined, {
        className: "docx-viewer",
        inWrapper: true,
        ignoreWidth: false,
        ignoreHeight: false,
        ignoreFonts: false,
        breakPages: true,
        ignoreLastRenderedPageBreak: false,
        experimental: false,
      }).catch((err) => {
        console.error("DOCX Render Error:", err);
        setLoadError("Failed to render the document visually.");
      });
    }
  }, [isLoading, loadError, docBlob]);

  // Dynamic cost calculation using user-configured pricing from settings.json
  const computedCost = useMemo(() => {
    if (!costMetrics || !pricingSettings) return null;

    // settings.json stores rates as "per 1M tokens" and "per 1K pages"
    const visionInputRate = (pricingSettings.visionInput || 0) / 1_000_000;
    const visionOutputRate = (pricingSettings.visionOutput || 0) / 1_000_000;
    const llmInputRate = (pricingSettings.llmInput || 0) / 1_000_000;
    const llmCompletionRate = (pricingSettings.llmCompletion || 0) / 1_000_000;
    const cuPageRate = (pricingSettings.contentUnderstanding || 0) / 1_000;
    const exchangeRate = pricingSettings.exchangeRate || 4.20;

    let costUsd = 0;
    costUsd += (costMetrics.vision_tokens_prompt || 0) * visionInputRate;
    costUsd += (costMetrics.vision_tokens_completion || 0) * visionOutputRate;
    costUsd += (costMetrics.llm_tokens_prompt || 0) * llmInputRate;
    costUsd += (costMetrics.llm_tokens_completion || 0) * llmCompletionRate;
    costUsd += (costMetrics.content_understanding_pages || 0) * cuPageRate;

    return {
      costUsd,
      costMyr: costUsd * exchangeRate
    };
  }, [costMetrics, pricingSettings]);

  const handleRetry = () => {
    if (!taskId) return;
    fetchDocxData(taskId);
  };

  useEffect(() => {
    setIsClient(true);
    const storedPreviewText = sessionStorage.getItem("documentPreview");

    if (storedPreviewText && !isFromHistory) {
      try {
        const preview = JSON.parse(storedPreviewText);
        let rawName = preview.filename || "Generated_Design_Document";
        if (rawName.startsWith("source_")) rawName = rawName.substring(7);
        if (rawName.toLowerCase().endsWith(".pptx")) rawName = rawName.slice(0, -5);

        // Remove spaces and hyphens, and ensure no multiple consecutive underscores
        rawName = rawName.replace(/[\s-]/g, "_").replace(/_+/g, "_");

        setDocName(`generated_${rawName}`);
      } catch (e) {
        console.error(e);
      }
    }

    if (taskId) {
      fetchDocxData(taskId);
    } else {
      router.push("/");
    }
  }, [router, fetchDocxData, taskId]);

  const handleDownloadDocx = async () => {
    if (!taskId) return;
    setIsDownloading(true);
    try {
      const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";
      const response = await fetch(`${API_BASE_URL}/api/download/${taskId}?filename=${encodeURIComponent(docName)}`);
      
      if (!response.ok) {
        throw new Error("Failed to download document");
      }
      
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${docName}.docx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error("Download error:", error);
      alert("Failed to download the document. Please try again.");
    } finally {
      setIsDownloading(false);
    }
  };

  const handleStartOver = () => {
    sessionStorage.clear();
    router.push("/");
  };

  if (!isClient) return null;

  return (
    <div
      className="d-flex w-100"
      style={{ fontFamily: "Inter, sans-serif", backgroundColor: "#f5f5f5", height: "calc(100vh - 73px)", overflow: "hidden" }}
    >
      {/* ── COLUMN 1: DOCUMENT VIEWER (LEFT - SCROLLABLE) ── */}
      <div
        id="scrollable-document-container"
        className="flex-grow-1 position-relative"
        style={{ height: "100%", overflowY: "auto", padding: "40px 20px" }}
      >
        <div className="d-flex justify-content-center">
          {isLoading ? (
            <div className="d-flex flex-column align-items-center justify-content-center" style={{ minHeight: "70vh" }}>
              <Spin indicator={<LoadingOutlined style={{ fontSize: 48 }} spin />} />
              <p className="text-muted mt-4 fw-medium fs-5">Rendering document preview...</p>
            </div>
          ) : loadError ? (
            <div className="d-flex flex-column align-items-center justify-content-center" style={{ minHeight: "70vh" }}>
              <div className="bg-white rounded-4 shadow-sm p-5 text-center d-flex flex-column align-items-center" style={{ maxWidth: "480px" }}>
                <div
                  className="d-flex align-items-center justify-content-center rounded-circle mb-4"
                  style={{ width: "64px", height: "64px", backgroundColor: "#fff2f0", border: "2px solid #ffccc7" }}
                >
                  <WarningOutlined style={{ fontSize: "28px", color: "#ff4d4f" }} />
                </div>
                <h5 className="fw-bold text-dark mb-2">Preview could not be loaded</h5>
                <p className="text-muted mb-1" style={{ fontSize: "14px" }}>{loadError}</p>
                <Button
                  type="primary"
                  icon={<ReloadOutlined />}
                  size="large"
                  onClick={handleRetry}
                  className="mt-4"
                  style={{ backgroundColor: "#2b5aee", borderRadius: "8px", fontWeight: 500 }}
                >
                  Retry Preview
                </Button>
              </div>
            </div>
          ) : (
            <div
              className="shadow-sm bg-white rounded"
              style={{
                maxWidth: "1000px",
                width: "100%",
                minHeight: "100%",
                // padding: "50px",
                overflowX: "auto"
              }}
            >
              {/* This div is now guaranteed to exist before renderAsync is called! */}
              <div ref={docxContainerRef} />
            </div>
          )}
        </div>

        <FloatButton.BackTop
          target={() => document.getElementById("scrollable-document-container") || window}
          visibilityHeight={300}
          type="primary"
          shape="circle"
          style={{ right: 500, bottom: 40, zIndex: 9999, width: "50px", height: "50px" }}
        />
      </div>

      {/* ── COLUMN 2: CONTROL SIDEBAR (RIGHT - FIXED) ── */}
      <div
        className="bg-white shadow-lg d-flex flex-column"
        style={{
          width: "440px",
          minWidth: "440px",
          height: "100%",
          overflowY: "auto",
          zIndex: 10
        }}
      >
        <div className="p-4">
          <Title level={4} className="mb-4" style={{ color: "#1f1f1f" }}>
            Document Details
          </Title>

          {/* Rename Field */}
          <div className="mb-4">
            <Text type="secondary" className="fw-bold d-block mb-2" style={{ fontSize: "12px", letterSpacing: "0.5px" }}>
              {isFromHistory ? "FILE NAME" : "RENAME FILE"}
            </Text>
            <Input
              size="large"
              value={docName}
              onChange={(e) => setDocName(e.target.value)}
              disabled={isFromHistory}
              suffix={!isFromHistory && <EditOutlined className="text-muted" />}
              addonAfter={<span style={{ fontWeight: 500, color: "#666" }}>.docx</span>}
              style={{ fontWeight: 500 }}
            />
          </div>

          {/* Cost Metrics — Dynamically computed using user-configured pricing */}
          {computedCost && (
            <div className="mb-4 p-3 bg-light rounded border">
              <Text type="secondary" className="fw-bold d-block mb-1" style={{ fontSize: "12px", letterSpacing: "0.5px" }}>
                ESTIMATED GENERATION COST
              </Text>
              <Text className="fs-4 fw-bold text-success">
                RM {computedCost.costMyr.toFixed(2)}
              </Text>
              <Text type="secondary" className="d-block mt-1" style={{ fontSize: "12px" }}>
                (USD ${computedCost.costUsd.toFixed(4)})
              </Text>
            </div>
          )}

          {/* --- OLD HARDCODED COST DISPLAY (fallback reference) ---
          {costMetrics && costMetrics.total_cost_myr !== undefined && (
            <div className="mb-4 p-3 bg-light rounded border">
              <Text type="secondary" className="fw-bold d-block mb-1" style={{ fontSize: "12px", letterSpacing: "0.5px" }}>
                ESTIMATED GENERATION COST
              </Text>
              <Text className="fs-4 fw-bold text-success">
                RM {Number(costMetrics.total_cost_myr).toFixed(2)}
              </Text>
            </div>
          )}
          */}
        </div>

        {/* Action Buttons */}
        <div className="p-4 mt-auto border-top bg-white">
          <Button
            type="primary"
            size="large"
            block
            icon={<FileWordOutlined className="me-2 text-white" />}
            onClick={handleDownloadDocx}
            disabled={isLoading || isDownloading}
            loading={isDownloading}
            className={`fw-medium text-white ${!isFromHistory ? 'mb-3' : ''}`}
            style={{ height: "50px", backgroundColor: "#2b5aee" }}
          >
            Download DOCX
          </Button>

          {!isFromHistory && (
          <Button
            danger
            block
            size="large"
            icon={<UndoOutlined className="me-2" />}
            onClick={handleStartOver}
            className="fw-medium"
            style={{ height: "50px" }}
          >
            Start Over
          </Button>
          )}
        </div>
      </div>
    </div>
  );
}
