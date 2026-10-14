"use client";

import { useEffect, useState, useRef, useCallback, useMemo } from "react";
import { Button, Input, Modal, Spin, FloatButton, Typography, Alert, Space } from "antd";
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
import { resolveCost } from "@/lib/cost";

const { Text, Title } = Typography;

interface ReviewViewProps {
  generatorType?: string;   // e.g. "design-document" (not needed when opened from History)
  fromHistory?: boolean;    // true when rendered under /history/[taskId]
}

export default function ReviewView({ generatorType = "design-document", fromHistory = false }: ReviewViewProps) {
  const router = useRouter();
  const params = useParams();
  const searchParams = useSearchParams();
  const taskId = params.taskId as string;
  const isFromHistory = fromHistory || searchParams.get('from') === 'history';

  const basePath = fromHistory ? "/history" : `/generate/${generatorType}`;

  const [docName, setDocName] = useState("Generated_Design_Document");
  const [taskRecord, setTaskRecord] = useState<any>(null);
  const [costMetrics, setCostMetrics] = useState<any>(null);
  const [fetchErrors, setFetchErrors] = useState<any[]>([]);
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
        setTaskRecord(statusData);
        if (statusData.cost_metrics) setCostMetrics(statusData.cost_metrics);
        if (statusData.generated_filename) setDocName(statusData.generated_filename);
        if (Array.isArray(statusData.fetch_errors)) setFetchErrors(statusData.fetch_errors);
      }

      // Fetch user-configured pricing from settings (used as fallback for legacy records)
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

  // Unified cost calculation: prefers locked cost_summary snapshot, falls back to settings
  const computedCost = useMemo(() => {
    return resolveCost(taskRecord, pricingSettings);
  }, [taskRecord, pricingSettings]);

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
      router.push(basePath);
    }
  }, [router, fetchDocxData, taskId, basePath]);

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
    router.push(basePath);
  };

  if (!isClient) return null;

  return (
    <div
      className="d-flex w-100"
      style={{ fontFamily: "Inter, sans-serif", backgroundColor: "#f5f5f5", height: "100%", overflow: "hidden" }}
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
            <Space.Compact block size="large">
              <Input
                size="large"
                value={docName}
                onChange={(e) => setDocName(e.target.value)}
                disabled={isFromHistory}
                suffix={!isFromHistory && <EditOutlined className="text-muted" />}
                style={{ fontWeight: 500 }}
              />
              <Space.Addon>
                <span style={{ fontWeight: 500, color: "#666" }}>.docx</span>
              </Space.Addon>
            </Space.Compact>
          </div>

          {/* Cost Metrics */}
          {computedCost && (
            <div className="mb-4 p-3 bg-light rounded border">
              <div className="d-flex justify-content-between align-items-center mb-1">
                <Text type="secondary" className="fw-bold d-block" style={{ fontSize: "12px", letterSpacing: "0.5px" }}>
                  {computedCost.isLocked ? "GENERATION COST (LOCKED)" : "ESTIMATED GENERATION COST"}
                </Text>
                {computedCost.isLocked && (
                  <span className="badge bg-success-subtle text-success border border-success-subtle" style={{ fontSize: "11px", fontWeight: 600 }}>
                    Rate Locked
                  </span>
                )}
              </div>
              <Text className="fs-4 fw-bold text-success">
                RM {computedCost.costMyr.toFixed(2)}
              </Text>
              <Text type="secondary" className="d-block mt-1" style={{ fontSize: "12px" }}>
                USD ${computedCost.costUsd.toFixed(4)} &bull; Rate: 1 USD = RM {computedCost.exchangeRate.toFixed(2)}
              </Text>
            </div>
          )}

          {/* 6.5: resources that could not be fetched from Azure (left out of the document) */}
          {fetchErrors.length > 0 && (
            <Alert
              type="warning"
              showIcon
              className="mb-4"
              title={`${fetchErrors.length} resource${fetchErrors.length > 1 ? "s were" : " was"} not documented`}
              description={
                <div style={{ fontSize: "12px" }}>
                  <div className="mb-2">These could not be read from Azure, so they are not in the document:</div>
                  <ul className="mb-0 ps-3">
                    {fetchErrors.map((e: any, i: number) => (
                      <li key={i}>
                        <span className="fw-medium">{String(e.id || "").split("/").pop()}</span>
                        {e.type ? <span className="text-muted"> ({e.type})</span> : null}
                      </li>
                    ))}
                  </ul>
                </div>
              }
            />
          )}
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
            className="fw-medium text-white mb-3"
            style={{ height: "50px", backgroundColor: "#2b5aee" }}
          >
            Download DOCX
          </Button>

          {isFromHistory && (
            <Button
              block
              size="large"
              icon={<ArrowLeftOutlined className="me-2" />}
              onClick={() => router.push("/history")}
              className="fw-medium"
              style={{ height: "50px" }}
            >
              Back to History
            </Button>
          )}

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
