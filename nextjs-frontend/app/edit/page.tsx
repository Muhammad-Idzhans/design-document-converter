"use client";

import { useEffect, useState, useRef, useCallback } from "react";
import { Button, Input, Modal, Spin, FloatButton } from "antd";
import {
  FileWordOutlined,
  FilePdfOutlined,
  DownloadOutlined,
  LeftOutlined,
  EditOutlined,
  ExclamationCircleOutlined,
  LoadingOutlined,
  ReloadOutlined,
  WarningOutlined
} from "@ant-design/icons";
import { useRouter } from "next/navigation";

export default function ReviewPage() {
  const router = useRouter();

  const [taskId, setTaskId] = useState<string | null>(null);
  const [docName, setDocName] = useState("Generated_Design_Document");
  const [pageCount, setPageCount] = useState<number>(0);
  const [pageImages, setPageImages] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [loadError, setLoadError] = useState("");
  const [isClient, setIsClient] = useState(false);
  const [retryCount, setRetryCount] = useState(0);
  const [isRetrying, setIsRetrying] = useState(false);

  const scrollRef = useRef<HTMLDivElement>(null);

  // Extracted preview-fetch logic so it can be called on retry too
  const fetchPreview = useCallback(async (tid: string) => {
    setIsLoading(true);
    setLoadError("");

    const MAX_AUTO_RETRIES = 2;
    let lastErr = "";

    for (let attempt = 1; attempt <= MAX_AUTO_RETRIES; attempt++) {
      try {
        const res = await fetch(`http://localhost:8000/api/prepare-preview/${tid}`);
        if (!res.ok) {
          const body = await res.text().catch(() => "");
          throw new Error(`Server returned ${res.status}${body ? `: ${body}` : ""}`);
        }
        const data = await res.json();
        setPageCount(data.page_count || 0);
        setPageImages(data.page_images || []);
        setIsLoading(false);
        setIsRetrying(false);
        return; // success – exit
      } catch (err: any) {
        lastErr = err.message || "Failed to prepare document preview.";
        console.warn(`Preview attempt ${attempt}/${MAX_AUTO_RETRIES} failed:`, lastErr);
        if (attempt < MAX_AUTO_RETRIES) {
          // Wait 1.5s before the next automatic retry
          await new Promise((r) => setTimeout(r, 1500));
        }
      }
    }

    // All automatic retries exhausted
    setLoadError(lastErr);
    setIsLoading(false);
    setIsRetrying(false);
  }, []);

  const handleRetry = () => {
    if (!taskId) return;
    setRetryCount((c) => c + 1);
    setIsRetrying(true);
    fetchPreview(taskId);
  };

  useEffect(() => {
    setIsClient(true);
    const storedTaskId = sessionStorage.getItem("documentTaskId");
    const storedPreviewText = sessionStorage.getItem("documentPreviewData");

    // Derive file name from the original PPTX upload
    if (storedPreviewText) {
      try {
        const preview = JSON.parse(storedPreviewText);
        let rawName = preview.filename || "Generated_Design_Document";
        if (rawName.startsWith("source_")) rawName = rawName.substring(7);
        if (rawName.toLowerCase().endsWith(".pptx")) rawName = rawName.slice(0, -5);
        setDocName(`${rawName}_generated`);
      } catch (e) {
        console.error(e);
      }
    }

    if (storedTaskId) {
      setTaskId(storedTaskId);
      fetchPreview(storedTaskId);
    } else {
      router.push("/");
    }
  }, [router, fetchPreview]);

  // ── Download Handlers ──
  const handleDownloadDocx = () => {
    if (!taskId) return;
    window.open(
      `http://localhost:8000/api/download/${taskId}?filename=${encodeURIComponent(docName)}`,
      "_blank"
    );
  };

  const handleDownloadPdf = () => {
    if (!taskId) return;
    window.open(
      `http://localhost:8000/api/download-pdf/${taskId}?filename=${encodeURIComponent(docName)}`,
      "_blank"
    );
  };

  const handleDownloadBoth = () => {
    handleDownloadDocx();
    setTimeout(() => handleDownloadPdf(), 1500);
  };

  const handleStartOver = () => {
    Modal.confirm({
      title: "Generate New Document",
      icon: <ExclamationCircleOutlined />,
      content:
        "Are you sure you want to start over? Any generated files that you haven\u2019t downloaded yet will be lost.",
      okText: "Yes, start over",
      cancelText: "Cancel",
      okButtonProps: { danger: true },
      onOk() {
        sessionStorage.clear();
        router.push("/");
      },
    });
  };

  if (!isClient) return null;

  return (
    <div
      className="container-fluid min-vh-100 d-flex flex-column p-0"
      style={{ fontFamily: "Inter, sans-serif", backgroundColor: "#e0e0e0" }}
    >
      {/* ── HEADER ── */}
      <header className="bg-white border-bottom px-4 py-3 d-flex justify-content-between align-items-center shadow-sm sticky-top z-3">
        <div className="d-flex align-items-center gap-3">
          <Button
            type="text"
            onClick={handleStartOver}
            icon={<LeftOutlined />}
            className="text-muted fw-bold"
          >
            Back to Start
          </Button>

          <div className="d-flex align-items-center bg-light px-3 py-1 rounded-pill border">
            <FileWordOutlined className="text-primary me-2" style={{ fontSize: "18px" }} />
            <Input
              value={docName}
              onChange={(e: any) => setDocName(e.target.value)}
              variant="borderless"
              className="fw-bold p-0 m-0 text-dark"
              style={{ width: "300px", fontSize: "15px" }}
              suffix={<EditOutlined className="text-muted" />}
            />
            <span className="text-muted mx-3">|</span>
            <span className="text-muted small fw-medium" style={{ whiteSpace: "nowrap" }}>
              {pageCount > 0 ? `${pageCount} Pages` : isLoading ? "Preparing\u2026" : "\u2013"}
            </span>
          </div>
        </div>

        <div className="d-flex gap-2">
          <Button
            type="default"
            onClick={handleDownloadDocx}
            icon={<FileWordOutlined />}
            style={{ borderColor: "#2b5aee", color: "#2b5aee" }}
            disabled={isLoading}
          >
            Download DOCX
          </Button>
          <Button
            type="default"
            danger
            onClick={handleDownloadPdf}
            icon={<FilePdfOutlined />}
            disabled={isLoading}
          >
            Download PDF
          </Button>
          <Button
            type="primary"
            style={{ backgroundColor: "#2b5aee" }}
            onClick={handleDownloadBoth}
            icon={<DownloadOutlined />}
            disabled={isLoading}
          >
            Download Both
          </Button>
        </div>
      </header>

      {/* ── DOCUMENT PAGE VIEWER ── */}
      <div
        className="flex-grow-1 d-flex justify-content-center py-4 px-3"
        style={{ backgroundColor: "#d6d6d6" }}
      >        {isLoading ? (
          <div
            className="d-flex flex-column align-items-center justify-content-center"
            style={{ minHeight: "70vh" }}
          >
            <Spin indicator={<LoadingOutlined style={{ fontSize: 48 }} spin />} />
            <p className="text-muted mt-4 fw-medium fs-5">
              {isRetrying ? "Retrying document preview\u2026" : "Preparing document preview\u2026"}
            </p>
            <p className="text-muted small">
              Converting each page to a viewable image. This may take a moment for large documents.
            </p>
          </div>
        ) : loadError ? (
          <div
            className="d-flex flex-column align-items-center justify-content-center"
            style={{ minHeight: "70vh" }}
          >
            <div
              className="bg-white rounded-4 shadow-sm p-5 text-center d-flex flex-column align-items-center"
              style={{ maxWidth: "480px" }}
            >
              <div
                className="d-flex align-items-center justify-content-center rounded-circle mb-4"
                style={{
                  width: "64px",
                  height: "64px",
                  backgroundColor: "#fff2f0",
                  border: "2px solid #ffccc7",
                }}
              >
                <WarningOutlined style={{ fontSize: "28px", color: "#ff4d4f" }} />
              </div>
              <h5 className="fw-bold text-dark mb-2">Preview could not be loaded</h5>
              <p className="text-muted mb-1" style={{ fontSize: "14px" }}>
                {loadError}
              </p>
              {retryCount > 0 && (
                <p className="text-muted mb-0" style={{ fontSize: "12px" }}>
                  Attempted {retryCount} manual {retryCount === 1 ? "retry" : "retries"}
                </p>
              )}
              <Button
                type="primary"
                icon={<ReloadOutlined />}
                size="large"
                onClick={handleRetry}
                className="mt-4"
                style={{
                  backgroundColor: "#2b5aee",
                  borderRadius: "8px",
                  fontWeight: 500,
                  height: "44px",
                  paddingInline: "28px",
                }}
              >
                Retry Preview
              </Button>
              <p className="text-muted mt-3 mb-0" style={{ fontSize: "12px" }}>
                The document was generated successfully &mdash; downloads still work even if the preview fails.
              </p>
            </div>
          </div>
        ) : pageImages.length > 0 ? (
          <div
            className="d-flex flex-column align-items-center gap-3"
            style={{ maxWidth: "900px", width: "100%" }}
          >
            {pageImages.map((url, idx) => (
              <div
                key={idx}
                className="shadow position-relative bg-white"
                style={{ width: "100%" }}
              >
                <img
                  src={`http://localhost:8000${url}`}
                  alt={`Page ${idx + 1}`}
                  style={{ width: "100%", display: "block" }}
                />
                <div
                  className="position-absolute bottom-0 end-0 m-2 px-2 py-1 bg-dark bg-opacity-75 text-white rounded"
                  style={{ fontSize: "12px", pointerEvents: "none" }}
                >
                  Page {idx + 1} of {pageCount}
                </div>
              </div>
            ))}
          </div>
        ) : (
          <div
            className="d-flex flex-column align-items-center justify-content-center text-muted"
            style={{ minHeight: "70vh" }}
          >
            <p>No document pages available for preview.</p>
          </div>
        )}
      </div>

      {/* ── Scroll to Top Button (Global window listener) ── */}
      <FloatButton.BackTop
        visibilityHeight={200}
        type="primary"
        shape="circle"
        style={{ right: 40, bottom: 40, zIndex: 9999, width: "50px", height: "50px" }}
      />
    </div>
  );
}
