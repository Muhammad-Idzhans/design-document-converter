// "use client";

// import { useEffect, useState, useRef, useCallback } from "react";
// import { Button, Input, Modal, Spin, FloatButton, Tag } from "antd";
// import {
//   FileWordOutlined,
//   FilePdfOutlined,
//   DownloadOutlined,
//   LeftOutlined,
//   EditOutlined,
//   ExclamationCircleOutlined,
//   LoadingOutlined,
//   ReloadOutlined,
//   WarningOutlined
// } from "@ant-design/icons";
// import { useRouter } from "next/navigation";

// export default function ReviewPage() {
//   const router = useRouter();

//   const [taskId, setTaskId] = useState<string | null>(null);
//   const [docName, setDocName] = useState("Generated_Design_Document");
//   const [pageCount, setPageCount] = useState<number>(0);
//   const [pageImages, setPageImages] = useState<string[]>([]);
//   const [costMetrics, setCostMetrics] = useState<any>(null);
//   const [isLoading, setIsLoading] = useState(true);
//   const [loadError, setLoadError] = useState("");
//   const [isClient, setIsClient] = useState(false);
//   const [retryCount, setRetryCount] = useState(0);
//   const [isRetrying, setIsRetrying] = useState(false);

//   const scrollRef = useRef<HTMLDivElement>(null);

//   // Extracted preview-fetch logic so it can be called on retry too
//   const fetchPreview = useCallback(async (tid: string) => {
//     setIsLoading(true);
//     setLoadError("");

//     const MAX_AUTO_RETRIES = 2;
//     let lastErr = "";

//     for (let attempt = 1; attempt <= MAX_AUTO_RETRIES; attempt++) {
//       try {
//         const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";
//         const res = await fetch(`${API_BASE_URL}/api/prepare-preview/${tid}`);
//         if (!res.ok) {
//           const body = await res.text().catch(() => "");
//           throw new Error(`Server returned ${res.status}${body ? `: ${body}` : ""}`);
//         }
//         const data = await res.json();
//         setPageCount(data.page_count || 0);
//         setPageImages(data.page_images || []);
//         if (data.cost_metrics) setCostMetrics(data.cost_metrics);
//         setIsLoading(false);
//         setIsRetrying(false);
//         return; // success – exit
//       } catch (err: any) {
//         lastErr = err.message || "Failed to prepare document preview.";
//         console.warn(`Preview attempt ${attempt}/${MAX_AUTO_RETRIES} failed:`, lastErr);
//         if (attempt < MAX_AUTO_RETRIES) {
//           // Wait 1.5s before the next automatic retry
//           await new Promise((r) => setTimeout(r, 1500));
//         }
//       }
//     }

//     // All automatic retries exhausted
//     setLoadError(lastErr);
//     setIsLoading(false);
//     setIsRetrying(false);
//   }, []);

//   const handleRetry = () => {
//     if (!taskId) return;
//     setRetryCount((c) => c + 1);
//     setIsRetrying(true);
//     fetchPreview(taskId);
//   };

//   useEffect(() => {
//     setIsClient(true);
//     const storedTaskId = sessionStorage.getItem("documentTaskId");
//     const storedPreviewText = sessionStorage.getItem("documentPreviewData");

//     // Derive file name from the original PPTX upload
//     if (storedPreviewText) {
//       try {
//         const preview = JSON.parse(storedPreviewText);
//         let rawName = preview.filename || "Generated_Design_Document";
//         if (rawName.startsWith("source_")) rawName = rawName.substring(7);
//         if (rawName.toLowerCase().endsWith(".pptx")) rawName = rawName.slice(0, -5);
//         setDocName(`${rawName}_generated`);
//       } catch (e) {
//         console.error(e);
//       }
//     }

//     if (storedTaskId) {
//       setTaskId(storedTaskId);
//       fetchPreview(storedTaskId);
//     } else {
//       router.push("/");
//     }
//   }, [router, fetchPreview]);

//   // ── Download Handlers ──
//   const handleDownloadDocx = () => {
//     if (!taskId) return;
//     Modal.info({
//       title: "Important: Viewing your Document",
//       content: (
//         <div>
//           <p>To ensure your document renders perfectly:</p>
//           <ol>
//             <li>Please open the downloaded file using <b>Microsoft Word</b> (not a web preview or Google Docs).</li>
//             <li>Upon opening, Microsoft Word will prompt you to update the fields (Table of Contents). <b>Click 'Yes'</b> to generate your final TOC.</li>
//           </ol>
//         </div>
//       ),
//       onOk() {
//         const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";
//         window.open(
//           `${API_BASE_URL}/api/download/${taskId}?filename=${encodeURIComponent(docName)}`,
//           "_blank"
//         );
//       },
//     });
//   };

//   const handleDownloadPdf = () => {
//     if (!taskId) return;
//     const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";
//     window.open(
//       `${API_BASE_URL}/api/download-pdf/${taskId}?filename=${encodeURIComponent(docName)}`,
//       "_blank"
//     );
//   };

//   const handleDownloadBoth = () => {
//     if (!taskId) return;
//     Modal.info({
//       title: "Important: Viewing your Document",
//       content: (
//         <div>
//           <p>To ensure your document renders perfectly:</p>
//           <ol>
//             <li>Please open the downloaded file using <b>Microsoft Word</b> (not a web preview or Google Docs).</li>
//             <li>Upon opening, Microsoft Word will prompt you to update the fields (Table of Contents). <b>Click 'Yes'</b> to generate your final TOC.</li>
//           </ol>
//         </div>
//       ),
//       onOk() {
//         const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";
//         window.open(
//           `${API_BASE_URL}/api/download/${taskId}?filename=${encodeURIComponent(docName)}`,
//           "_blank"
//         );
//         setTimeout(() => handleDownloadPdf(), 1500);
//       },
//     });
//   };

//   const handleStartOver = () => {
//     Modal.confirm({
//       title: "Generate New Document",
//       icon: <ExclamationCircleOutlined />,
//       content:
//         "Are you sure you want to start over? Any generated files that you haven\u2019t downloaded yet will be lost.",
//       okText: "Yes, start over",
//       cancelText: "Cancel",
//       okButtonProps: { danger: true },
//       onOk() {
//         sessionStorage.clear();
//         router.push("/");
//       },
//     });
//   };

//   if (!isClient) return null;

//   return (
//     <div
//       className="container-fluid min-vh-100 d-flex flex-column p-0"
//       style={{ fontFamily: "Inter, sans-serif", backgroundColor: "#e0e0e0" }}
//     >
//       {/* ── HEADER ── */}
//       <header className="bg-white border-bottom px-4 py-3 d-flex justify-content-between align-items-center shadow-sm sticky-top z-3">
//         <div className="d-flex align-items-center gap-3">
//           <Button
//             type="text"
//             onClick={handleStartOver}
//             icon={<LeftOutlined />}
//             className="text-muted fw-bold"
//           >
//             Back to Start
//           </Button>

//           <div className="d-flex align-items-center bg-light px-3 py-1 rounded-pill border">
//             <FileWordOutlined className="text-primary me-2" style={{ fontSize: "18px" }} />
//             <Input
//               value={docName}
//               onChange={(e: any) => setDocName(e.target.value)}
//               variant="borderless"
//               className="fw-bold p-0 m-0 text-dark"
//               style={{ width: "300px", fontSize: "15px" }}
//               suffix={<EditOutlined className="text-muted" />}
//             />
//             <span className="text-muted mx-3">|</span>
//             <span className="text-muted small fw-medium" style={{ whiteSpace: "nowrap" }}>
//               {pageCount > 0 ? `${pageCount} Pages` : isLoading ? "Preparing\u2026" : "\u2013"}
//             </span>
//           </div>
//           {costMetrics && costMetrics.total_cost_myr !== undefined && (
//             <Tag color="green" className="py-1 px-3 fs-6 rounded-pill border-0 shadow-sm" style={{ fontWeight: 600 }}>
//               Estimated Generation Cost: RM {Number(costMetrics.total_cost_myr).toFixed(2)}
//             </Tag>
//           )}
//         </div>

//         <div className="d-flex gap-2">
//           <Button
//             type="default"
//             onClick={handleDownloadDocx}
//             icon={<FileWordOutlined />}
//             style={{ borderColor: "#2b5aee", color: "#2b5aee" }}
//             disabled={isLoading}
//           >
//             Download DOCX
//           </Button>
//           <Button
//             type="default"
//             danger
//             onClick={handleDownloadPdf}
//             icon={<FilePdfOutlined />}
//             disabled={isLoading}
//           >
//             Download PDF
//           </Button>
//           {/* <Button
//             type="primary"
//             style={{ backgroundColor: "#2b5aee" }}
//             onClick={handleDownloadBoth}
//             icon={<DownloadOutlined />}
//             disabled={isLoading}
//           >
//             Download Both
//           </Button> */}
//         </div>
//       </header>

//       {/* ── DOCUMENT PAGE VIEWER ── */}
//       <div
//         className="flex-grow-1 d-flex justify-content-center py-4 px-3"
//         style={{ backgroundColor: "#d6d6d6" }}
//       >        {isLoading ? (
//         <div
//           className="d-flex flex-column align-items-center justify-content-center"
//           style={{ minHeight: "70vh" }}
//         >
//           <Spin indicator={<LoadingOutlined style={{ fontSize: 48 }} spin />} />
//           <p className="text-muted mt-4 fw-medium fs-5">
//             {isRetrying ? "Retrying document preview\u2026" : "Preparing document preview\u2026"}
//           </p>
//           <p className="text-muted small">
//             Converting each page to a viewable image. This may take a moment for large documents.
//           </p>
//         </div>
//       ) : loadError ? (
//         <div
//           className="d-flex flex-column align-items-center justify-content-center"
//           style={{ minHeight: "70vh" }}
//         >
//           <div
//             className="bg-white rounded-4 shadow-sm p-5 text-center d-flex flex-column align-items-center"
//             style={{ maxWidth: "480px" }}
//           >
//             <div
//               className="d-flex align-items-center justify-content-center rounded-circle mb-4"
//               style={{
//                 width: "64px",
//                 height: "64px",
//                 backgroundColor: "#fff2f0",
//                 border: "2px solid #ffccc7",
//               }}
//             >
//               <WarningOutlined style={{ fontSize: "28px", color: "#ff4d4f" }} />
//             </div>
//             <h5 className="fw-bold text-dark mb-2">Preview could not be loaded</h5>
//             <p className="text-muted mb-1" style={{ fontSize: "14px" }}>
//               {loadError}
//             </p>
//             {retryCount > 0 && (
//               <p className="text-muted mb-0" style={{ fontSize: "12px" }}>
//                 Attempted {retryCount} manual {retryCount === 1 ? "retry" : "retries"}
//               </p>
//             )}
//             <Button
//               type="primary"
//               icon={<ReloadOutlined />}
//               size="large"
//               onClick={handleRetry}
//               className="mt-4"
//               style={{
//                 backgroundColor: "#2b5aee",
//                 borderRadius: "8px",
//                 fontWeight: 500,
//                 height: "44px",
//                 paddingInline: "28px",
//               }}
//             >
//               Retry Preview
//             </Button>
//             <p className="text-muted mt-3 mb-0" style={{ fontSize: "12px" }}>
//               The document was generated successfully &mdash; downloads still work even if the preview fails.
//             </p>
//           </div>
//         </div>
//       ) : pageImages.length > 0 ? (
//         <div
//           className="d-flex flex-column align-items-center gap-3"
//           style={{ maxWidth: "900px", width: "100%" }}
//         >
//           {pageImages.map((url, idx) => (
//             <div
//               key={idx}
//               className="shadow position-relative bg-white"
//               style={{ width: "100%" }}
//             >
//               <img
//                 src={`${process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000"}${url}`}
//                 alt={`Page ${idx + 1}`}
//                 style={{ width: "100%", display: "block" }}
//               />
//               <div
//                 className="position-absolute bottom-0 end-0 m-2 px-2 py-1 bg-dark bg-opacity-75 text-white rounded"
//                 style={{ fontSize: "12px", pointerEvents: "none" }}
//               >
//                 Page {idx + 1} of {pageCount}
//               </div>
//             </div>
//           ))}
//         </div>
//       ) : (
//         <div
//           className="d-flex flex-column align-items-center justify-content-center text-muted"
//           style={{ minHeight: "70vh" }}
//         >
//           <p>No document pages available for preview.</p>
//         </div>
//       )}
//       </div>

//       {/* ── Scroll to Top Button (Global window listener) ── */}
//       <FloatButton.BackTop
//         visibilityHeight={200}
//         type="primary"
//         shape="circle"
//         style={{ right: 40, bottom: 40, zIndex: 9999, width: "50px", height: "50px" }}
//       />
//     </div>
//   );
// }



// "use client";

// import { useEffect, useState, useRef, useCallback } from "react";
// import { Button, Input, Modal, Spin, FloatButton, Tag } from "antd";
// import {
//   FileWordOutlined,
//   FilePdfOutlined,
//   LeftOutlined,
//   EditOutlined,
//   ExclamationCircleOutlined,
//   LoadingOutlined,
//   ReloadOutlined,
//   WarningOutlined
// } from "@ant-design/icons";
// import { useRouter } from "next/navigation";
// import { renderAsync } from "docx-preview";

// export default function ReviewPage() {
//   const router = useRouter();

//   const [taskId, setTaskId] = useState<string | null>(null);
//   const [docName, setDocName] = useState("Generated_Design_Document");
//   const [costMetrics, setCostMetrics] = useState<any>(null);
//   const [isLoading, setIsLoading] = useState(true);
//   const [loadError, setLoadError] = useState("");
//   const [isClient, setIsClient] = useState(false);

//   const docxContainerRef = useRef<HTMLDivElement>(null);

//   const fetchAndRenderDocx = useCallback(async (tid: string) => {
//     setIsLoading(true);
//     setLoadError("");

//     try {
//       const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";

//       const statusRes = await fetch(`${API_BASE_URL}/api/status/${tid}`);
//       if (statusRes.ok) {
//         const statusData = await statusRes.json();
//         if (statusData.cost_metrics) setCostMetrics(statusData.cost_metrics);
//       }

//       const docxRes = await fetch(`${API_BASE_URL}/api/download/${tid}?filename=preview`);
//       if (!docxRes.ok) {
//         throw new Error("Failed to fetch the DOCX file from the server.");
//       }
//       const blob = await docxRes.blob();

//       if (docxContainerRef.current) {
//         await renderAsync(blob, docxContainerRef.current, undefined, {
//           className: "docx-viewer",
//           inWrapper: true,
//           ignoreWidth: false,
//           ignoreHeight: false,
//           ignoreFonts: false,
//           breakPages: true,
//           ignoreLastRenderedPageBreak: false,
//           experimental: false,
//         });
//       }

//       setIsLoading(false);
//     } catch (err: any) {
//       console.error("DOCX Render Error:", err);
//       setLoadError(err.message || "Failed to render the document.");
//       setIsLoading(false);
//     }
//   }, []);

//   const handleRetry = () => {
//     if (!taskId) return;
//     fetchAndRenderDocx(taskId);
//   };

//   useEffect(() => {
//     setIsClient(true);
//     const storedTaskId = sessionStorage.getItem("documentTaskId");
//     const storedPreviewText = sessionStorage.getItem("documentPreviewData");

//     if (storedPreviewText) {
//       try {
//         const preview = JSON.parse(storedPreviewText);
//         let rawName = preview.filename || "Generated_Design_Document";
//         if (rawName.startsWith("source_")) rawName = rawName.substring(7);
//         if (rawName.toLowerCase().endsWith(".pptx")) rawName = rawName.slice(0, -5);
//         setDocName(`${rawName}_generated`);
//       } catch (e) {
//         console.error(e);
//       }
//     }

//     if (storedTaskId) {
//       setTaskId(storedTaskId);
//       fetchAndRenderDocx(storedTaskId);
//     } else {
//       router.push("/");
//     }
//   }, [router, fetchAndRenderDocx]);

//   const handleDownloadDocx = () => {
//     if (!taskId) return;
//     const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";
//     window.open(
//       `${API_BASE_URL}/api/download/${taskId}?filename=${encodeURIComponent(docName)}`,
//       "_blank"
//     );
//   };

//   const handleDownloadPdf = () => {
//     if (!taskId) return;
//     const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";
//     window.open(
//       `${API_BASE_URL}/api/download-pdf/${taskId}?filename=${encodeURIComponent(docName)}`,
//       "_blank"
//     );
//   };

//   const handleStartOver = () => {
//     Modal.confirm({
//       title: "Generate New Document",
//       icon: <ExclamationCircleOutlined />,
//       content: "Are you sure you want to start over? Any generated files that you haven’t downloaded yet will be lost.",
//       okText: "Yes, start over",
//       cancelText: "Cancel",
//       okButtonProps: { danger: true },
//       onOk() {
//         sessionStorage.clear();
//         router.push("/");
//       },
//     });
//   };

//   if (!isClient) return null;

//   return (
//     <div
//       className="container-fluid min-vh-100 d-flex flex-column p-0"
//       style={{ fontFamily: "Inter, sans-serif", backgroundColor: "#e0e0e0" }}
//     >
//       {/* ── DOCUMENT PAGE VIEWER ── */}
//       {/* Added pb-5 and marginBottom so the fixed footer doesn't overlap the bottom of the document */}
//       <div
//         className="flex-grow-1 d-flex justify-content-center py-4 px-3"
//         style={{ backgroundColor: "#d6d6d6", paddingBottom: "100px" }}
//       >
//         {isLoading ? (
//           <div className="d-flex flex-column align-items-center justify-content-center" style={{ minHeight: "70vh" }}>
//             <Spin indicator={<LoadingOutlined style={{ fontSize: 48 }} spin />} />
//             <p className="text-muted mt-4 fw-medium fs-5">Rendering document preview...</p>
//           </div>
//         ) : loadError ? (
//           <div className="d-flex flex-column align-items-center justify-content-center" style={{ minHeight: "70vh" }}>
//             <div className="bg-white rounded-4 shadow-sm p-5 text-center d-flex flex-column align-items-center" style={{ maxWidth: "480px" }}>
//               <div
//                 className="d-flex align-items-center justify-content-center rounded-circle mb-4"
//                 style={{ width: "64px", height: "64px", backgroundColor: "#fff2f0", border: "2px solid #ffccc7" }}
//               >
//                 <WarningOutlined style={{ fontSize: "28px", color: "#ff4d4f" }} />
//               </div>
//               <h5 className="fw-bold text-dark mb-2">Preview could not be loaded</h5>
//               <p className="text-muted mb-1" style={{ fontSize: "14px" }}>{loadError}</p>
//               <Button
//                 type="primary"
//                 icon={<ReloadOutlined />}
//                 size="large"
//                 onClick={handleRetry}
//                 className="mt-4"
//                 style={{ backgroundColor: "#2b5aee", borderRadius: "8px", fontWeight: 500 }}
//               >
//                 Retry Preview
//               </Button>
//             </div>
//           </div>
//         ) : (
//           <div
//             className="shadow bg-white rounded mb-5"
//             style={{
//               maxWidth: "1000px",
//               width: "100%",
//               minHeight: "800px",
//               padding: "40px",
//               overflowX: "auto"
//             }}
//           >
//             <div ref={docxContainerRef} />
//           </div>
//         )}
//       </div>

//       {/* ── BOTTOM CONTROL PANEL ── */}
//       <div className="bg-white border-top px-4 py-3 d-flex justify-content-between align-items-center shadow-lg position-fixed bottom-0 w-100 z-3">
//         <div className="d-flex align-items-center gap-3">
//           <Button
//             type="text"
//             onClick={handleStartOver}
//             icon={<LeftOutlined />}
//             className="text-muted fw-bold"
//           >
//             Back to Start
//           </Button>

//           <div className="d-flex align-items-center bg-light px-3 py-1 rounded-pill border">
//             <FileWordOutlined className="text-primary me-2" style={{ fontSize: "18px" }} />
//             <Input
//               value={docName}
//               onChange={(e: any) => setDocName(e.target.value)}
//               variant="borderless"
//               className="fw-bold p-0 m-0 text-dark"
//               style={{ width: "300px", fontSize: "15px" }}
//               suffix={<EditOutlined className="text-muted" />}
//             />
//             <span className="text-muted mx-3">|</span>
//             <span className="text-muted small fw-medium" style={{ whiteSpace: "nowrap" }}>
//               DOCX Preview
//             </span>
//           </div>
//           {costMetrics && costMetrics.total_cost_myr !== undefined && (
//             <Tag color="green" className="py-1 px-3 fs-6 rounded-pill border-0 shadow-sm" style={{ fontWeight: 600 }}>
//               Estimated Cost: RM {Number(costMetrics.total_cost_myr).toFixed(2)}
//             </Tag>
//           )}
//         </div>

//         <div className="d-flex gap-2">
//           <Button
//             type="default"
//             onClick={handleDownloadDocx}
//             icon={<FileWordOutlined />}
//             style={{ borderColor: "#2b5aee", color: "#2b5aee" }}
//             disabled={isLoading}
//           >
//             Download DOCX
//           </Button>
//           <Button
//             type="default"
//             danger
//             onClick={handleDownloadPdf}
//             icon={<FilePdfOutlined />}
//             disabled={isLoading}
//           >
//             Download PDF
//           </Button>
//         </div>
//       </div>

//       {/* Lifted the BackTop button slightly higher so it sits above the new bottom bar */}
//       <FloatButton.BackTop
//         visibilityHeight={200}
//         type="primary"
//         shape="circle"
//         style={{ right: 40, bottom: 100, zIndex: 9999, width: "50px", height: "50px" }}
//       />
//     </div>
//   );
// }

// "use client";

// import { useEffect, useState, useRef, useCallback } from "react";
// import { Button, Input, Modal, Spin, FloatButton, Typography } from "antd";
// import {
//   FileWordOutlined,
//   EditOutlined,
//   ExclamationCircleOutlined,
//   LoadingOutlined,
//   ReloadOutlined,
//   WarningOutlined,
//   ArrowLeftOutlined,
// } from "@ant-design/icons";
// import { useRouter } from "next/navigation";
// import { renderAsync } from "docx-preview";

// const { Text, Title } = Typography;

// export default function ReviewPage() {
//   const router = useRouter();

//   const [taskId, setTaskId] = useState<string | null>(null);
//   const [docName, setDocName] = useState("Generated_Design_Document");
//   const [costMetrics, setCostMetrics] = useState<any>(null);
//   const [isLoading, setIsLoading] = useState(true);
//   const [loadError, setLoadError] = useState("");
//   const [isClient, setIsClient] = useState(false);

//   const docxContainerRef = useRef<HTMLDivElement>(null);

//   const fetchAndRenderDocx = useCallback(async (tid: string) => {
//     setIsLoading(true);
//     setLoadError("");

//     try {
//       const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";

//       const statusRes = await fetch(`${API_BASE_URL}/api/status/${tid}`);
//       if (statusRes.ok) {
//         const statusData = await statusRes.json();
//         if (statusData.cost_metrics) setCostMetrics(statusData.cost_metrics);
//       }

//       const docxRes = await fetch(`${API_BASE_URL}/api/download/${tid}?filename=preview`);
//       if (!docxRes.ok) {
//         throw new Error("Failed to fetch the DOCX file from the server.");
//       }
//       const blob = await docxRes.blob();

//       if (docxContainerRef.current) {
//         await renderAsync(blob, docxContainerRef.current, undefined, {
//           className: "docx-viewer",
//           inWrapper: true,
//           ignoreWidth: false,
//           ignoreHeight: false,
//           ignoreFonts: false,
//           breakPages: true,
//           ignoreLastRenderedPageBreak: false,
//           experimental: false,
//         });
//       }

//       setIsLoading(false);
//     } catch (err: any) {
//       console.error("DOCX Render Error:", err);
//       setLoadError(err.message || "Failed to render the document.");
//       setIsLoading(false);
//     }
//   }, []);

//   const handleRetry = () => {
//     if (!taskId) return;
//     fetchAndRenderDocx(taskId);
//   };

//   useEffect(() => {
//     setIsClient(true);
//     const storedTaskId = sessionStorage.getItem("documentTaskId");
//     const storedPreviewText = sessionStorage.getItem("documentPreviewData");

//     if (storedPreviewText) {
//       try {
//         const preview = JSON.parse(storedPreviewText);
//         let rawName = preview.filename || "Generated_Design_Document";
//         if (rawName.startsWith("source_")) rawName = rawName.substring(7);
//         if (rawName.toLowerCase().endsWith(".pptx")) rawName = rawName.slice(0, -5);
//         setDocName(`${rawName}_generated`);
//       } catch (e) {
//         console.error(e);
//       }
//     }

//     if (storedTaskId) {
//       setTaskId(storedTaskId);
//       fetchAndRenderDocx(storedTaskId);
//     } else {
//       router.push("/");
//     }
//   }, [router, fetchAndRenderDocx]);

//   const handleDownloadDocx = () => {
//     if (!taskId) return;
//     const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";
//     window.open(
//       `${API_BASE_URL}/api/download/${taskId}?filename=${encodeURIComponent(docName)}`,
//       "_blank"
//     );
//   };

//   const handleStartOver = () => {
//     Modal.confirm({
//       title: "Start Over",
//       icon: <ExclamationCircleOutlined />,
//       content: "Are you sure you want to start over? Any generated files that you haven’t downloaded yet will be lost.",
//       okText: "Yes, start over",
//       cancelText: "Cancel",
//       okButtonProps: { danger: true },
//       onOk() {
//         sessionStorage.clear();
//         router.push("/");
//       },
//     });
//   };

//   if (!isClient) return null;

//   return (
//     <div
//       className="d-flex w-100 flex-grow-1"
//       /* CRITICAL CHANGE 1: 
//         We subtract an estimated 70px to account for your global <Header />. 
//         If your header is taller/shorter, adjust this '70px' value.
//         'overflow: hidden' prevents the main browser window from scrolling entirely.
//       */
//       style={{ fontFamily: "Inter, sans-serif", backgroundColor: "#f5f5f5", overflow: "hidden" }}
//     >
//       {/* ── COLUMN 1: DOCUMENT VIEWER (LEFT - SCROLLABLE) ── */}
//       <div
//         id="scrollable-document-container"
//         className="flex-grow-1 position-relative"
//         /* CRITICAL CHANGE 2: Only this specific div is allowed to scroll vertically */
//         style={{ overflowY: "auto", padding: "40px 20px" }}
//       >
//         <div className="d-flex justify-content-center">
//           {isLoading ? (
//             <div className="d-flex flex-column align-items-center justify-content-center" style={{ minHeight: "70vh" }}>
//               <Spin indicator={<LoadingOutlined style={{ fontSize: 48 }} spin />} />
//               <p className="text-muted mt-4 fw-medium fs-5">Rendering document preview...</p>
//             </div>
//           ) : loadError ? (
//             <div className="d-flex flex-column align-items-center justify-content-center" style={{ minHeight: "70vh" }}>
//               <div className="bg-white rounded-4 shadow-sm p-5 text-center d-flex flex-column align-items-center" style={{ maxWidth: "480px" }}>
//                 <div
//                   className="d-flex align-items-center justify-content-center rounded-circle mb-4"
//                   style={{ width: "64px", height: "64px", backgroundColor: "#fff2f0", border: "2px solid #ffccc7" }}
//                 >
//                   <WarningOutlined style={{ fontSize: "28px", color: "#ff4d4f" }} />
//                 </div>
//                 <h5 className="fw-bold text-dark mb-2">Preview could not be loaded</h5>
//                 <p className="text-muted mb-1" style={{ fontSize: "14px" }}>{loadError}</p>
//                 <Button
//                   type="primary"
//                   icon={<ReloadOutlined />}
//                   size="large"
//                   onClick={handleRetry}
//                   className="mt-4"
//                   style={{ backgroundColor: "#2b5aee", borderRadius: "8px", fontWeight: 500 }}
//                 >
//                   Retry Preview
//                 </Button>
//               </div>
//             </div>
//           ) : (
//             <div
//               className="shadow-sm bg-white rounded"
//               style={{
//                 maxWidth: "1000px",
//                 width: "100%",
//                 minHeight: "100%",
//                 padding: "50px",
//                 overflowX: "auto"
//               }}
//             >
//               <div ref={docxContainerRef} />
//             </div>
//           )}
//         </div>

//         {/* CRITICAL CHANGE 3: Target the div so the button knows what container to scroll back to top */}
//         <FloatButton.BackTop
//           target={() => document.getElementById("scrollable-document-container") || window}
//           visibilityHeight={300}
//           type="primary"
//           shape="circle"
//           style={{ right: 380, bottom: 40, zIndex: 9999, width: "50px", height: "50px" }}
//         />
//       </div>

//       {/* ── COLUMN 2: CONTROL SIDEBAR (RIGHT - FIXED) ── */}
//       <div
//         className="bg-white shadow-lg d-flex flex-column"
//         /* CRITICAL CHANGE 4: No scroll styles applied here. It naturally takes 100% of the available calc height */
//         style={{ width: "360px", minWidth: "360px", zIndex: 10 }}
//       >
//         <div className="p-4">
//           <Title level={4} className="mb-4" style={{ color: "#1f1f1f" }}>
//             Document Details
//           </Title>

//           {/* Rename Field */}
//           <div className="mb-4">
//             <Text type="secondary" className="fw-bold d-block mb-2" style={{ fontSize: "12px", letterSpacing: "0.5px" }}>
//               FILE NAME
//             </Text>
//             <Input
//               size="large"
//               value={docName}
//               onChange={(e) => setDocName(e.target.value)}
//               suffix={<EditOutlined className="text-muted" />}
//               style={{ fontWeight: 500 }}
//             />
//           </div>

//           {/* Cost Metrics */}
//           {costMetrics && costMetrics.total_cost_myr !== undefined && (
//             <div className="mb-4 p-3 bg-light rounded border">
//               <Text type="secondary" className="fw-bold d-block mb-1" style={{ fontSize: "12px", letterSpacing: "0.5px" }}>
//                 ESTIMATED GENERATION COST
//               </Text>
//               <Text className="fs-4 fw-bold text-success">
//                 RM {Number(costMetrics.total_cost_myr).toFixed(2)}
//               </Text>
//             </div>
//           )}
//         </div>

//         {/* Action Buttons - 'mt-auto' forces this block to the very bottom of the sidebar */}
//         <div className="p-4 mt-auto border-top bg-white">
//           <Button
//             type="primary"
//             size="large"
//             block
//             icon={<FileWordOutlined className="me-2 text-white" />}
//             onClick={handleDownloadDocx}
//             disabled={isLoading}
//             className="mb-3 fw-medium text-white"
//             style={{ height: "50px", backgroundColor: "#2b5aee" }}
//           >
//             Download DOCX
//           </Button>

//           <Button
//             danger
//             block
//             size="large"
//             icon={<ReloadOutlined className="me-2" />}
//             onClick={handleStartOver}
//             className="fw-medium"
//             style={{ height: "50px" }}
//           >
//             Start Over
//           </Button>
//         </div>
//       </div>
//     </div>
//   );
// }










"use client";

import { useEffect, useState, useRef, useCallback } from "react";
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
import { useRouter } from "next/navigation";
import { renderAsync } from "docx-preview";

const { Text, Title } = Typography;

export default function ReviewPage() {
  const router = useRouter();

  const [taskId, setTaskId] = useState<string | null>(null);
  const [docName, setDocName] = useState("Generated_Design_Document");
  const [costMetrics, setCostMetrics] = useState<any>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [loadError, setLoadError] = useState("");
  const [isClient, setIsClient] = useState(false);

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

      // Fetch Cost Metrics
      const statusRes = await fetch(`${API_BASE_URL}/api/status/${tid}`);
      if (statusRes.ok) {
        const statusData = await statusRes.json();
        if (statusData.cost_metrics) setCostMetrics(statusData.cost_metrics);
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

  const handleRetry = () => {
    if (!taskId) return;
    fetchDocxData(taskId);
  };

  useEffect(() => {
    setIsClient(true);
    const storedTaskId = sessionStorage.getItem("documentTaskId");
    const storedPreviewText = sessionStorage.getItem("documentPreviewData");

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
      fetchDocxData(storedTaskId); // Updated function name
    } else {
      router.push("/");
    }
  }, [router, fetchDocxData]);

  const handleDownloadDocx = () => {
    if (!taskId) return;
    const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";
    window.open(
      `${API_BASE_URL}/api/download/${taskId}?filename=${encodeURIComponent(docName)}`,
      "_blank"
    );
  };

  const handleStartOver = () => {
    Modal.confirm({
      title: "Start Over",
      icon: <ExclamationCircleOutlined />,
      content: "Are you sure you want to start over? Any generated files that you haven’t downloaded yet will be lost.",
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
          style={{ right: 400, bottom: 40, zIndex: 9999, width: "50px", height: "50px" }}
        />
      </div>

      {/* ── COLUMN 2: CONTROL SIDEBAR (RIGHT - FIXED) ── */}
      <div
        className="bg-white shadow-lg d-flex flex-column"
        style={{
          width: "360px",
          minWidth: "360px",
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
              FILE NAME
            </Text>
            <Input
              size="large"
              value={docName}
              onChange={(e) => setDocName(e.target.value)}
              suffix={<EditOutlined className="text-muted" />}
              style={{ fontWeight: 500 }}
            />
          </div>

          {/* Cost Metrics */}
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
        </div>

        {/* Action Buttons */}
        <div className="p-4 mt-auto border-top bg-white">
          <Button
            type="primary"
            size="large"
            block
            icon={<FileWordOutlined className="me-2 text-white" />}
            onClick={handleDownloadDocx}
            disabled={isLoading}
            className="mb-3 fw-medium text-white"
            style={{ height: "50px", backgroundColor: "#2b5aee" }}
          >
            Download DOCX
          </Button>

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
        </div>
      </div>
    </div>
  );
}