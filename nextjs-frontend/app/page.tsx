"use client";
import { useState, useEffect } from "react";
import { Upload, Button, theme } from "antd";
import { CloudUploadOutlined, PictureOutlined, CheckCircleOutlined } from "@ant-design/icons";
import { useRouter } from "next/navigation";

const { Dragger } = Upload;

export default function UploadPage() {
  const router = useRouter();
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [selectedLogo, setSelectedLogo] = useState<File | null>(null);
  const [logoPreviewUrl, setLogoPreviewUrl] = useState<string | null>(null);
  const [isUploading, setIsUploading] = useState(false);

  // Manage object URL memory safely
  useEffect(() => {
    if (selectedLogo) {
      const objectUrl = URL.createObjectURL(selectedLogo);
      setLogoPreviewUrl(objectUrl);
      return () => URL.revokeObjectURL(objectUrl);
    } else {
      setLogoPreviewUrl(null);
    }
  }, [selectedLogo]);

  return (
    <div className="container-fluid min-vh-100 bg-light d-flex flex-column align-items-center py-5">

      {/* Header Section */}
      <div className="text-center mt-4 mb-5">
        <h2 className="fw-bold mb-2">Generate Design Document</h2>
        <p className="text-muted mb-0">Upload your PowerPoint architecture deck to begin.</p>
      </div>

      <div className="w-100" style={{ maxWidth: "720px" }}>

        {/* Drag & Drop Section */}
        <Dragger
          name="file"
          multiple={false}
          accept=".pptx"
          showUploadList={false}
          className="custom-dragger shadow-sm"
          beforeUpload={(file: any) => {
            setSelectedFile(file);
            return false; // Prevent automatic HTTP upload, we handle it manually
          }}
          style={{ padding: selectedFile ? "50px 0" : "60px 0" }}
        >
          <div className="d-flex flex-column align-items-center">
            {selectedFile ? (
              <>
                <div
                  className="upload-icon d-flex align-items-center justify-content-center rounded-circle text-success mx-auto mb-3"
                  style={{ width: "64px", height: "64px", fontSize: "32px", backgroundColor: "#e6ffed" }}
                >
                  <CheckCircleOutlined />
                </div>
                <h5 className="fw-semibold mb-1 text-success"><span className="fw-bold fst-italic">{selectedFile.name}</span><br />has been uploaded</h5>
                <p className="text-muted small mb-4">Would you like to proceed to the next step?</p>

                <div className="d-flex gap-3 mt-2">
                  <Button size="large" onClick={(e: any) => {
                    // Letting this bubble up re-opens the Dragger file dialog
                  }}>
                    Select another .pptx file
                  </Button>
                  <Button
                    type="primary"
                    size="large"
                    onClick={async (e: any) => {
                      e.stopPropagation(); // Stop Dragger upload dialog from opening
                      if (!selectedFile) return;
                      
                      setIsUploading(true);
                      
                      try {
                        const formData = new FormData();
                        formData.append("file", selectedFile);
                        if (selectedLogo) {
                          formData.append("logo", selectedLogo);
                        }
                        
                        const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";
                        const res = await fetch(`${API_BASE_URL}/api/upload`, {
                          method: "POST",
                          body: formData
                        });
                        
                        if (!res.ok) throw new Error("Server responded with status " + res.status);
                        
                        const data = await res.json();
                        if (data.error) throw new Error(data.error);
                        
                        if (typeof window !== "undefined") {
                          sessionStorage.setItem("documentTaskId", data.task_id);
                          sessionStorage.setItem("documentPreview", JSON.stringify(data.preview));
                        }
                        
                        router.push('/preview');
                      } catch (err: any) {
                        console.error(err);
                        alert("Failed to upload to backend: " + err.message);
                      } finally {
                        setIsUploading(false);
                      }
                    }}
                    loading={isUploading}
                  >
                    Next: Preview Slides
                  </Button>
                </div>
              </>
            ) : (
              <>
                <div
                  className="upload-icon d-flex align-items-center justify-content-center rounded-circle text-primary mx-auto mb-3"
                  style={{ width: "64px", height: "64px", fontSize: "32px", backgroundColor: "#e6f0ff" }}
                >
                  <CloudUploadOutlined />
                </div>
                <h5 className="fw-semibold mb-1">Drag & drop your PPTX file here</h5>
                <p className="text-muted small mb-4">Support for .pptx files only</p>

                <Button type="primary" size="large" onClick={(e: any) => e.preventDefault()}>
                  Browse Files
                </Button>
              </>
            )}
          </div>
        </Dragger>

        {/* Uploaded Client Logo Section */}
        <div className="bg-white border rounded p-4 mt-4 shadow-sm">
          <div className="d-flex align-items-center justify-content-between">
            <div className="d-flex align-items-center">
              <div
                className="bg-light border rounded d-flex align-items-center justify-content-center me-3"
                style={{ width: "48px", height: "48px", fontSize: "22px" }}
              >
                <PictureOutlined className="text-secondary" />
              </div>
              <div>
                <h6 className="fw-semibold mb-1">Upload Client Logo</h6>
                <p className="text-muted small mb-0">Optional &bull; png, jpeg, jpg, svg+xml</p>
              </div>
            </div>

            <Upload
              name="logo"
              accept="image/png, image/jpeg, image/jpg, image/svg+xml"
              showUploadList={false}
              beforeUpload={(file: any) => {
                setSelectedLogo(file);
                return false;
              }}
            >
              <Button size="middle">{selectedLogo ? "Select another logo" : "Select"}</Button>
            </Upload>
          </div>

          {logoPreviewUrl && selectedLogo && (
            <div className="mt-4 pt-4 border-top d-flex flex-column align-items-center" style={{ animation: "fadeIn 0.3s ease-in-out" }}>
              <p className="text-muted small mb-3 fw-semibold text-uppercase" style={{ letterSpacing: "0.5px" }}>Logo Preview</p>
              <img
                src={logoPreviewUrl}
                alt="Client Logo Preview"
                style={{ maxHeight: "140px", maxWidth: "100%", objectFit: "contain" }}
                className="border rounded p-3 bg-light mb-3 shadow-sm"
              />
              <span className="text-secondary small fw-medium">{selectedLogo.name}</span>
            </div>
          )}
        </div>

      </div>
    </div>
  );
}
