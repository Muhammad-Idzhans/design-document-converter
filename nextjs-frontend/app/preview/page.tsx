"use client";

import { useEffect, useState } from "react";
import { Button, Modal, Spin } from "antd";
import { PlayCircleOutlined, AppstoreOutlined } from "@ant-design/icons";
import { useRouter } from "next/navigation";

export default function PreviewPage() {
  const router = useRouter();
  
  const [taskId, setTaskId] = useState<string | null>(null);
  const [previewData, setPreviewData] = useState<any>(null);
  const [isClient, setIsClient] = useState(false);
  
  // Modal State
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [selectedSlide, setSelectedSlide] = useState<any>(null);

  useEffect(() => {
    setIsClient(true);
    // Load data from session storage passed by Phase 1 Upload
    const storedTaskId = sessionStorage.getItem("documentTaskId");
    const storedPreview = sessionStorage.getItem("documentPreview");

    if (storedTaskId && storedPreview) {
      setTaskId(storedTaskId);
      try {
        setPreviewData(JSON.parse(storedPreview));
      } catch (e) {
        console.error("Failed to parse preview data");
      }
    }
  }, []);

  const handleStartProcessing = async () => {
    if (!taskId) return;
    
    // Hit the /api/process endpoint to trigger background AI processing
    try {
      const res = await fetch(`http://localhost:8000/api/process/${taskId}`, {
        method: "POST"
      });
      if (!res.ok) throw new Error("Failed to start processing");
      
      // Navigate to Step 3
      router.push("/processing");
    } catch (err) {
      alert("Failed to connect to backend to start processing.");
      console.error(err);
    }
  };

  const openSlideModal = (slide: any) => {
    setSelectedSlide(slide);
    setIsModalOpen(true);
  };

  if (!isClient) {
    return <div className="min-vh-100 d-flex justify-content-center align-items-center"><Spin size="large" /></div>;
  }

  if (!previewData) {
    return (
      <div className="min-vh-100 d-flex flex-column justify-content-center align-items-center bg-light">
        <h4>No slides to preview</h4>
        <p className="text-muted">Please go back to the upload page and submit a PowerPoint file.</p>
        <Button onClick={() => router.push("/")} type="primary" className="mt-3">Return to Upload</Button>
      </div>
    );
  }

  return (
    <div className="container-fluid min-vh-100 bg-light py-5 px-md-5">
      
      {/* Header Section */}
      <div className="d-flex flex-wrap justify-content-between align-items-end mb-4 border-bottom pb-4">
        <div>
          <h2 className="fw-bold mb-2">Slide Preview</h2>
          <p className="text-muted mb-0" style={{ fontSize: "15px" }}>
            File loaded: <span className="fw-semibold text-dark">{previewData.filename}</span> &bull; {previewData.total_slides} slides extracted &bull; {previewData.diagrams_detected} diagrams detected.
          </p>
        </div>
        
        <Button 
          type="primary" 
          size="large" 
          icon={<PlayCircleOutlined />} 
          style={{ backgroundColor: "#2b5aee", padding: "0 24px", height: "46px", borderRadius: "8px", fontWeight: "500", marginTop: "16px" }}
          onClick={handleStartProcessing}
        >
          Start AI Processing
        </Button>
      </div>

      {/* Grid Section */}
      <div className="row row-cols-1 row-cols-sm-2 row-cols-md-3 row-cols-lg-4 g-4 pb-5">
        {previewData.slides && previewData.slides.map((slide: any, idx: number) => (
          <div className="col" key={idx}>
            <div 
              className="card h-100 shadow-sm border-0 custom-slide-card" 
              style={{ cursor: "pointer", transition: "transform 0.2s, box-shadow 0.2s", borderRadius: "10px", overflow: "hidden" }}
              onClick={() => openSlideModal(slide)}
            >
              
              {/* Slide Image Thumbnail */}
              <div 
                className="bg-light d-flex flex-column justify-content-center align-items-center position-relative" 
                style={{ height: "180px", overflow: "hidden" }}
              >
                {slide.thumbnail_url ? (
                   <img 
                     src={`http://localhost:8000${slide.thumbnail_url}`} 
                     alt={`Slide ${slide.slide_number}`} 
                     style={{ width: "100%", height: "100%", objectFit: "cover", objectPosition: "top center" }} 
                   />
                ) : (
                   <AppstoreOutlined className="text-secondary opacity-25" style={{ fontSize: "40px" }} />
                )}
                
                <div 
                  className="position-absolute top-0 start-0 m-3 px-2 py-1 bg-white rounded shadow-sm"
                  style={{ fontSize: "11px", fontWeight: "700", opacity: 0.9 }}
                >
                  SLIDE {slide.slide_number}
                </div>
              </div>

              {/* Presenter Notes */}
              <div className="card-body p-3 bg-white d-flex flex-column" style={{ minHeight: "80px", borderTop: "1px solid #f0f0f0" }}>
                <span className="text-muted" style={{ fontSize: "13px", display: '-webkit-box', WebkitLineClamp: 3, WebkitBoxOrient: 'vertical', overflow: 'hidden' }}>
                  {slide.notes || "No presenter note"}
                </span>
              </div>

            </div>
          </div>
        ))}
      </div>

      {/* Enlarge Slide Modal */}
      <Modal 
        title={<span className="fw-bold">Slide {selectedSlide?.slide_number} Preview</span>}
        open={isModalOpen} 
        onCancel={() => setIsModalOpen(false)}
        footer={[
          <Button key="close" onClick={() => setIsModalOpen(false)}>
            Close
          </Button>
        ]}
        width={800}
        centered
      >
        {selectedSlide && (
          <div className="py-3">
            <div className="bg-light rounded d-flex align-items-center justify-content-center border overflow-hidden" style={{ height: "400px", marginBottom: "20px" }}>
              {selectedSlide.thumbnail_url ? (
                   <img 
                     src={`http://localhost:8000${selectedSlide.thumbnail_url}`} 
                     alt={`Slide ${selectedSlide.slide_number}`} 
                     style={{ width: "100%", height: "100%", objectFit: "contain" }} 
                   />
                ) : (
                  <div className="text-center text-muted">
                    <h5>{selectedSlide.title || `Slide ${selectedSlide.slide_number}`}</h5>
                    <p>(Visual slide image placeholder)</p>
                  </div>
                )}
            </div>
            
            <h6 className="fw-bold text-dark">Extracted Presenter Notes:</h6>
            <div className="bg-secondary bg-opacity-10 p-3 rounded text-dark" style={{ minHeight: "80px", whiteSpace: "pre-wrap", fontSize: "14px" }}>
              {selectedSlide.notes || "No presenter note"}
            </div>
          </div>
        )}
      </Modal>

      {/* Minimal CSS for hover states added dynamically or via globals */}
      <style jsx global>{`
        .custom-slide-card:hover {
          transform: translateY(-4px);
          box-shadow: 0 10px 20px rgba(0,0,0,0.08) !important;
        }
      `}</style>
    </div>
  );
}
