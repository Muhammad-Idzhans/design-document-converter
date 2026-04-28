# 📄 Enfrasys Design Document Generator
An intelligent, enterprise-ready web application designed to automatically convert pre-sales architectural presentations (.pptx) into comprehensive, formal Microsoft Word Solution Design Documents (.docx). Powered by Microsoft Foundry (Agentic AI) and Azure AI Content Understanding.

## 📖 Overview
This project provides a Next.js-based frontend for the design document generator UI, and a FastAPI-based backend for the document generation process. In Enfrasys, most pre-sales staff and solutions architects will create a Design Deck first in `.pptx` format. This deck is presented to the customer first hand, before the proper, formal design document in `.docx` format is produced. This tool accelerates the process of drafting that final `.docx` file by intelligently extracting and analyzing the presentation content.

The solution leverages Microsoft AI Services such as the Microsoft Foundry Agentic solution, Azure OpenAI Vision, and Azure AI Content Understanding, together with open-source technologies (like python-pptx) to execute the document transformation with high fidelity.

## ✨ Key Features
- **Automated Document Generation:** Seamlessly converts uploaded `.pptx` presentations into fully formatted `.docx` design documents.
- **Interactive AI Context:** Users can manually override or add specific AI context (presenter notes) to individual slides, providing targeted guidance for the document generation.
- **Cost Estimation:** Transparently calculates and displays the estimated price for generating a specific document based on token usage and slide count.
- **Custom Document Naming:** Users can rename the generated design document before the final output is produced.
- **Agentic AI Orchestration:** Powered by Microsoft Foundry Agentic Architecture (Orchestrator and Writer agents) utilizing GPT-4.1 to dynamically construct a bespoke Table of Contents and generate technical content that maps directly to the provided slides.
- **Advanced Content Extraction:** Deep parsing of slide text, speaker notes, and tables using `python-pptx` and Azure AI Content Understanding.
- **Vision AI Analysis:** Automatic extraction of diagrams and architecture images, which are then analyzed and described by Azure OpenAI Vision (GPT-4o).
- **Modern Web UI:** A sleek, responsive Next.js frontend featuring drag-and-drop file uploads, real-time status polling, and interactive slide previews.

## 🏗️ Architecture
The system consists of a decoupled frontend and backend:
1. **Frontend (Next.js):** Handles the user interface, file uploads, progress polling, interactive slide previews, and context editing.
2. **Backend (FastAPI):** Exposes REST endpoints to manage background tasks for file processing. It orchestrates content extraction and communication with the Azure AI ecosystem (OpenAI Vision, Content Understanding, and Agentic endpoints) to assemble the final Word document.

## 💻 Tech Stack
- **Frontend:** [Next.js](https://nextjs.org/) (React Framework), Ant Design, Bootstrap, TailwindCSS
- **Backend API:** [FastAPI](https://fastapi.tiangolo.com/), Uvicorn
- **Document Processing:** `python-pptx`, `python-docx`, PyMuPDF (`fitz`), Pandoc
- **AI Orchestration:** Microsoft Foundry (Agentic AI Architecture)
- **Vision & Extraction:** Azure OpenAI (GPT-4o Vision), Azure AI Content Understanding
- **Deployment:** Docker, Azure Web App (Linux), [Vercel.com](https://vercel.com/)

## 📋 Prerequisites
Before setting up the project locally, ensure you have the following installed:
- [Node.js](https://nodejs.org/en/) (v18 or higher)
- [Python](https://www.python.org/) (v3.9 or higher)
- An active Azure Subscription with the necessary Microsoft Foundry (Azure OpenAI endpoints) and Azure AI Content Understanding resources provisioned.

## ⚙️ Environment Variables
### Backend (`python-backend-development/.env`)
Create a `.env` file in the backend directory and configure the following variables:
```env
# Azure OpenAI Vision Settings
AZURE_OPENAI_ENDPOINT="<your-vision-endpoint>"
AZURE_OPENAI_KEY="<your-vision-key>"
AZURE_OPENAI_DEPLOYMENT_NAME="<your-vision-deployment-name-e.g-gpt-4o>"

# Microsoft Foundry Agentic Settings
AGENT_OPENAI_ENDPOINT="<your-agent-endpoint>"
AGENT_OPENAI_KEY="<your-agent-key>"
AGENT_ASSISTANT_ID="<your-agent-assistant-id>"
ORCHESTRATOR_DEPLOYMENT="<your-orchestrator-deployment-name-e.g-gpt-4o>"
WRITER_DEPLOYMENT="<your-writer-deployment-name-e.g-gpt-4o>"

# Azure AI Content Understanding
CONTENT_UNDERSTANDING_ENDPOINT="<your-cu-endpoint>"
CONTENT_UNDERSTANDING_KEY="<your-cu-key>"
```

### Frontend (`nextjs-frontend/.env.local`)
Create a `.env.local` file in the frontend directory:
```env
NEXT_PUBLIC_API_URL="http://localhost:8000"
```

## 🚀 Local Setup

#### **1. Clone the repository:**
```cmd
git clone https://github.com/Muhammad-Idzhans/design-document-converter.git
cd design-document-converter
```

#### **2. Start the Backend (FastAPI):**
```cmd
cd python-backend-deployment
pip install -r requirements.txt
python app.py
```

#### **3. Start the Frontend (Next.js):**
Open a new terminal window:
```cmd
cd nextjs-frontend
npm install
npm run dev
```

#### **4. Access the application:**
Open [http://localhost:3000](http://localhost:3000) in your web browser.

## ☁️ **[Backend]** Deployment to Azure Web App
The backend is packaged as a Docker container and hosted on Azure Web App for reliable deployment.

#### **1. Build and Push the Docker Image:**
First, build the Docker image using the provided `Dockerfile` and push it to your public Docker Hub registry.
```cmd
cd python-backend-deployment
docker build -t <your-docker-username>/<image-name>:latest .
docker push <your-docker-username>/<image-name>:latest
```

#### **2. Create the Azure Web App**
In the Microsoft Azure Portal, create a new Web App. In the Basics tab, apply the following settings:
- **Publish**: Container
- **Operating System:** Linux
- **Region:** Choose a region (e.g., Southeast Asia)
- **Linux Plan:** P0v3 (Recommended)

#### **3. Configure Container Settings**
Next, navigate to the Container tab during the setup process and configure it to pull your image from Docker Hub:
- **Image Source:** Other Container Registries
- **Access Type:** Public
- **Registry server URL:** https://index.docker.io
- **Image and tag:** <your-docker-username>/<image-name>:latest
- **Port:** 8000

#### **4. Configure Environment Variables:**
Navigate to **Settings -> Environment variables** in your Web App. 
First, add all the standard variables from your local `.env` file (e.g., endpoints, keys). 
Next, add the following Azure-specific App Service variables to ensure the container routes correctly and saves files persistently:
- `WEBSITES_PORT` = `8000`
- `WEBSITES_ENABLE_APP_SERVICE_STORAGE` = `true`
- `OUTPUT_DIR` = `/home/site/wwwroot/outputs`

#### **5. Enable "Always On":**
To prevent the container from spinning down and causing cold-start delays:
- Navigate to **Settings -> Configuration** (or **General settings**).
- Locate the **Always On** toggle and set it to **On**.
- Click **Save**.

#### **6. Restart and Verify:**
Navigate to the Web App's **Overview** page and click **Restart**. Wait a few moments, then check the Log Stream or navigate to the provided Azure URL to ensure the container has successfully started and is ready to process documents.

## 🖱️ Usage Guide

**1. Upload Presentation**
![Upload PPTX file to the UI](media/designDocumentConverter-upload.png)
- Navigate to the frontend application.
- Drag and drop your `.pptx` design deck into the upload area or click the **"Browse Files"** button to select a file. 
- Optionally, upload a client logo to be embedded in the final design document.
- Then, click the **"Next: Preview Slides"** button to continue.

**2. Processing & Extraction**
![Preview Processed slides](media/designDocumentConverter-preview.png)
![Edit and add AI context](media/designDocumentConverter-previewContext.png)
- Once uploaded, the backend will initiate a background process to extract images, content and keynotes from the PPTX file.
- The UI will display a preview of extracted data from each slides.
- Optionally, you may click on of the slides to add extra context to be included in the design document.
- Once you ready, you can click the **"Start AI Processing"** button at the top right corner to continue.

**3. Processing Begins**
![Processing in Progress](media/designDocumentConverter-loading.png)
- Once you click **"Start AI Processing"**, you will be redirected to a new page where you can monitor the processing status.
- Click **""Preview Generated Document** to view the processed document.


**4. Final Review & Generate**
![Preview Generated Document](media/designDocumentConverter-download.png)
- Review the extracted slide thumbnails and data.
- **Cost Estimation:** View the estimated price for generating the document based on your current configurations.
- **Rename:** Optionally rename the document to your preferred title.
- Download your fully formatted, bespoke Solution Design Document in `.docx` format.

---

<div align="center">
  <em>Developed for Enfrasys by Muhammad Idzhans Khairi</em>
</div>
