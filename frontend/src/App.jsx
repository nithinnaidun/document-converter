import { useRef, useState } from "react";
import "./App.css";

const API_URL = "http://127.0.0.1:8000";

const conversions = {
  "pdf-to-word": {
    label: "PDF → Word",
    output: ".docx",
    icon: "📄",
    extensions: [".pdf"],
    endpoint: "pdf-to-word",
    description: "Convert PDF into editable Word",
  },

  "pdf-to-excel": {
    label: "PDF → Excel",
    output: ".xlsx",
    icon: "📊",
    extensions: [".pdf"],
    endpoint: "pdf-to-excel",
    description: "Extract tables into Excel",
  },

  "pdf-to-jpg": {
    label: "PDF → JPG",
    output: ".jpg",
    icon: "🖼️",
    extensions: [".pdf"],
    endpoint: "pdf-to-jpg",
    description: "Convert PDF pages to images",
  },

  "pdf-to-txt": {
    label: "PDF → TXT",
    output: ".txt",
    icon: "📝",
    extensions: [".pdf"],
    endpoint: "pdf-to-txt",
    description: "Extract text from PDF",
  },

  "word-to-pdf": {
    label: "Word → PDF",
    output: ".pdf",
    icon: "📘",
    extensions: [".doc", ".docx"],
    endpoint: "word-to-pdf",
    description: "Convert Word document to PDF",
  },

  "word-to-txt": {
    label: "Word → TXT",
    output: ".txt",
    icon: "📝",
    extensions: [".doc", ".docx"],
    endpoint: "word-to-txt",
    description: "Extract Word text",
  },

  "excel-to-pdf": {
    label: "Excel → PDF",
    output: ".pdf",
    icon: "📗",
    extensions: [".xls", ".xlsx"],
    endpoint: "excel-to-pdf",
    description: "Convert spreadsheet to PDF",
  },

  "excel-to-xlsx": {
    label: "Excel → XLSX",
    output: ".xlsx",
    icon: "📊",
    extensions: [".xls", ".xlsx"],
    endpoint: "excel-to-xlsx",
    description: "Clean and convert workbook",
  },

  "excel-to-csv": {
    label: "Excel → CSV",
    output: ".csv",
    icon: "📋",
    extensions: [".xls", ".xlsx"],
    endpoint: "excel-to-csv",
    description: "Convert spreadsheet to CSV",
  },

  "ppt-to-pdf": {
    label: "PowerPoint → PDF",
    output: ".pdf",
    icon: "📙",
    extensions: [".ppt", ".pptx"],
    endpoint: "ppt-to-pdf",
    description: "Convert presentation to PDF",
  },

  "ppt-to-pptx": {
    label: "PowerPoint → PPTX",
    output: ".pptx",
    icon: "📙",
    extensions: [".ppt", ".pptx"],
    endpoint: "ppt-to-pptx",
    description: "Convert presentation to PPTX",
  },

  "image-to-pdf": {
    label: "JPG/PNG → PDF",
    output: ".pdf",
    icon: "🖼️",
    extensions: [".jpg", ".jpeg", ".png"],
    endpoint: "image-to-pdf",
    description: "Convert images to PDF",
  },

  "image-to-word": {
    label: "JPG/PNG → Word",
    output: ".docx",
    icon: "🖼️",
    extensions: [".jpg", ".jpeg", ".png"],
    endpoint: "image-to-word",
    description: "Convert image to editable Word",
  },

  "image-to-excel": {
    label: "Image → Excel",
    output: ".xlsx",
    icon: "📊",
    extensions: [".jpg", ".jpeg", ".png"],
    endpoint: "image-to-excel",
    description: "OCR and extract table data",
  },

  "txt-to-pdf": {
    label: "TXT → PDF",
    output: ".pdf",
    icon: "📝",
    extensions: [".txt"],
    endpoint: "txt-to-pdf",
    description: "Convert text to PDF",
  },

  "txt-to-word": {
    label: "TXT → Word",
    output: ".docx",
    icon: "📝",
    extensions: [".txt"],
    endpoint: "txt-to-word",
    description: "Convert text to Word",
  },

  "bank-to-excel": {
    label: "Bank Statement → Excel",
    output: ".xlsx",
    icon: "🏦",
    extensions: [".pdf", ".jpg", ".jpeg", ".png"],
    endpoint: "bank-to-excel",
    description: "Extract transactions into clean Excel",
  },

  "bank-to-csv": {
    label: "Bank Statement → CSV",
    output: ".csv",
    icon: "🏦",
    extensions: [".pdf", ".jpg", ".jpeg", ".png"],
    endpoint: "bank-to-csv",
    description: "Extract transactions into CSV",
  },

  "bank-to-pdf": {
    label: "Bank Statement → PDF",
    output: ".pdf",
    icon: "🏦",
    extensions: [".pdf", ".jpg", ".jpeg", ".png"],
    endpoint: "bank-to-pdf",
    description: "Create clean PDF statement",
  },

  "invoice-to-excel": {
    label: "Invoice → Excel",
    output: ".xlsx",
    icon: "🧾",
    extensions: [".pdf", ".jpg", ".jpeg", ".png"],
    endpoint: "invoice-to-excel",
    description: "Extract invoice data into Excel",
  },

  "invoice-to-word": {
    label: "Invoice → Word",
    output: ".docx",
    icon: "🧾",
    extensions: [".pdf", ".jpg", ".jpeg", ".png"],
    endpoint: "invoice-to-word",
    description: "Convert invoice to Word",
  },

  "invoice-to-pdf": {
    label: "Invoice → PDF",
    output: ".pdf",
    icon: "🧾",
    extensions: [".pdf", ".jpg", ".jpeg", ".png"],
    endpoint: "invoice-to-pdf",
    description: "Create clean PDF invoice",
  },

  "receipt-to-excel": {
    label: "Receipt → Excel",
    output: ".xlsx",
    icon: "🧾",
    extensions: [".pdf", ".jpg", ".jpeg", ".png"],
    endpoint: "receipt-to-excel",
    description: "Extract receipt data",
  },

  "receipt-to-pdf": {
    label: "Receipt → PDF",
    output: ".pdf",
    icon: "🧾",
    extensions: [".pdf", ".jpg", ".jpeg", ".png"],
    endpoint: "receipt-to-pdf",
    description: "Create PDF from receipt",
  },

  "form-to-excel": {
    label: "Form → Excel",
    output: ".xlsx",
    icon: "📋",
    extensions: [".pdf", ".jpg", ".jpeg", ".png"],
    endpoint: "form-to-excel",
    description: "Extract form fields",
  },

  "form-to-word": {
    label: "Form → Word",
    output: ".docx",
    icon: "📋",
    extensions: [".pdf", ".jpg", ".jpeg", ".png"],
    endpoint: "form-to-word",
    description: "Convert form to editable Word",
  },
};

const allowedTypes = [
  ".pdf",
  ".doc",
  ".docx",
  ".xls",
  ".xlsx",
  ".ppt",
  ".pptx",
  ".jpg",
  ".jpeg",
  ".png",
  ".txt",
];

function App() {
  const fileInputRef = useRef(null);

  const [file, setFile] = useState(null);
  const [format, setFormat] = useState("");
  const [dragging, setDragging] = useState(false);
  const [status, setStatus] = useState("");
  const [converting, setConverting] = useState(false);
  const [progress, setProgress] = useState(0);
  const [downloadUrl, setDownloadUrl] = useState("");
  const [downloadName, setDownloadName] = useState("");

  const getExtension = (name) => {
    return "." + name.split(".").pop().toLowerCase();
  };

  const handleFile = (selectedFile) => {
    if (!selectedFile) return;

    const extension = getExtension(selectedFile.name);

    if (!allowedTypes.includes(extension)) {
      setStatus("❌ This file type is not supported.");
      return;
    }

    if (selectedFile.size > 25 * 1024 * 1024) {
      setStatus("❌ File size must be less than 25 MB.");
      return;
    }

    setFile(selectedFile);
    setStatus("");
    setDownloadUrl("");
    setDownloadName("");

    const compatible = Object.entries(conversions).filter(
      ([, item]) => item.extensions.includes(extension)
    );

    if (compatible.length > 0) {
      setFormat(compatible[0][0]);
    }
  };

  const handleFileChange = (e) => {
    handleFile(e.target.files[0]);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setDragging(false);
    handleFile(e.dataTransfer.files[0]);
  };

  const chooseConversion = (key) => {
    setFormat(key);
    setStatus("");

    setTimeout(() => {
      document
        .getElementById("converter")
        ?.scrollIntoView({
          behavior: "smooth",
          block: "center",
        });
    }, 50);
  };

  const uploadAndConvert = async () => {
    if (!file) {
      setStatus("❌ Please upload a document first.");
      return;
    }

    if (!format) {
      setStatus("❌ Please choose a conversion.");
      return;
    }

    const conversion = conversions[format];
    const extension = getExtension(file.name);

    if (!conversion.extensions.includes(extension)) {
      setStatus(
        `❌ ${conversion.label} does not support ${extension.toUpperCase()} files.`
      );
      return;
    }

    try {
      setConverting(true);
      setProgress(10);
      setDownloadUrl("");

      setStatus("⏳ Uploading document...");

      const formData = new FormData();
      formData.append("file", file);

      const uploadResponse = await fetch(
        `${API_URL}/api/upload`,
        {
          method: "POST",
          body: formData,
        }
      );

      if (!uploadResponse.ok) {
        throw new Error("Upload failed");
      }

      const uploadData = await uploadResponse.json();

      setProgress(35);
      setStatus(`🔍 Analyzing ${file.name}...`);

      await new Promise((resolve) =>
        setTimeout(resolve, 400)
      );

      setProgress(55);
      setStatus(
        `⚙️ Converting to ${conversion.label}...`
      );

      const convertResponse = await fetch(
        `${API_URL}/api/convert/${conversion.endpoint}/${uploadData.file_id}`,
        {
          method: "POST",
        }
      );

      if (!convertResponse.ok) {
        let message = "Conversion failed";

        try {
          const errorData =
            await convertResponse.json();

          message =
            errorData.detail || message;
        } catch {}

        throw new Error(message);
      }

      const convertData =
        await convertResponse.json();

      setProgress(80);
      setStatus("📦 Preparing your converted file...");

      const downloadResponse = await fetch(
        `${API_URL}/api/download/${convertData.file_id}`
      );

      if (!downloadResponse.ok) {
        throw new Error("Download failed");
      }

      const blob =
        await downloadResponse.blob();

      const blobUrl =
        URL.createObjectURL(blob);

      const outputName =
        convertData.filename ||
        `converted-document${conversion.output}`;

      setDownloadUrl(blobUrl);
      setDownloadName(outputName);
      setProgress(100);

      const link =
        document.createElement("a");

      link.href = blobUrl;
      link.download = outputName;

      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      setStatus(
        "✅ Conversion completed successfully!"
      );

      setConverting(false);

    } catch (error) {
      console.error(error);

      setProgress(0);

      setStatus(
        `❌ ${error.message || "Something went wrong."}`
      );

      setConverting(false);
    }
  };

  const removeFile = () => {
    setFile(null);
    setFormat("");
    setStatus("");
    setProgress(0);
    setDownloadUrl("");
    setDownloadName("");

    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  return (
    <div className="app">

      {/* Animated background */}

      <div className="background-orb orb-one"></div>
      <div className="background-orb orb-two"></div>
      <div className="background-orb orb-three"></div>


      {/* HEADER */}

      <header className="header">

        <div className="logo">
          <span className="logo-document">
            📄
          </span>

          Docu<span>Convert</span>
        </div>

        <nav>
          <a href="#conversion-types">
            Conversions
          </a>

          <a href="#converter">
            Converter
          </a>
        </nav>

      </header>


      {/* HERO */}

      <main>

        <section className="hero-section">

          <div className="floating-document document-one">
            📄
          </div>

          <div className="floating-document document-two">
            📊
          </div>

          <div className="floating-document document-three">
            🧾
          </div>


          <div className="badge">
            ✨ FREE DOCUMENT CONVERTER
          </div>

          <h1>
            Convert your documents
            <br />

            <span>
              quickly and accurately
            </span>
          </h1>

          <p className="subtitle">
            Convert PDF, Word, Excel,
            PowerPoint, images, bank
            statements, invoices and more.
            <br />

            No registration. No login.
            No permanent storage.
          </p>

        </section>


        {/* CONVERSION TOOLS */}

        <section
          id="conversion-types"
          className="conversion-types"
        >

          <div className="section-heading">

            <span className="section-badge">
              CONVERSION TOOLS
            </span>

            <h2>
              Choose what you want to convert
            </h2>

            <p>
              Select a conversion type, then
              upload your document.
            </p>

          </div>


          <div className="conversion-grid">

            {Object.entries(conversions).map(
              ([key, conversion]) => (

                <button
                  key={key}
                  type="button"
                  className={
                    `conversion-card ${
                      format === key
                        ? "selected"
                        : ""
                    }`
                  }
                  onClick={() =>
                    chooseConversion(key)
                  }
                >

                  <div className="conversion-icon">
                    {conversion.icon}
                  </div>

                  <div className="conversion-content">

                    <strong>
                      {conversion.label}
                    </strong>

                    <span>
                      {conversion.description}
                    </span>

                  </div>

                  {format === key && (
                    <div className="selected-mark">
                      ✓
                    </div>
                  )}

                </button>

              )
            )}

          </div>

        </section>


        {/* CONVERTER */}

        <section
          id="converter"
          className="converter-card"
        >

          <div className="converter-title">

            <div className="step-circle">
              1
            </div>

            <div>

              <span>
                UPLOAD
              </span>

              <h2>
                Upload your document
              </h2>

              <p>
                Drag & drop or click to browse
              </p>

            </div>

          </div>


          {/* UPLOAD AREA */}

          <div
            className={
              `drop-zone ${
                dragging ? "dragging" : ""
              } ${file ? "has-file" : ""}`
            }

            onDragOver={(e) => {
              e.preventDefault();
              setDragging(true);
            }}

            onDragLeave={() =>
              setDragging(false)
            }

            onDrop={handleDrop}

            onClick={() =>
              fileInputRef.current?.click()
            }
          >

            <input
              ref={fileInputRef}
              type="file"
              hidden
              accept=".pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.jpg,.jpeg,.png,.txt"
              onChange={handleFileChange}
            />

            {!file ? (
              <>

                <div className="upload-3d">
                  <div className="upload-cube">
                    <span>📄</span>
                  </div>
                </div>

                <h2>
                  Drop your document here
                </h2>

                <p>
                  or click to browse from your computer
                </p>

                <small>
                  PDF • DOC • DOCX • XLS • XLSX •
                  PPT • PPTX • JPG • PNG • TXT
                  <br />
                  Maximum file size: 25 MB
                </small>

              </>
            ) : (

              <div className="selected-file-large">

                <div className="large-file-icon">
                  📄
                </div>

                <div>

                  <h3>
                    {file.name}
                  </h3>

                  <p>
                    {(file.size / 1024 / 1024).toFixed(2)}
                    {" "}MB
                  </p>

                  <span>
                    File ready for conversion
                  </span>

                </div>

              </div>

            )}

          </div>


          {/* FILE CONTROLS */}

          {file && (
            <div className="file-control">

              <div className="mini-file">
                📄
              </div>

              <div className="file-control-info">

                <strong>
                  {file.name}
                </strong>

                <span>
                  {getExtension(file.name).toUpperCase()}
                  {" "}•{" "}
                  {(file.size / 1024 / 1024).toFixed(2)}
                  {" "}MB
                </span>

              </div>

              <button
                type="button"
                className="remove-button"
                onClick={(e) => {
                  e.stopPropagation();
                  removeFile();
                }}
              >
                ✕
              </button>

            </div>
          )}


          {/* SELECTED CONVERSION */}

          {format &&
            conversions[format] && (

              <div className="selected-conversion">

                <div className="selected-left">

                  <div className="selected-conversion-icon">
                    {conversions[format].icon}
                  </div>

                  <div>

                    <small>
                      OUTPUT FORMAT
                    </small>

                    <strong>
                      {conversions[format].label}
                    </strong>

                  </div>

                </div>

                <button
                  type="button"
                  onClick={() =>
                    document
                      .getElementById(
                        "conversion-types"
                      )
                      ?.scrollIntoView({
                        behavior: "smooth",
                      })
                  }
                >
                  Change
                </button>

              </div>
            )}


          {/* PROGRESS */}

          {converting && (
            <div className="progress-container">

              <div className="progress-top">

                <span>
                  Processing document
                </span>

                <strong>
                  {progress}%
                </strong>

              </div>

              <div className="progress-track">

                <div
                  className="progress-bar"
                  style={{
                    width: `${progress}%`,
                  }}
                ></div>

              </div>

            </div>
          )}


          {/* CONVERT BUTTON */}

          <button
            type="button"
            className="convert-button"
            onClick={uploadAndConvert}
            disabled={
              !file ||
              !format ||
              converting
            }
          >

            <span className="button-icon">
              {converting ? "⏳" : "⚡"}
            </span>

            {converting
              ? "Processing Document..."
              : "Convert Document"}

            <span className="button-arrow">
              →
            </span>

          </button>


          {/* STATUS */}

          {status && (
            <div
              className={
                `status ${
                  status.startsWith("❌")
                    ? "error"
                    : status.startsWith("✅")
                    ? "success"
                    : ""
                }`
              }
            >
              {status}
            </div>
          )}


          {/* DOWNLOAD */}

          {downloadUrl && (

            <div className="download-result">

              <div className="download-check">
                ✓
              </div>

              <div>

                <h3>
                  Your document is ready
                </h3>

                <p>
                  {downloadName}
                </p>

              </div>

              <a
                className="download-button"
                href={downloadUrl}
                download={downloadName}
              >
                ⬇ Download
              </a>

            </div>

          )}


          <div className="privacy-note">
            🔒 Temporary processing •
            Files automatically deleted
            after processing
          </div>

        </section>

      </main>


      {/* FOOTER */}

      <footer>

        <div className="footer-logo">
          📄 DocuConvert
        </div>

        <p>
          Developed by{" "}
          <strong>
            Nithin Naidu N
          </strong>
        </p>

        <small>
          © 2026 DocuConvert
        </small>

      </footer>

    </div>
  );
}

export default App;