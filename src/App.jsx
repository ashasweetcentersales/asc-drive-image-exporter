
import React, { useEffect, useMemo, useRef, useState } from "react";
import ExcelJS from "exceljs";
import {
  ArrowDownToLine,
  Check,
  ChevronDown,
  CircleAlert,
  FileSpreadsheet,
  FolderOpen,
  Image as ImageIcon,
  Link2,
  LoaderCircle,
  LogIn,
  LogOut,
  Search,
  ShieldCheck,
  X,
  Copy,
  Files
} from "lucide-react";

const DRIVE_API = "https://www.googleapis.com/drive/v3";
const FOLDER_MIME = "application/vnd.google-apps.folder";
const IMAGE_MIME_PREFIX = "image/";

// --------------------------------------------------
// IMAGE PREVIEW
// --------------------------------------------------

function DriveImage({ file, token, className = "", alt = "" }) {
  const [src, setSrc] = useState("");
  const [failed, setFailed] = useState(false);

  useEffect(() => {
    let active = true;
    let objectUrl = "";

    setSrc("");
    setFailed(false);

    if (!token || !file?.id) return undefined;

    fetch(
      `${DRIVE_API}/files/${encodeURIComponent(file.id)}?alt=media`,
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    )
      .then((response) => {
        if (!response.ok) {
          throw new Error("Preview unavailable");
        }

        return response.blob();
      })
      .then((blob) => {
        if (!blob.type.startsWith("image/")) {
          throw new Error("Invalid image response");
        }

        objectUrl = URL.createObjectURL(blob);

        if (active) {
          setSrc(objectUrl);
        }
      })
      .catch(() => {
        if (active) setFailed(true);
      });

    return () => {
      active = false;

      if (objectUrl) {
        URL.revokeObjectURL(objectUrl);
      }
    };
  }, [file?.id, token]);

  if (src) {
    return (
      <img
        className={className}
        src={src}
        alt={alt}
        loading="lazy"
      />
    );
  }

  return (
    <span
      className={`thumb-fallback ${
        failed ? "preview-unavailable" : ""
      }`}
    >
      <ImageIcon size={26} />
    </span>
  );
}

// --------------------------------------------------
// GOOGLE DRIVE LINK EXTRACTION
// --------------------------------------------------

function extractDriveId(input) {
  const value = input.trim();

  if (/^[a-zA-Z0-9_-]{10,}$/.test(value)) {
    return value;
  }

  const patterns = [
    /\/folders\/([a-zA-Z0-9_-]+)/,
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /\/d\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/
  ];

  for (const pattern of patterns) {
    const match = value.match(pattern);

    if (match) return match[1];
  }

  throw new Error(
    "Enter a valid Google Drive folder URL or image URL."
  );
}

function extension(name, mimeType) {
  const dot = name.lastIndexOf(".");

  if (dot > -1 && dot < name.length - 1) {
    return name.slice(dot + 1).toUpperCase();
  }

  return (
    ((mimeType || "").split("/")[1] || "IMAGE")
      .split("+")[0]
      .toUpperCase()
  );
}

function getDriveLinks(file) {
  return {
    copyLink:
      file.webViewLink ||
      `https://drive.google.com/file/d/${file.id}/view`,

    directLink:
      `https://lh3.googleusercontent.com/d/${file.id}`
  };
}

// --------------------------------------------------
// DRIVE API: GET FILE METADATA
// --------------------------------------------------

async function getDriveMetadata(token, fileId) {
  const params = new URLSearchParams({
    fields:
      "id,name,mimeType,webViewLink,size,md5Checksum,modifiedTime",
    supportsAllDrives: "true"
  });

  const response = await fetch(
    `${DRIVE_API}/files/${encodeURIComponent(fileId)}?${params}`,
    {
      headers: {
        Authorization: `Bearer ${token}`
      }
    }
  );

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      data?.error?.message ||
      "Unable to access this Google Drive item."
    );
  }

  return data;
}

// --------------------------------------------------
// DRIVE API: LIST FOLDER CONTENTS
// --------------------------------------------------

async function driveList(token, folderId) {
  const all = [];
  let pageToken = "";

  do {
    const params = new URLSearchParams({
      q: `'${folderId.replace(/'/g, "\\'")}' in parents and trashed = false`,
      pageSize: "1000",
      orderBy: "name",
      fields:
        "nextPageToken,files(id,name,mimeType,webViewLink,thumbnailLink,size,md5Checksum,modifiedTime)",
      supportsAllDrives: "true",
      includeItemsFromAllDrives: "true"
    });

    if (pageToken) {
      params.set("pageToken", pageToken);
    }

    const response = await fetch(
      `${DRIVE_API}/files?${params}`,
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    );

    const data = await response.json();

    if (!response.ok) {
      throw new Error(
        data?.error?.message ||
        "Google Drive folder request failed."
      );
    }

    all.push(...(data.files || []));
    pageToken = data.nextPageToken || "";

  } while (pageToken);

  return all;
}

// --------------------------------------------------
// IMAGE OBJECT
// --------------------------------------------------

function createImageRecord(file, folderPath = "Individual image") {
  const links = getDriveLinks(file);

  return {
    id: file.id,
    name: file.name,
    format: extension(file.name, file.mimeType),
    mimeType: file.mimeType,

    copyLink: links.copyLink,
    directLink: links.directLink,

    folderPath,

    size: Number(file.size || 0),
    md5Checksum: file.md5Checksum || "",
    modifiedTime: file.modifiedTime || "",
    thumbnailLink: file.thumbnailLink || "",

    exactDuplicate: false,
    repeatedName: false
  };
}

// --------------------------------------------------
// SCAN FOLDER RECURSIVELY
// --------------------------------------------------

async function scanFolder(token, rootId, recursive, onProgress) {
  const queue = [
    {
      id: rootId,
      path: ""
    }
  ];

  const visited = new Set();
  const images = [];

  let foldersScanned = 0;

  while (queue.length) {
    const current = queue.shift();

    if (visited.has(current.id)) continue;

    visited.add(current.id);
    foldersScanned += 1;

    onProgress?.({
      foldersScanned,
      imageCount: images.length,
      folderName: current.path || "Root folder"
    });

    const entries = await driveList(token, current.id);

    for (const file of entries) {
      if (file.mimeType === FOLDER_MIME) {
        if (recursive && !visited.has(file.id)) {
          queue.push({
            id: file.id,
            path: current.path
              ? `${current.path}/${file.name}`
              : file.name
          });
        }
      } else if (
        (file.mimeType || "").startsWith(IMAGE_MIME_PREFIX)
      ) {
        images.push(
          createImageRecord(
            file,
            current.path || "Root folder"
          )
        );
      }
    }

    onProgress?.({
      foldersScanned,
      imageCount: images.length,
      folderName: current.path || "Root folder"
    });
  }

  return enrichDuplicates(images, foldersScanned);
}

// --------------------------------------------------
// SINGLE IMAGE PROCESSING
// --------------------------------------------------

function processSingleImage(file, onProgress) {
  if (!file.mimeType?.startsWith(IMAGE_MIME_PREFIX)) {
    throw new Error(
      "The supplied Google Drive file is not a supported image."
    );
  }

  const image = createImageRecord(file, "Individual image");

  onProgress?.({
    foldersScanned: 0,
    imageCount: 1,
    folderName: file.name
  });

  return enrichDuplicates([image], 0);
}

// --------------------------------------------------
// DUPLICATE DETECTION
// --------------------------------------------------

function enrichDuplicates(images, foldersScanned) {
  const checksumCounts = new Map();
  const nameCounts = new Map();

  for (const image of images) {
    if (image.md5Checksum) {
      checksumCounts.set(
        image.md5Checksum,
        (checksumCounts.get(image.md5Checksum) || 0) + 1
      );
    }

    const key = image.name.trim().toLowerCase();

    nameCounts.set(
      key,
      (nameCounts.get(key) || 0) + 1
    );
  }

  const enriched = images.map((image) => ({
    ...image,

    exactDuplicate:
      !!image.md5Checksum &&
      checksumCounts.get(image.md5Checksum) > 1,

    repeatedName:
      nameCounts.get(image.name.trim().toLowerCase()) > 1
  }));

  return {
    images: enriched,
    foldersScanned,

    exactDuplicateGroups:
      [...checksumCounts.values()].filter(
        (count) => count > 1
      ).length,

    exactDuplicateFiles:
      enriched.filter((image) => image.exactDuplicate).length,

    repeatedNameFiles:
      enriched.filter((image) => image.repeatedName).length
  };
}

// --------------------------------------------------
// UTILITIES
// --------------------------------------------------

function formatBytes(bytes) {
  if (!bytes) return "—";

  if (bytes < 1024) {
    return `${bytes} B`;
  }

  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)} KB`;
  }

  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

// --------------------------------------------------
// MAIN APPLICATION
// --------------------------------------------------

function App() {
  const clientId =
    import.meta.env.VITE_GOOGLE_CLIENT_ID || "";

  const [folderUrl, setFolderUrl] = useState("");
  const [recursive, setRecursive] = useState(true);

  const [token, setToken] = useState("");
  const [status, setStatus] = useState("idle");

  const [progress, setProgress] = useState({
    foldersScanned: 0,
    imageCount: 0,
    folderName: ""
  });

  const [images, setImages] = useState([]);
  const [copiedId, setCopiedId] = useState("");
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");

  const [showSettings, setShowSettings] = useState(false);

  const [query, setQuery] = useState("");
  const [formatFilter, setFormatFilter] = useState("ALL");
  const [duplicateOnly, setDuplicateOnly] = useState(false);
  const [previewMode, setPreviewMode] = useState("grid");

  const tokenClient = useRef(null);

  // --------------------------------------------------
  // GOOGLE OAUTH INITIALISATION
  // --------------------------------------------------

  useEffect(() => {
    if (
      !window.google?.accounts?.oauth2 ||
      !clientId.trim()
    ) {
      return;
    }

    tokenClient.current =
      window.google.accounts.oauth2.initTokenClient({
        client_id: clientId.trim(),

        scope:
          "https://www.googleapis.com/auth/drive.readonly",

        callback: (response) => {
          if (response.error) {
            setError(
              response.error_description ||
              response.error ||
              "Google sign-in failed."
            );

            setStatus("idle");
            return;
          }

          setToken(response.access_token);
          setError("");

          setNotice(
            "Google Drive connected successfully."
          );

          setStatus("idle");
        }
      });

  }, [clientId]);

  // --------------------------------------------------
  // FILTERS
  // --------------------------------------------------

  const formats = useMemo(
    () =>
      [...new Set(images.map((item) => item.format))].sort(),
    [images]
  );

  const filteredImages = useMemo(() => {
    const term = query.trim().toLowerCase();

    return images.filter((item) => {
      const matchesTerm =
        !term ||
        `${item.name} ${item.format} ${item.folderPath}`
          .toLowerCase()
          .includes(term);

      const matchesFormat =
        formatFilter === "ALL" ||
        item.format === formatFilter;

      const matchesDuplicate =
        !duplicateOnly ||
        item.exactDuplicate ||
        item.repeatedName;

      return (
        matchesTerm &&
        matchesFormat &&
        matchesDuplicate
      );
    });

  }, [images, query, formatFilter, duplicateOnly]);

  // --------------------------------------------------
  // GOOGLE SIGN-IN
  // --------------------------------------------------

  function signIn() {
    setError("");
    setNotice("");

    if (!clientId.trim()) {
      setShowSettings(true);

      setError(
        "Google OAuth Client ID is not configured."
      );

      return;
    }

    if (!window.google?.accounts?.oauth2) {
      setError(
        "Google sign-in is loading. Please try again."
      );

      return;
    }

    tokenClient.current?.requestAccessToken({
      prompt: token ? "" : "consent"
    });
  }

  // --------------------------------------------------
  // MAIN SCAN FUNCTION
  // SUPPORTS FOLDER AND SINGLE IMAGE LINKS
  // --------------------------------------------------

  async function handleExport() {
    setError("");
    setNotice("");
    setImages([]);

    let driveId;

    try {
      driveId = extractDriveId(folderUrl);
    } catch (err) {
      setError(err.message);
      return;
    }

    if (!token) {
      setError(
        "Connect your Google account before scanning."
      );

      return;
    }

    setStatus("scanning");

    setProgress({
      foldersScanned: 0,
      imageCount: 0,
      folderName: "Identifying Drive item..."
    });

    try {
      // First identify whether the supplied link is a folder or file.
      const metadata = await getDriveMetadata(
        token,
        driveId
      );

      let result;

      if (metadata.mimeType === FOLDER_MIME) {
        // FOLDER LINK
        result = await scanFolder(
          token,
          driveId,
          recursive,
          setProgress
        );

      } else if (
        metadata.mimeType?.startsWith(IMAGE_MIME_PREFIX)
      ) {
        // SINGLE IMAGE LINK
        result = processSingleImage(
          metadata,
          setProgress
        );

      } else {
        throw new Error(
          "Unsupported Drive item. Please provide a folder or image file URL."
        );
      }

      setImages(result.images);
      setStatus("done");

      if (metadata.mimeType === FOLDER_MIME) {
        setNotice(
          `Folder scan complete: ${result.images.length} images found in ${result.foldersScanned} folder(s).`
        );
      } else {
        setNotice(
          `Individual image processed successfully: ${metadata.name}`
        );
      }

      if (!result.images.length) {
        setNotice(
          "Scan complete. No image files were found in this folder."
        );
      }

    } catch (err) {
      setStatus("idle");

      setError(
        err.message ||
        "Unable to process this Google Drive link."
      );
    }
  }

  // --------------------------------------------------
  // EXCEL EXPORT
  // --------------------------------------------------

  async function downloadExcel() {
    if (!images.length) return;

    const workbook = new ExcelJS.Workbook();

    workbook.creator = "Asha Sweet Center";
    workbook.subject = "Product image inventory";
    workbook.title =
      "Asha Sweet Center - Drive Image Index";

    workbook.created = new Date();

    const sheet = workbook.addWorksheet("Images", {
      views: [
        {
          state: "frozen",
          ySplit: 1,
          showGridLines: false
        }
      ]
    });

    sheet.columns = [
      {
        header: "Image Name",
        key: "name",
        width: 38
      },
      {
        header: "Format",
        key: "format",
        width: 12
      },
      {
        header: "Copy Link URL",
        key: "copyLink",
        width: 48
      },
      {
        header: "Direct Image URL",
        key: "directLink",
        width: 58
      },
      {
        header: "Folder Path",
        key: "folderPath",
        width: 32
      },
      {
        header: "File Size",
        key: "size",
        width: 14
      },
      {
        header: "Exact Duplicate",
        key: "exactDuplicate",
        width: 17
      },
      {
        header: "Repeated Filename",
        key: "repeatedName",
        width: 19
      },
      {
        header: "Last Modified",
        key: "modifiedTime",
        width: 23
      }
    ];

    sheet.addRows(
      images.map((item) => ({
        name: item.name,
        format: item.format,

        copyLink: {
          text: "Open in Google Drive",
          hyperlink: item.copyLink
        },

        directLink: {
          text: item.directLink,
          hyperlink: item.directLink
        },

        folderPath: item.folderPath,
        size: formatBytes(item.size),

        exactDuplicate:
          item.exactDuplicate ? "Yes" : "No",

        repeatedName:
          item.repeatedName ? "Yes" : "No",

        modifiedTime:
          item.modifiedTime
            ? new Date(item.modifiedTime).toLocaleString()
            : ""
      }))
    );

    const header = sheet.getRow(1);
    header.height = 30;

    header.eachCell((cell) => {
      cell.font = {
        bold: true,
        color: { argb: "FFFFFFFF" },
        name: "Arial",
        size: 10
      };

      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF9D1D32" }
      };

      cell.alignment = {
        vertical: "middle",
        horizontal: "left"
      };

      cell.border = {
        bottom: {
          style: "medium",
          color: { argb: "FFD7B36A" }
        }
      };
    });

    sheet.autoFilter = {
      from: "A1",
      to: "I1"
    };

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;

      row.height = 23;

      row.eachCell((cell) => {
        cell.font = {
          name: "Arial",
          size: 10,
          color: { argb: "FF333333" }
        };

        cell.alignment = {
          vertical: "middle"
        };

        if (rowNumber % 2 === 0) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFF8F1" }
          };
        }
      });

      [3, 4].forEach((column) => {
        row.getCell(column).font = {
          color: { argb: "FF0563C1" },
          underline: true,
          size: 10
        };
      });

      if (row.getCell(7).value === "Yes") {
        row.getCell(7).font = {
          bold: true,
          color: { argb: "FFC0392B" }
        };
      }
    });

    const summary = workbook.addWorksheet("Summary");

    summary.columns = [
      { width: 28 },
      { width: 24 }
    ];

    summary.addRows([
      ["ASHA SWEET CENTER", "PRODUCT IMAGE EXPORT"],
      ["Total images", images.length],
      ["Folders scanned", progress.foldersScanned],
      [
        "Exact duplicate files",
        images.filter((item) => item.exactDuplicate).length
      ],
      [
        "Repeated filenames",
        images.filter((item) => item.repeatedName).length
      ],
      ["Export date", new Date().toLocaleString()]
    ]);

    summary.getRow(1).font = {
      bold: true,
      color: { argb: "FFFFFFFF" },
      size: 13
    };

    summary.getRow(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF9D1D32" }
    };

    const buffer = await workbook.xlsx.writeBuffer();

    const blob = new Blob(
      [buffer],
      {
        type:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      }
    );

    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");

    anchor.href = url;

    anchor.download =
      `ASC-Product-Image-Index-${new Date()
        .toISOString()
        .slice(0, 10)}.xlsx`;

    anchor.click();

    URL.revokeObjectURL(url);
  }

  // --------------------------------------------------
  // CSV EXPORT
  // --------------------------------------------------

  function downloadCsv() {
    if (!images.length) return;

    const headers = [
      "Image Name",
      "Format",
      "Copy Link URL",
      "Direct Image URL",
      "Folder Path",
      "File Size",
      "Exact Duplicate",
      "Repeated Filename",
      "Last Modified"
    ];

    const quote = (value) =>
      `"${String(value ?? "").replace(/"/g, '""')}"`;

    const rows = images.map((item) => [
      item.name,
      item.format,
      item.copyLink,
      item.directLink,
      item.folderPath,
      formatBytes(item.size),
      item.exactDuplicate ? "Yes" : "No",
      item.repeatedName ? "Yes" : "No",
      item.modifiedTime
    ]);

    const csv =
      "\uFEFF" +
      [headers, ...rows]
        .map((row) => row.map(quote).join(","))
        .join("\r\n");

    const blob = new Blob(
      [csv],
      {
        type: "text/csv;charset=utf-8;"
      }
    );

    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");

    anchor.href = url;

    anchor.download =
      `ASC-Product-Image-Index-${new Date()
        .toISOString()
        .slice(0, 10)}.csv`;

    anchor.click();

    URL.revokeObjectURL(url);
  }

  // --------------------------------------------------
  // DISCONNECT
  // --------------------------------------------------
async function copyDirectLink(image) {
  try {
    await navigator.clipboard.writeText(image.directLink);
    setCopiedId(image.id);

    setTimeout(() => {
      setCopiedId("");
    }, 2000);
  } catch (error) {
    setError("Unable to copy link. Please copy it manually.");
  }
}
  function disconnect() {
    setToken("");
    setImages([]);
    setStatus("idle");

    setNotice("Google Drive disconnected.");
    setError("");
  }

  // --------------------------------------------------
  // SUMMARY
  // --------------------------------------------------

  const exactDuplicateCount =
    images.filter((item) => item.exactDuplicate).length;

  const repeatedNameCount =
    images.filter((item) => item.repeatedName).length;

  // --------------------------------------------------
  // USER INTERFACE
  // --------------------------------------------------

  return (
    <main className="app-shell">

      {/* HEADER */}

      <header className="topbar">

        <a
          className="brand"
          href="/"
          aria-label="Asha Sweet Center Image Exporter"
        >
          <img
            className="brand-logo"
            src="/asha-logo.png"
            alt="Asha Sweet Center"
          />

          <span className="brand-divider" />

          <span className="brand-product">
            PRODUCT IMAGE
            <br />
            <b>EXPORTER</b>
          </span>
        </a>

        <div className="top-actions">

          <span className="internal-pill">
            INTERNAL TOOL
          </span>

          {token ? (
            <button
              className="account-pill"
              onClick={disconnect}
              title="Disconnect Google account"
            >
              <span className="online-dot" />
              Drive connected
              <LogOut size={15} />
            </button>
          ) : (
            <span className="secure-label">
              <ShieldCheck size={15} />
              Read-only access
            </span>
          )}

          <button
            className="settings-button"
            onClick={() => setShowSettings(!showSettings)}
          >
            Settings
            <ChevronDown size={15} />
          </button>

        </div>

      </header>

      {/* INTRODUCTION */}

      <section className="hero">

        <div className="eyebrow">
          ASHA SWEET CENTER / INTERNAL OPERATIONS
        </div>

        <h1>
          Product images,
          <br />
          <span>neatly organised.</span>
        </h1>

        <p className="hero-copy">
          Create a searchable image inventory from your Google
          Drive product folders or individual image links.
          Preview files, identify duplicates and export clean
          Excel or CSV reports.
        </p>

      </section>

      {/* MAIN WORKSPACE */}

      <section className="workspace">

        <div className="panel">

          {/* STEP 1 */}

          <div className="panel-heading">

            <div className="step-icon">
              <span>01</span>
            </div>

            <div>
              <h2>Connect Google Drive</h2>

              <p>
                Use your Asha Sweet Center Google account or
                an account with access to the required files.
              </p>
            </div>

            <div
              className={`connection-state ${
                token ? "connected" : ""
              }`}
            >
              <span />
              {token ? "Connected" : "Not connected"}
            </div>

          </div>

          <button
            className={`google-button ${
              token ? "google-connected" : ""
            }`}
            onClick={token ? disconnect : signIn}
          >
            {token ? (
              <>
                <Check size={18} />
                Google Drive connected
              </>
            ) : (
              <>
                <LogIn size={18} />
                Continue with Google
              </>
            )}
          </button>

          <p className="microcopy">
            <ShieldCheck size={13} />
            Read-only permission. Files are not modified.
          </p>

          <div className="section-divider" />

          {/* STEP 2 */}

          <div className="panel-heading folder-heading">

            <div className="step-icon">
              <span>02</span>
            </div>

            <div>
              <h2>Enter Google Drive link</h2>

              <p>
                Paste either a folder URL or an individual image URL.
              </p>
            </div>

          </div>

          <label
            className="field-label"
            htmlFor="folder-url"
          >
            GOOGLE DRIVE FOLDER OR IMAGE LINK
          </label>

          <div className="url-input-wrap">

            <Link2 size={18} />

            <input
              id="folder-url"
              value={folderUrl}
              onChange={(e) => setFolderUrl(e.target.value)}
              placeholder="Paste Google Drive folder or image URL..."
            />

            {folderUrl && (
              <button
                className="clear-input"
                onClick={() => setFolderUrl("")}
                aria-label="Clear folder link"
              >
                <X size={16} />
              </button>
            )}

          </div>

          <label className="check-row">

            <input
              type="checkbox"
              checked={recursive}
              onChange={(e) => setRecursive(e.target.checked)}
            />

            <span className="custom-check">
              <Check size={12} />
            </span>

            <span>
              <strong>Include subfolders</strong>

              <small>
                Applies when the supplied link is a folder.
              </small>
            </span>

          </label>

          <button
            className="primary-button"
            disabled={status === "scanning"}
            onClick={handleExport}
          >
            {status === "scanning" ? (
              <>
                <LoaderCircle className="spin" size={18} />
                Processing Drive link…
              </>
            ) : (
              <>
                <FolderOpen size={18} />
                Process Drive link
              </>
            )}
          </button>

          {status === "scanning" && (
            <div className="progress-box">

              <div className="progress-top">
                <span>
                  Processing: {progress.folderName}
                </span>

                <span>
                  {progress.imageCount} images
                </span>
              </div>

              <div className="progress-track">
                <span />
              </div>

              <small>
                {progress.foldersScanned} folder(s) checked ·
                {" "}
                {progress.imageCount} image(s) discovered
              </small>

            </div>
          )}

          {error && (
            <div className="message error-message">
              <CircleAlert size={17} />
              <span>{error}</span>
            </div>
          )}

          {notice && !error && (
            <div className="message success-message">
              <Check size={17} />
              <span>{notice}</span>
            </div>
          )}

        </div>

        {/* SIDE INFORMATION */}

        <aside className="side-panel">

          <div className="side-card">

            <div className="side-icon">
              <FileSpreadsheet size={21} />
            </div>

            <h3>Made for product-image operations</h3>

            <p>
              Keep product photography organised for catalogues,
              listings and internal teams.
            </p>

            <div className="included-list">

              {[
                "Folder and individual image links",
                "Image name and format",
                "Drive and direct image links",
                "Thumbnail previews",
                "Duplicate file detection",
                "Formatted Excel and CSV"
              ].map((item) => (
                <div key={item}>
                  <Check size={14} />
                  {item}
                </div>
              ))}

            </div>

          </div>

          <div className="privacy-note">

            <ShieldCheck size={17} />

            <div>
              <strong>Private to your workflow</strong>

              <p>
                Folder scanning and report generation run in your
                browser. No image files are uploaded to this app.
              </p>
            </div>

          </div>

        </aside>

      </section>

      {/* RESULTS */}

      {images.length > 0 && (
        <section className="results-section">

          <div className="results-header">

            <div>
              <div className="eyebrow">
                PRODUCT IMAGE INVENTORY
              </div>

              <h2>
                {images.length.toLocaleString()} images found
              </h2>

              <p>
                {progress.foldersScanned} folders scanned ·
                {" "}
                {formats.length} image formats detected
              </p>
            </div>

            <div className="export-actions">

              <button
                className="csv-button"
                onClick={downloadCsv}
              >
                Export CSV
              </button>

              <button
                className="download-button"
                onClick={downloadExcel}
              >
                <ArrowDownToLine size={18} />
                Download Excel
              </button>

            </div>

          </div>
          {images.length === 1 &&
  images[0].folderPath === "Individual image" && (
    <div className="single-link-result">
      <div className="single-link-heading">
        <div>
          <div className="eyebrow">DIRECT IMAGE LINK</div>
          <h3>{images[0].name}</h3>
          <p>Your Google-hosted image URL is ready to copy.</p>
        </div>

        <button
          className="copy-link-button"
          onClick={() => copyDirectLink(images[0])}
        >
          {copiedId === images[0].id ? (
            <Check size={16} />
          ) : (
            <Copy size={16} />
          )}

          {copiedId === images[0].id ? "Copied!" : "Copy link"}
        </button>
      </div>

      <div className="direct-link-field">
        <input
          type="text"
          readOnly
          value={images[0].directLink}
          onFocus={(e) => e.target.select()}
        />
      </div>
    </div>
  )}

          {/* STATISTICS */}

          <div className="stat-grid">

            <div className="stat-card">
              <span>Total images</span>

              <strong>
                {images.length.toLocaleString()}
              </strong>
            </div>

            <div className="stat-card">
              <span>Exact duplicate files</span>

              <strong
                className={
                  exactDuplicateCount ? "stat-warning" : ""
                }
              >
                {exactDuplicateCount}
              </strong>
            </div>

            <div className="stat-card">
              <span>Repeated filenames</span>

              <strong>{repeatedNameCount}</strong>
            </div>

            <div className="stat-card">
              <span>Folders scanned</span>

              <strong>{progress.foldersScanned}</strong>
            </div>

          </div>

          {/* SEARCH AND FILTERS */}

          <div className="results-toolbar">

            <div className="toolbar-title">
              <ImageIcon size={15} />

              {filteredImages.length} of {images.length} shown
            </div>

            <div className="filter-controls">

              <div className="search-wrap">
                <Search size={15} />

                <input
                  value={query}
                  onChange={(e) => setQuery(e.target.value)}
                  placeholder="Search name or folder..."
                />
              </div>

              <select
                value={formatFilter}
                onChange={(e) => setFormatFilter(e.target.value)}
                aria-label="Filter by image format"
              >
                <option value="ALL">All formats</option>

                {formats.map((format) => (
                  <option key={format} value={format}>
                    {format}
                  </option>
                ))}
              </select>

              <button
                className={`filter-toggle ${
                  duplicateOnly ? "active" : ""
                }`}
                onClick={() => setDuplicateOnly(!duplicateOnly)}
              >
                <Copy size={14} />
                Duplicates
              </button>

              <div className="view-toggle">

                <button
                  className={
                    previewMode === "grid" ? "selected" : ""
                  }
                  onClick={() => setPreviewMode("grid")}
                  title="Grid view"
                >
                  <ImageIcon size={15} />
                </button>

                <button
                  className={
                    previewMode === "list" ? "selected" : ""
                  }
                  onClick={() => setPreviewMode("list")}
                  title="List view"
                >
                  <Files size={15} />
                </button>

              </div>

            </div>

          </div>

          {/* GRID VIEW */}

          {previewMode === "grid" ? (

            <div className="image-grid">

              {filteredImages.slice(0, 240).map((item) => (

                <article
                  className="image-card"
                  key={item.id}
                >

                  <a
                    className="thumb-wrap"
                    href={item.copyLink}
                    target="_blank"
                    rel="noreferrer"
                    title="Open image in Google Drive"
                  >
                    <DriveImage
                      file={item}
                      token={token}
                      alt={item.name}
                    />
                  </a>

                  <div className="image-card-body">

                    <div
                      className="image-name"
                      title={item.name}
                    >
                      {item.name}
                    </div>

                    <div className="image-meta">
                      <span>{item.format}</span>
                      <span>{formatBytes(item.size)}</span>
                    </div>

                    <div className="image-flags">

                      {item.exactDuplicate && (
                        <span className="duplicate-tag">
                          Exact duplicate
                        </span>
                      )}

                      {!item.exactDuplicate &&
                        item.repeatedName && (
                          <span className="name-tag">
                            Repeated name
                          </span>
                        )}

                    </div>

                    <a
                      className="drive-link"
                      href={item.copyLink}
                      target="_blank"
                      rel="noreferrer"
                    >
                      Open in Drive ↗
                    </a>

                  </div>

                </article>

              ))}

            </div>

          ) : (

            /* TABLE VIEW */

            <div className="table-wrap">

              <table>

                <thead>
                  <tr>
                    <th>PREVIEW</th>
                    <th>IMAGE NAME</th>
                    <th>FORMAT</th>
                    <th>SIZE</th>
                    <th>FOLDER</th>
                    <th>CHECK</th>
                    <th>DRIVE LINK</th>
                  </tr>
                </thead>

                <tbody>

                  {filteredImages.slice(0, 500).map((item) => (

                    <tr key={item.id}>

                      <td>
                        <DriveImage
                          file={item}
                          token={token}
                          className="table-thumb"
                          alt=""
                        />
                      </td>

                      <td className="file-cell">
                        {item.name}
                      </td>

                      <td>
                        <span className="format-tag">
                          {item.format}
                        </span>
                      </td>

                      <td>{formatBytes(item.size)}</td>

                      <td className="folder-cell">
                        {item.folderPath}
                      </td>

                      <td>
                        {item.exactDuplicate ? (
                          <span className="duplicate-tag">
                            Duplicate
                          </span>
                        ) : item.repeatedName ? (
                          <span className="name-tag">
                            Repeated name
                          </span>
                        ) : (
                          "—"
                        )}
                      </td>

                      <td>
                        <a
                          href={item.copyLink}
                          target="_blank"
                          rel="noreferrer"
                        >
                          Open ↗
                        </a>
                      </td>

                    </tr>

                  ))}

                </tbody>

              </table>

            </div>

          )}

          {filteredImages.length >
            (previewMode === "grid" ? 240 : 500) && (

            <p className="table-footnote">
              Preview is limited for performance. Excel and CSV
              exports include all {images.length} images.
            </p>

          )}

        </section>
      )}

      {/* FOOTER */}

      <footer>
        <span>ASHA SWEET CENTER · EST. 1951</span>

        <span>
          Internal product-image utility · Google Drive read-only
        </span>
      </footer>

      {/* SETTINGS */}

      {showSettings && (

        <div
          className="modal-backdrop"
          onMouseDown={(e) => {
            if (e.target === e.currentTarget) {
              setShowSettings(false);
            }
          }}
        >

          <section className="settings-modal">

            <button
              className="modal-close"
              onClick={() => setShowSettings(false)}
              aria-label="Close settings"
            >
              <X size={18} />
            </button>

            <img
              className="settings-logo"
              src="/asha-logo.png"
              alt="Asha Sweet Center"
            />

            <h2>Google Drive connection</h2>

            <p>
              The OAuth Client ID is loaded from your local
              environment configuration.
            </p>

            <div className="setup-hint">

              <strong>Environment configuration</strong>

              <p>
                Create a <code>.env.local</code> file in the
                project root and add:
                {" "}
                <code>
                  VITE_GOOGLE_CLIENT_ID=your-client-id
                </code>
                .
                Restart the development server after saving.
              </p>

            </div>

            <div className="setup-hint">

              <strong>Google Cloud setup</strong>

              <ol>
                <li>
                  Enable Google Drive API.
                </li>

                <li>
                  Configure Google Auth Platform.
                </li>

                <li>
                  Create an OAuth Client ID using Web application.
                </li>

                <li>
                  Add localhost and your Vercel domain under
                  Authorized JavaScript origins.
                </li>
              </ol>

              <a
                href="https://console.cloud.google.com/auth/clients"
                target="_blank"
                rel="noreferrer"
              >
                Open Google Cloud OAuth clients ↗
              </a>

            </div>

            <button
              className="primary-button"
              onClick={() => setShowSettings(false)}
            >
              Close
            </button>

          </section>

        </div>

      )}

    </main>
  );
}

export default App;
