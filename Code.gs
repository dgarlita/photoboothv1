const DEFAULT_FOLDER_ID = "1UvhMAEjGSuIfg8fuWdbyzPtUBzz01iYR";

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";
    const eventId = (e && e.parameter && e.parameter.eventId) ? e.parameter.eventId : "";
    const folderId = (e && e.parameter && e.parameter.folderId) ? e.parameter.folderId : DEFAULT_FOLDER_ID;

    if (action === "list" && eventId) {
      const photos = listEventPhotos(folderId, eventId);
      return jsonResponse({ status: "success", photos: photos });
    }

    return jsonResponse({
      status: "ok",
      message: "Rafqa Production API Aktif",
      folderId: folderId
    });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: "error", message: "Payload POST kosong" });
    }

    const data = JSON.parse(e.postData.contents);
    const folderId = data.folderId || data.rootFolderId || DEFAULT_FOLDER_ID;
    const action = data.action || "upload";

    if (action === "upload") {
      const result = savePhotoToDrive(folderId, data);
      return jsonResponse({
        status: "success",
        fileId: result.fileId,
        fileUrl: result.fileUrl,
        fileName: result.fileName,
        folderName: result.folderName
      });
    }

    if (action === "like" && data.fileId) {
      const likes = incrementPhotoLike(data.fileId, data.delta || 1);
      return jsonResponse({ status: "success", fileId: data.fileId, likes: likes });
    }

    return jsonResponse({ status: "error", message: "Action tidak dikenali: " + action });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function savePhotoToDrive(rootFolderId, data) {
  const rawBase64 = data.imageBase64 || data.image || data.photo || data.dataUrl || "";
  if (!rawBase64) {
    throw new Error("Payload imageBase64 kosong.");
  }

  const rootFolder = DriveApp.getFolderById(rootFolderId);
  const eventFolderName = sanitizeFolderName(data.eventName || data.eventId || "Rafqa-Event");
  const eventFolder = getOrCreateSubfolder(rootFolder, eventFolderName);

  const base64Clean = String(rawBase64).replace(/^data:image\/[a-zA-Z0-9+.-]+;base64,/, "");
  const bytes = Utilities.base64Decode(base64Clean);
  const fileName = data.filename || data.fileName || ("RP_" + Date.now() + ".jpg");
  const mimeType = data.mimeType || "image/jpeg";

  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = eventFolder.createFile(blob);

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  file.setDescription(JSON.stringify({
    eventId: data.eventId || "",
    guestName: data.guestName || "Tamu Undangan",
    guestWish: data.guestWish || "",
    likes: 0,
    timestamp: data.timestamp || new Date().toISOString()
  }));

  const fileId = file.getId();
  return {
    fileId: fileId,
    fileName: file.getName(),
    folderName: eventFolder.getName(),
    fileUrl: "https://lh3.googleusercontent.com/d/" + fileId
  };
}

function incrementPhotoLike(fileId, delta) {
  const file = DriveApp.getFileById(fileId);
  let meta = {};
  try {
    meta = JSON.parse(file.getDescription() || "{}");
  } catch (err) {}
  const nextLikes = Math.max(0, (Number(meta.likes) || 0) + (Number(delta) || 1));
  meta.likes = nextLikes;
  file.setDescription(JSON.stringify(meta));
  return nextLikes;
}

function listEventPhotos(rootFolderId, eventId) {
  const rootFolder = DriveApp.getFolderById(rootFolderId);
  const folders = rootFolder.getFolders();
  const results = [];

  while (folders.hasNext()) {
    const folder = folders.next();
    const files = folder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType().indexOf("image/") !== 0) continue;

      let meta = {};
      try {
        meta = JSON.parse(file.getDescription() || "{}");
      } catch (err) {}

      if (meta.eventId && meta.eventId !== eventId) continue;

      const fileId = file.getId();
      results.push({
        id: fileId,
        name: file.getName(),
        guestName: meta.guestName || "Tamu Undangan",
        guestWish: meta.guestWish || "",
        likes: Number(meta.likes) || 0,
        url: "https://lh3.googleusercontent.com/d/" + fileId,
        timestamp: Utilities.formatDate(file.getDateCreated(), "Asia/Jakarta", "HH:mm")
      });
    }
  }

  return results.reverse();
}

function getOrCreateSubfolder(parentFolder, folderName) {
  const iter = parentFolder.getFoldersByName(folderName);
  if (iter.hasNext()) return iter.next();
  return parentFolder.createFolder(folderName);
}

function sanitizeFolderName(name) {
  return String(name).replace(/[\\/:*?"<>|]/g, "-").trim().slice(0, 80);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Jalankan fungsi ini sekali di editor Apps Script untuk memberikan izin akses Google Drive
function authorizeDriveAccess() {
  const folder = DriveApp.getFolderById(DEFAULT_FOLDER_ID);
  Logger.log("Izin Google Drive aktif untuk folder: " + folder.getName());
}
