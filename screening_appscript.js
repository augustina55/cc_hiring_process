const SHEET_ID     = '1cr_poEvii21bVATswTJizDBa2bHT4AS6vQugANCy7Po';
const VIDEO_FOLDER = 'cc_hiring_videos';
const TEMP_FOLDER  = 'cc_hiring_temp';
// Set via: Project Settings > Script Properties > OPENAI_API_KEY
const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
const SHEET_NAME = "Sheet1";

function testOpenAI() {

  const response = UrlFetchApp.fetch(
    "https://api.openai.com/v1/models",
    {
      method: "get",
      headers: {
        Authorization: "Bearer " + OPENAI_API_KEY
      },
      muteHttpExceptions: true
    }
  );

  Logger.log(response.getResponseCode());
  Logger.log(response.getContentText());
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'chunk')    return uploadChunk(data);
    if (data.action === 'finalize') return finalizeUpload(data);
    if (data.action === 'submit')   return submitForm(data);
    return jsonResp({ ok: false, error: 'Unknown action: ' + data.action });
  } catch(err) {
    return jsonResp({ ok: false, error: err.toString() });
  }
}

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getFolder(name) {
  const props = PropertiesService.getScriptProperties();
  const key   = 'folderId_' + name;

  const cachedId = props.getProperty(key);
  if (cachedId) {
    try { return DriveApp.getFolderById(cachedId); } catch (e) { /* deleted — fall through */ }
  }

  // Concurrent chunk uploads can race here: DriveApp.getFoldersByName()
  // doesn't enforce uniqueness, so two simultaneous calls can each fail to
  // find the folder and both create their own "cc_hiring_temp", scattering
  // chunks across duplicate folders (finalize then reports "No chunks
  // found" if it happens to look in the wrong one). A lock plus caching the
  // resolved folder ID makes every call converge on the same folder.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const cachedId2 = props.getProperty(key);
    if (cachedId2) {
      try { return DriveApp.getFolderById(cachedId2); } catch (e) { /* deleted — fall through */ }
    }
    const it = DriveApp.getFoldersByName(name);
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder(name);
    props.setProperty(key, folder.getId());
    return folder;
  } finally {
    lock.releaseLock();
  }
}

function uploadChunk(data) {
  try {
    const name  = data.uploadId + '__' + String(data.index).padStart(6, '0');
    const bytes = Utilities.base64Decode(data.chunk);
    const blob  = Utilities.newBlob(bytes, 'application/octet-stream', name);
    getFolder(TEMP_FOLDER).createFile(blob);
    return jsonResp({ ok: true });
  } catch(err) {
    return jsonResp({ ok: false, error: 'chunk failed: ' + err.toString() });
  }
}

function finalizeUpload(data) {
  try {
    const tempFolder  = getFolder(TEMP_FOLDER);
    const videoFolder = getFolder(VIDEO_FOLDER);

    const chunkMap = {};
    const iter = tempFolder.getFiles();
    while (iter.hasNext()) {
      const f = iter.next();
      const n = f.getName();
      if (n.startsWith(data.uploadId + '__')) {
        chunkMap[parseInt(n.split('__')[1], 10)] = f;
      }
    }

    const indices = Object.keys(chunkMap).map(Number).sort((a, b) => a - b);
    if (!indices.length) return jsonResp({ ok: false, error: 'No chunks found' });

    const parts = [];
    let total = 0;
    for (const i of indices) {
      const b = chunkMap[i].getBlob().getBytes();
      parts.push(b); total += b.length;
    }
    const all = new Array(total);
    let off = 0;
    for (const b of parts) { for (let i = 0; i < b.length; i++) all[off++] = b[i]; }
    for (const i of indices) chunkMap[i].setTrashed(true);

    const mime  = data.mimeType || 'video/mp4';
    const ext   = mime.includes('webm') ? 'webm' : mime.includes('quicktime') ? 'mov' : 'mp4';
    const safe  = (data.applicantName || 'applicant').replace(/[^a-zA-Z0-9 _-]/g, '_');
    const fname = safe + '_' + Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd') + '.' + ext;

    const file = videoFolder.createFile(Utilities.newBlob(all, mime, fname));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return jsonResp({ ok: true, url: 'https://drive.google.com/file/d/' + file.getId() + '/view' });
  } catch(err) {
    return jsonResp({ ok: false, error: 'finalize failed: ' + err.toString() });
  }
}

function submitForm(data) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp', 'Name', 'Phone',
        'Video URL', 'Submission Type', 'Transcript_text'
      ]);
    }
    sheet.appendRow([
      new Date(),
      data.name,
      data.phone,
      data.videoUrl,
      data.submissionType,
      data.transcript || ''          // speech-to-text from recording
    ]);
  } catch(sheetErr) {
    Logger.log('Sheet write failed: ' + sheetErr);
  }

  return jsonResp({ ok: true });
}


//video_trancription 

function processCoachApplications() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);

  const data = sheet.getDataRange().getValues();

  for (let row = 1; row < data.length; row++) {

    const processed = data[row][15]; // P — Processed marker

    if (processed === "DONE") continue;

    const videoUrl = data[row][3]; // D — Video URL

    if (!videoUrl) continue;

    try {

      Logger.log("Processing Row " + (row + 1));

      const fileId = extractDriveFileId(videoUrl);

      const transcript = transcribeVideo(fileId);

      const analysis = evaluateCoach(transcript);

      sheet.getRange(row + 1, 6).setValue(transcript);           // F — Transcript_text
      sheet.getRange(row + 1, 7).setValue(analysis.communication);
      sheet.getRange(row + 1, 8).setValue(analysis.teaching);
      sheet.getRange(row + 1, 9).setValue(analysis.kid_friendly);
      sheet.getRange(row + 1, 10).setValue(analysis.engagement);
      sheet.getRange(row + 1, 11).setValue(analysis.chess_accuracy);
      sheet.getRange(row + 1, 12).setValue(analysis.overall);
      sheet.getRange(row + 1, 13).setValue(analysis.status);
      sheet.getRange(row + 1, 14).setValue(
        analysis.strengths.join("\n")
      );
      sheet.getRange(row + 1, 15).setValue(
        analysis.concerns.join("\n")
      );
      sheet.getRange(row + 1, 16).setValue("DONE");             // P — Processed marker

    } catch (e) {

      Logger.log(e);

      sheet.getRange(row + 1, 16)
        .setValue("ERROR: " + e.message);
    }
  }
}


//drive file id
function extractDriveFileId(url) {

  const match = url.match(/[-\w]{25,}/);

  if (!match) {
    throw new Error("Invalid Drive URL");
  }

  return match[0];
}

//trancription 
function transcribeVideo(fileId) {

  const file = DriveApp.getFileById(fileId);

  const blob = file.getBlob();

  const response = UrlFetchApp.fetch(
    "https://api.openai.com/v1/audio/transcriptions",
    {
      method: "post",
      headers: {
        Authorization: "Bearer " + OPENAI_API_KEY
      },
      payload: {
        model: "gpt-4o-mini-transcribe",
        file: blob
      },
      muteHttpExceptions: true
    }
  );

  const result = JSON.parse(
    response.getContentText()
  );

  if (!result.text) {
    throw new Error(response.getContentText());
  }

  return result.text;
}

function evaluateCoach(transcript) {

  const prompt = `
You are a senior hiring manager evaluating chess coaches who teach children aged 6-12.

Analyze this transcript.

Return ONLY valid JSON.

{
  "communication": 0,
  "teaching": 0,
  "kid_friendly": 0,
  "engagement": 0,
  "chess_accuracy": 0,
  "overall": 0,

  "status": "",

  "english_level": "",
  "teaching_style": "",
  "recommended_student_level": "",

  "strengths": [],
  "concerns": [],
  "red_flags": [],

  "summary": "",
  "hire_recommendation": ""
}
Scoring:

Communication:
Grammar, fluency, confidence.

Teaching:
Structured explanation,
examples,
clarity.

Kid Friendly:
Simple language,
encouraging tone.

Engagement:
Interactive teaching,
questions,
student involvement.

Chess Accuracy:
Correctness of chess concept.

Status:

Strongly Recommended:
overall >= 8.5

Recommended:
overall >= 7

Borderline:
overall >= 6

Reject:
overall < 6

Transcript:

${transcript}
`;

  const payload = {
    model: "gpt-4.1-mini",
    response_format: {
      type: "json_object"
    },
    messages: [
      {
        role: "user",
        content: prompt
      }
    ],
    temperature: 0.2
  };

  const response = UrlFetchApp.fetch(
    "https://api.openai.com/v1/chat/completions",
    {
      method: "post",
      headers: {
        Authorization: "Bearer " + OPENAI_API_KEY,
        "Content-Type": "application/json"
      },
      payload: JSON.stringify(payload)
    }
  );

  const result = JSON.parse(
    response.getContentText()
  );

  return JSON.parse(
    result.choices[0].message.content
  );
}

