const SHEET_ID     = '1cr_poEvii21bVATswTJizDBa2bHT4AS6vQugANCy7Po';
const VIDEO_FOLDER = 'cc_hiring_videos';
const TEMP_FOLDER  = 'cc_hiring_temp';
// Set via: Project Settings > Script Properties > GROQ_API_KEY
const GROQ_API_KEY = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
const SHEET_NAME = "Sheet2";
const TRANSCRIBE_URL = 'https://screening-phi.vercel.app/api/transcribe';

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp', 'Name', 'Phone',                          // A-C
      'Video URL', 'Submission Type',                        // D-E
      'Transcript_text',                                     // F
      'Communication', 'Communication Reason',               // G-H
      'Teaching', 'Teaching Reason',                          // I-J
      'Kid Friendly', 'Kid Friendly Reason',                  // K-L
      'Engagement', 'Engagement Reason',                      // M-N
      'Chess Accuracy', 'Chess Accuracy Reason',              // O-P
      'Overall', 'Status', 'Decision Reason',                 // Q-S
      'Transcript Confidence', 'Transcript Confidence Reason',// T-U
      'Strengths', 'Concerns',                                // V-W
      'Processed'                                             // X
    ]);
  }
  return sheet;
}

function testGroq() {

  const response = UrlFetchApp.fetch(
    "https://api.groq.com/openai/v1/models",
    {
      method: "get",
      headers: {
        Authorization: "Bearer " + GROQ_API_KEY
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

    let totalSize = 0;
    for (const i of indices) totalSize += chunkMap[i].getSize();

    const mime  = data.mimeType || 'video/mp4';
    const ext   = mime.includes('webm') ? 'webm' : mime.includes('quicktime') ? 'mov' : 'mp4';
    const safe  = (data.applicantName || 'applicant').replace(/[^a-zA-Z0-9 _-]/g, '_');
    const fname = safe + '_' + Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd') + '.' + ext;

    // Utilities.newBlob() caps out around 50MB, so a video of any real
    // length can never be assembled into a single in-memory blob. Instead,
    // stream the (already ≤3MB) chunk files straight into Drive via a
    // resumable upload session — each PUT stays small regardless of the
    // total file size.
    const token = ScriptApp.getOAuthToken();
    const initResp = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ name: fname, parents: [videoFolder.getId()], mimeType: mime }),
      muteHttpExceptions: true
    });
    if (initResp.getResponseCode() !== 200) {
      throw new Error('Could not start resumable upload: ' + initResp.getContentText());
    }
    const uploadUrl = initResp.getHeaders()['Location'] || initResp.getHeaders()['location'];
    if (!uploadUrl) throw new Error('Resumable upload session had no Location header');

    let offset = 0, fileId = null;
    for (const i of indices) {
      const chunkBlob = chunkMap[i].getBlob();
      const size  = chunkBlob.getBytes().length;
      const start = offset, end = offset + size - 1;
      const isLast = (i === indices[indices.length - 1]);

      const putResp = UrlFetchApp.fetch(uploadUrl, {
        method: 'put',
        headers: { 'Content-Range': 'bytes ' + start + '-' + end + '/' + totalSize },
        payload: chunkBlob,
        muteHttpExceptions: true
      });
      const code = putResp.getResponseCode();
      offset += size;

      if (isLast) {
        if (code !== 200 && code !== 201) {
          throw new Error('Resumable upload did not complete: ' + putResp.getContentText());
        }
        fileId = JSON.parse(putResp.getContentText()).id;
      } else if (code !== 308) {
        throw new Error('Resumable upload chunk failed at index ' + i + ': ' + putResp.getContentText());
      }
    }
    if (!fileId) throw new Error('Resumable upload finished without a file id');

    for (const i of indices) chunkMap[i].setTrashed(true);

    const file = DriveApp.getFileById(fileId);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return jsonResp({ ok: true, url: 'https://drive.google.com/file/d/' + file.getId() + '/view' });
  } catch(err) {
    return jsonResp({ ok: false, error: 'finalize failed: ' + err.toString() });
  }
}

function submitForm(data) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAME);
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

    const processed = data[row][23]; // X — Processed marker

    if (processed === "DONE") continue;

    const videoUrl = data[row][3]; // D — Video URL

    if (!videoUrl) continue;

    try {

      Logger.log("Processing Row " + (row + 1));

      const fileId = extractDriveFileId(videoUrl);

      const transcript = transcribeViaVercel(fileId);

      const analysis = evaluateCoach(transcript);

      sheet.getRange(row + 1, 6).setValue(transcript);                    // F  — Transcript_text
      sheet.getRange(row + 1, 7).setValue(analysis.communication);        // G
      sheet.getRange(row + 1, 8).setValue(analysis.communication_reason); // H
      sheet.getRange(row + 1, 9).setValue(analysis.teaching);             // I
      sheet.getRange(row + 1, 10).setValue(analysis.teaching_reason);     // J
      sheet.getRange(row + 1, 11).setValue(analysis.kid_friendly);        // K
      sheet.getRange(row + 1, 12).setValue(analysis.kid_friendly_reason); // L
      sheet.getRange(row + 1, 13).setValue(analysis.engagement);          // M
      sheet.getRange(row + 1, 14).setValue(analysis.engagement_reason);   // N
      sheet.getRange(row + 1, 15).setValue(analysis.chess_accuracy);      // O
      sheet.getRange(row + 1, 16).setValue(analysis.chess_accuracy_reason); // P
      sheet.getRange(row + 1, 17).setValue(analysis.overall);             // Q
      sheet.getRange(row + 1, 18).setValue(analysis.status);              // R
      sheet.getRange(row + 1, 19).setValue(analysis.decision_reason);     // S
      sheet.getRange(row + 1, 20).setValue(analysis.transcript_confidence); // T
      sheet.getRange(row + 1, 21).setValue(analysis.transcript_confidence_reason); // U
      sheet.getRange(row + 1, 22).setValue(
        (analysis.strengths || []).join("\n")
      );
      sheet.getRange(row + 1, 23).setValue(
        (analysis.concerns || []).join("\n")
      );
      sheet.getRange(row + 1, 24).setValue("DONE");                       // X — Processed marker

    } catch (e) {

      Logger.log(e);

      sheet.getRange(row + 1, 24)
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
// Apps Script can't do the audio extraction a large video needs before
// transcription (Utilities.newBlob()/getBlob() cap out around 50MB, and
// Groq's own upload limit is ~25MB — both well under a multi-minute video).
// Vercel downloads the video from Drive directly, strips it to a small
// audio-only track with ffmpeg, and transcribes that instead.
function transcribeViaVercel(fileId) {
  const response = UrlFetchApp.fetch(TRANSCRIBE_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ fileId }),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code !== 200) {
    throw new Error('Transcription service returned ' + code + ': ' + text.slice(0, 300));
  }

  const result = JSON.parse(text);
  if (!result.ok) throw new Error(result.error || 'Transcription failed');

  return result.transcript;
}

function evaluateCoach(transcript) {

  const prompt = `
You are a senior hiring manager evaluating chess coaches who teach children aged 6-12.
You are working ONLY from a speech-to-text transcript — there is no audio or video signal,
so tone of voice, facial expression, and body language cannot be assessed. Base every
judgment strictly on what the words themselves show, and say so explicitly when something
can't be determined from text alone (e.g. actual vocal warmth or energy).

Every score MUST be justified by a "_reason" field that either quotes or closely paraphrases
a specific part of the transcript. Do not give a reason that just restates the score in words
("engagement is high because coach is engaging") — point at the actual evidence. If the
transcript has nothing supporting a category, say so plainly and score it low.

For "engagement" specifically, look for concrete textual signals of interaction with the
(imagined) student — not just enthusiasm. Evidence includes:
- Direct address ("you", "can you see", "your turn")
- Checking for understanding ("does that make sense?", "do you see why?")
- Posing questions and pausing for a response, even a rhetorical one
- Encouragement or positive reinforcement ("great job", "well done", "nice try")
- Inviting the student to try something ("now you try", "let's practice this together")
A transcript that is a one-way, uninterrupted lecture with no direct address or questions
should score LOW on engagement even if the explanation itself is otherwise clear — note that
explicitly in "engagement_reason".

Also assess whether the transcript itself is usable for evaluation. Speech-to-text can produce
garbled or truncated output; if the transcript is very short, incoherent, or clearly missing
large portions of speech, set "transcript_confidence" to "low" and say why in
"transcript_confidence_reason" — a low-confidence transcript means the scores below are
unreliable and the application should be reviewed manually.

Return ONLY valid JSON in this exact shape:

{
  "communication": 0, "communication_reason": "",
  "teaching": 0, "teaching_reason": "",
  "kid_friendly": 0, "kid_friendly_reason": "",
  "engagement": 0, "engagement_reason": "",
  "chess_accuracy": 0, "chess_accuracy_reason": "",
  "overall": 0,

  "status": "",
  "decision_reason": "",

  "transcript_confidence": "high",
  "transcript_confidence_reason": "",

  "strengths": [],
  "concerns": [],
  "red_flags": []
}

Scoring guide (each 0-10):

Communication:
Grammar, fluency, confidence in how ideas are expressed.

Teaching:
Structured explanation, use of examples, clarity of the chess concept.

Kid Friendly:
Simple language a 6-12 year old could follow, encouraging tone in the wording used.

Engagement:
Concrete interaction signals as described above — NOT just how enthusiastic the topic sounds.

Chess Accuracy:
Correctness of the chess concept actually explained.

"overall" is your holistic judgment (not a strict average) of hireability for teaching kids.

Status (based on "overall"):
Strongly Recommended: overall >= 8.5
Recommended: overall >= 7
Borderline: overall >= 6
Reject: overall < 6

"decision_reason" must be 2-4 sentences explaining specifically why this status was chosen,
citing the strongest piece of supporting evidence from the transcript (quote or close
paraphrase) — a hiring manager should be able to read only this field and understand the call.

Transcript:

${transcript}
`;

  // Confirmed available on this account via testGroq() (GET /v1/models).
  // If this ever 404s again, re-run testGroq() and pick another id that
  // lists "json_mode" in supported_features.
  const payload = {
    model: "openai/gpt-oss-120b",
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
    "https://api.groq.com/openai/v1/chat/completions",
    {
      method: "post",
      headers: {
        Authorization: "Bearer " + GROQ_API_KEY,
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

