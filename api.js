/* ====================================================================
 * APPLICANT REGISTRATION — API ABSTRACTION
 * --------------------------------------------------------------------
 * SINGLE INTEGRATION POINT between the registration form and the
 * Google Apps Script backend. No other file contains the endpoint.
 *
 * WHERE TO CONFIGURE THE APPS SCRIPT URL:
 *   Set APPS_SCRIPT_URL below to your Apps Script Web App "/exec" URL
 *   (Deploy > New deployment > Web app, Execute as: Me,
 *   Who has access: Anyone).
 *
 *   GitHub Pages has no environment variables, so this constant fills
 *   the role of VITE_APPS_SCRIPT_URL. You may also inject it at runtime
 *   BEFORE this script loads:
 *       window.VITE_APPS_SCRIPT_URL = 'https://script.google.com/...';
 *
 * No credentials or secrets belong in this file — the Web App URL is
 * public by design; all Sheet/Drive/Gmail access stays server-side.
 * ==================================================================== */

window.REGISTRATION_API = (function () {

  var APPS_SCRIPT_URL =
    window.VITE_APPS_SCRIPT_URL ||
    'https://script.google.com/macros/s/AKfycbwgB4E6gqU5aVlaa-dNbsC6UotAIRXnwiMeWGDnaGH8Et4seXELY7lQdx0gb40JJwJMOA/exec';

  var CONFIG = {
    APPS_SCRIPT_URL: APPS_SCRIPT_URL,
    REQUEST_TIMEOUT_MS: 60000,

    // EDIT ME: venues offered for "Location for Interview:"
    INTERVIEW_LOCATIONS: [
      'PESO Cagayan - Provincial Capitol, Tuguegarao City',
      'Provincial Capitol Grounds, Tuguegarao City'
    ],

    MAX_FILE_MB: 5
  };

  function ApiError(userMessage) {
    var e = new Error(userMessage);
    e.name = 'ApiError';
    e.userMessage = userMessage;
    return e;
  }

  var MSG_NETWORK = 'We could not submit your registration. Please check your internet connection and try again.';
  var MSG_BACKEND = 'We could not submit your registration. Please try again in a few minutes.';
  var MSG_TIMEOUT = 'The submission took too long and may not have gone through. Please check your connection and try again.';

  var CODE_MESSAGES = {
    busy: 'The system is busy right now. Please try again in a minute.',
    rate_limited: 'Too many registration attempts for this email right now. Please try again later.',
    rejected: 'This submission could not be accepted. Please contact the registration desk for assistance.',
    validation: 'Some required information was missing or invalid. Please review your answers and try again.',
    storage_failed: 'We could not save your registration. Please try again.',
    upload_failed: 'File upload failed. Please check the file type and size (PDF, DOC, DOCX, JPG, or PNG; max 5 MB).',
    malformed: 'We could not process your submission. Please refresh the page and try again.',
    internal: 'We could not submit your registration. Please try again in a few minutes.'
  };

  function messageForCode(code) {
    return CODE_MESSAGES[code] || MSG_BACKEND;
  }

  // text/plain body = CORS-simple request, no preflight; Apps Script handles it via doPost
  function submitApplicant(fields, files, extras) {
    files = files || {};
    extras = extras || {};
    if (!CONFIG.APPS_SCRIPT_URL || CONFIG.APPS_SCRIPT_URL.indexOf('https://script.google.com/') !== 0) {
      return Promise.reject(ApiError(MSG_BACKEND));
    }

    return Promise.resolve()
      .then(function () { return files.resume ? fileToDataUrl(files.resume) : null; })
      .then(function (resume) { return files.pwdId ? fileToDataUrl(files.pwdId).then(function (p) { return [resume, p]; }) : [resume, null]; })
      .then(function (uploads) {
        var payload = {
          action: 'register',
          honeypot: String(extras.honeypot || ''),
          fields: fields,
          files: {
            resume: uploads[0] ? {
              filename: files.resume.name,
              mimeType: files.resume.type || guessMime(files.resume.name),
              dataUrl: uploads[0]
            } : null,
            pwdId: uploads[1] ? {
              filename: files.pwdId.name,
              mimeType: files.pwdId.type || guessMime(files.pwdId.name),
              dataUrl: uploads[1]
            } : null
          }
        };

        var controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
        var timer = controller
          ? setTimeout(function () { controller.abort(); }, CONFIG.REQUEST_TIMEOUT_MS)
          : null;

        return fetch(CONFIG.APPS_SCRIPT_URL, {
          method: 'POST',
          redirect: 'follow',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify(payload),
          signal: controller ? controller.signal : undefined
        }).then(function (res) {
          if (timer) clearTimeout(timer);
          return res.json().catch(function () { throw ApiError(MSG_BACKEND); }).then(function (data) {
            if (!res.ok) { console.error('Registration backend error', data); throw ApiError(MSG_BACKEND); }
            return normalizeResponse(data);
          });
        }, function (err) {
          if (timer) clearTimeout(timer);
          console.error(err);
          throw ApiError(err && err.name === 'AbortError' ? MSG_TIMEOUT : MSG_NETWORK);
        });
      });
  }

  function normalizeResponse(data) {
    if (data && data.status === 'success') {
      return {
        ok: true,
        duplicate: false,
        uniqueId: data.uniqueId || '',
        referenceNo: data.referenceNo || '',
        emailSent: !!data.emailSent
      };
    }
    if (data && data.status === 'duplicate') {
      // Minimal payload: reference number only — no internal UUID.
      return {
        ok: true,
        duplicate: true,
        referenceNo: data.referenceNo || '',
        message: data.message || ''
      };
    }
    console.error('Registration rejected by backend', data);
    throw ApiError(messageForCode(data && data.code));
  }

  function fileToDataUrl(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () { resolve(String(reader.result)); };
      reader.onerror = function () { reject(ApiError('The selected file could not be read. Please try a different file.')); };
      reader.readAsDataURL(file);
    });
  }

  function guessMime(name) {
    var ext = String(name).split('.').pop().toLowerCase();
    switch (ext) {
      case 'pdf': return 'application/pdf';
      case 'doc': return 'application/msword';
      case 'docx': return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
      case 'jpg':
      case 'jpeg': return 'image/jpeg';
      case 'png': return 'image/png';
      default: return 'application/octet-stream';
    }
  }

  /* ── Form Choices API ── */

  function jsonpGet(params, onDone) {
    var callbackName = 'fcCb_' + Date.now() + Math.floor(Math.random() * 1000);
    var script = document.createElement('script');
    var timer = setTimeout(function () {
      onDone({ status: 'error', message: 'Request timed out.' });
      delete window[callbackName]; script.remove();
    }, 10000);
    window[callbackName] = function (data) {
      clearTimeout(timer); onDone(data);
      delete window[callbackName]; script.remove();
    };
    script.onerror = function () {
      clearTimeout(timer);
      onDone({ status: 'error', message: 'Cannot reach configuration service.' });
      delete window[callbackName]; script.remove();
    };
    var query = Object.keys(params).map(function (k) { return k + '=' + encodeURIComponent(params[k]); }).join('&');
    script.src = CONFIG.APPS_SCRIPT_URL + '?' + query + '&callback=' + callbackName;
    document.body.appendChild(script);
  }

  // Public READ of active choices. onDone receives
  // {status:'success', fields:{...}, choices:{fieldKey:[{value,label}]}}
  // or {status:'error'} — callers must fall back to built-in defaults.
  function fetchFormChoices(onDone) {
    jsonpGet({ action: 'form_choices' }, onDone);
  }

  // Administrative mutation. Resolves the parsed response; rejects with a
  // user-facing message otherwise. The admin key travels only in the POST
  // body (never in a URL/query string).
  function adminFormChoice(action, data, adminKey) {
    var body = Object.assign({}, data || {}, { action: action, adminKey: String(adminKey || '') });
    return fetch(CONFIG.APPS_SCRIPT_URL, {
      method: 'POST',
      redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body)
    }).then(function (res) {
      return res.json().catch(function () { throw new Error('Unexpected server response.'); });
    }).then(function (out) {
      if (!res.ok) throw new Error(out && out.message ? out.message : 'Request failed.');
      if (out.status === 'error') throw new Error(out.message || 'Request failed.');
      return out;
    });
  }

  return {
    CONFIG: CONFIG,
    submitApplicant: submitApplicant,
    fetchFormChoices: fetchFormChoices,
    adminFormChoice: adminFormChoice,
    ApiError: ApiError
  };
})();
