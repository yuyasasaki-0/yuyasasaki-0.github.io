"use strict";
(function () {
  var root =
    typeof globalThis !== "undefined" ? globalThis : typeof window !== "undefined" ? window : undefined;
  if (!root) return;

  var ns = (root.MailChecker = root.MailChecker || {});
  var storage = ns.storage;

  function normalizeString(value) {
    return value == null ? "" : String(value).trim();
  }

  function lower(value) {
    return value == null ? "" : String(value).toLowerCase();
  }

  function withTimeout(promise, ms, timeoutValue) {
    return new Promise(function (resolve) {
      var done = false;
      var id = null;

      try {
        if (typeof ms === "number" && isFinite(ms) && ms > 0) {
          id = setTimeout(function () {
            if (done) return;
            done = true;
            resolve(timeoutValue);
          }, ms);
        }
      } catch (_e) {}

      Promise.resolve(promise)
        .then(function (value) {
          if (done) return;
          done = true;
          try {
            if (id != null) clearTimeout(id);
          } catch (_e2) {}
          resolve(value);
        })
        .catch(function (_err) {
          if (done) return;
          done = true;
          try {
            if (id != null) clearTimeout(id);
          } catch (_e3) {}
          resolve(timeoutValue);
        });
    });
  }

  function canUseEws() {
    try {
      return (
        typeof Office !== "undefined" &&
        Office &&
        Office.context &&
        Office.context.mailbox &&
        typeof Office.context.mailbox.makeEwsRequestAsync === "function"
      );
    } catch (_e) {
      return false;
    }
  }

  function pMakeEwsRequest(requestXml) {
    return new Promise(function (resolve) {
      try {
        if (!canUseEws()) return resolve({ ok: false, error: "EWS is not available in this client." });
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(String(requestXml || ""), function (asyncResult) {
          try {
            if (!asyncResult || asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
              var msg = "";
              try {
                msg = asyncResult && asyncResult.error && asyncResult.error.message ? asyncResult.error.message : "";
              } catch (_e2) {}
              return resolve({ ok: false, error: msg || "EWS request failed." });
            }
            resolve({ ok: true, value: String(asyncResult.value || "") });
          } catch (_e3) {
            resolve({ ok: false, error: "EWS request failed." });
          }
        });
      } catch (_e) {
        resolve({ ok: false, error: "EWS request failed." });
      }
    });
  }

  function parseXml(xmlText) {
    try {
      if (!xmlText) return null;
      var doc = new DOMParser().parseFromString(String(xmlText), "text/xml");
      return doc;
    } catch (_e) {
      return null;
    }
  }

  function findFirstByLocalName(rootNode, localName) {
    if (!rootNode) return null;
    try {
      var all = rootNode.getElementsByTagName("*");
      for (var i = 0; i < all.length; i++) {
        if (all[i] && all[i].localName === localName) return all[i];
      }
    } catch (_e) {}
    return null;
  }

  function findAllByLocalName(rootNode, localName) {
    var out = [];
    if (!rootNode) return out;
    try {
      var all = rootNode.getElementsByTagName("*");
      for (var i = 0; i < all.length; i++) {
        if (all[i] && all[i].localName === localName) out.push(all[i]);
      }
    } catch (_e) {}
    return out;
  }

  function textContent(node) {
    try {
      return node && node.textContent != null ? String(node.textContent) : "";
    } catch (_e) {
      return "";
    }
  }

  function getEwsResponseCode(doc) {
    if (!doc) return "";
    var node = findFirstByLocalName(doc, "ResponseCode");
    return normalizeString(textContent(node));
  }

  function getEwsResponseClass(doc) {
    if (!doc) return "";
    var msg = findFirstByLocalName(doc, "ResponseMessage");
    try {
      if (!msg || !msg.getAttribute) return "";
      return normalizeString(msg.getAttribute("ResponseClass"));
    } catch (_e) {
      return "";
    }
  }

  function extractEwsMessageText(doc) {
    if (!doc) return "";
    var msgText = findFirstByLocalName(doc, "MessageText");
    return normalizeString(textContent(msgText));
  }

  function isEwsSuccess(doc) {
    var rc = lower(getEwsResponseClass(doc));
    if (rc === "success") return true;
    var code = lower(getEwsResponseCode(doc));
    if (code === "noerror") return true;
    return false;
  }

  function buildSoapEnvelope(bodyInnerXml) {
    return (
      '<?xml version="1.0" encoding="utf-8"?>' +
      '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
      'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
      'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
      'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
      "<soap:Header>" +
      '<t:RequestServerVersion Version="Exchange2013" />' +
      "</soap:Header>" +
      "<soap:Body>" +
      String(bodyInnerXml || "") +
      "</soap:Body>" +
      "</soap:Envelope>"
    );
  }

  function buildExpandDlRequest(emailAddress) {
    var email = normalizeString(emailAddress);
    return buildSoapEnvelope(
      "<m:ExpandDL>" +
        "<m:Mailbox>" +
        "<t:EmailAddress>" +
        escapeXml(email) +
        "</t:EmailAddress>" +
        "</m:Mailbox>" +
        "</m:ExpandDL>"
    );
  }

  function buildResolveNamesRequest(unresolvedEntry, searchScope) {
    var entry = normalizeString(unresolvedEntry);
    var scope = normalizeString(searchScope) || "Contacts";
    return buildSoapEnvelope(
      '<m:ResolveNames ReturnFullContactData="true" SearchScope="' +
        escapeXmlAttribute(scope) +
        '">' +
        "<m:UnresolvedEntry>" +
        escapeXml(entry) +
        "</m:UnresolvedEntry>" +
        "</m:ResolveNames>"
    );
  }

  function escapeXml(text) {
    return String(text || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function escapeXmlAttribute(text) {
    return escapeXml(text).replace(/\r?\n/g, " ");
  }

  function parseDlExpansion(doc) {
    var out = [];
    if (!doc) return out;
    var dl = findFirstByLocalName(doc, "DLExpansion");
    if (!dl) return out;

    var mailboxes = findAllByLocalName(dl, "Mailbox");
    for (var i = 0; i < mailboxes.length; i++) {
      var mb = mailboxes[i];
      if (!mb) continue;
      var name = normalizeString(textContent(findFirstByLocalName(mb, "Name")));
      var email = normalizeString(textContent(findFirstByLocalName(mb, "EmailAddress")));
      var mailboxType = normalizeString(textContent(findFirstByLocalName(mb, "MailboxType")));
      if (!email) continue;
      out.push({ emailAddress: email, displayName: name, mailboxType: mailboxType });
    }

    return out;
  }

  function parseResolveNames(doc) {
    if (!doc) return { resolutions: [] };
    var rs = findFirstByLocalName(doc, "ResolutionSet");
    if (!rs) return { resolutions: [] };

    var resolutions = [];
    var resNodes = findAllByLocalName(rs, "Resolution");
    for (var i = 0; i < resNodes.length; i++) {
      var res = resNodes[i];
      if (!res) continue;

      var mailbox = findFirstByLocalName(res, "Mailbox");
      var name = normalizeString(textContent(findFirstByLocalName(mailbox, "Name")));
      var email = normalizeString(textContent(findFirstByLocalName(mailbox, "EmailAddress")));
      var mailboxType = normalizeString(textContent(findFirstByLocalName(mailbox, "MailboxType")));

      resolutions.push({ emailAddress: email, displayName: name, mailboxType: mailboxType });
    }

    return { resolutions: resolutions };
  }

  var CONTACTS_CACHE_KEY = "mailchecker.cache.contacts.v1";
  var DL_CACHE_KEY = "mailchecker.cache.dl.v1";

  function nowMs() {
    try {
      return Date.now();
    } catch (_e) {
      return 0;
    }
  }

  async function getCacheDoc(key) {
    if (!storage || typeof storage.getJson !== "function") return { schemaVersion: 1, entries: {} };
    var doc = null;
    try {
      doc = await storage.getJson(key);
    } catch (_e) {
      doc = null;
    }
    if (!doc || typeof doc !== "object") return { schemaVersion: 1, entries: {} };
    if (!doc.entries || typeof doc.entries !== "object") doc.entries = {};
    return doc;
  }

  async function setCacheDoc(key, doc) {
    if (!storage || typeof storage.setJson !== "function") return;
    try {
      await storage.setJson(key, doc || null);
    } catch (_e) {}
  }

  function isFresh(entry, ttlMs) {
    if (!entry || typeof entry !== "object") return false;
    if (typeof entry.ts !== "number") return false;
    if (ttlMs <= 0) return false;
    return nowMs() - entry.ts < ttlMs;
  }

  async function expandDlCached(emailAddress, options) {
    var email = normalizeString(emailAddress);
    if (!email) return { ok: false, members: [], error: "Missing email address." };

    var ttlMs = 24 * 60 * 60 * 1000; // 1 day
    var timeoutMs = 1200;
    try {
      if (options && typeof options.ttlMs === "number") ttlMs = options.ttlMs;
      if (options && typeof options.timeoutMs === "number") timeoutMs = options.timeoutMs;
    } catch (_e) {}

    var key = lower(email);
    var doc = await getCacheDoc(DL_CACHE_KEY);
    var cached = doc.entries[key];
    if (isFresh(cached, ttlMs) && Array.isArray(cached.members)) {
      return { ok: true, members: cached.members, cached: true };
    }

    var requestXml = buildExpandDlRequest(email);
    var response = await withTimeout(pMakeEwsRequest(requestXml), timeoutMs, { ok: false, error: "Timeout" });
    if (!response || !response.ok) {
      return { ok: false, members: [], error: (response && response.error) || "EWS failed." };
    }

    var docXml = parseXml(response.value);
    if (!docXml) return { ok: false, members: [], error: "Failed to parse EWS response." };
    if (!isEwsSuccess(docXml)) {
      return {
        ok: false,
        members: [],
        error: getEwsResponseCode(docXml) || extractEwsMessageText(docXml) || "EWS error.",
      };
    }

    var members = parseDlExpansion(docXml);
    doc.entries[key] = { ts: nowMs(), members: members };
    await setCacheDoc(DL_CACHE_KEY, doc);
    return { ok: true, members: members, cached: false };
  }

  async function resolveInContactsCached(emailAddress, options) {
    var email = normalizeString(emailAddress);
    if (!email) return { ok: false, value: null, error: "Missing email address." };

    var ttlMs = 7 * 24 * 60 * 60 * 1000; // 7 days
    var timeoutMs = 900;
    try {
      if (options && typeof options.ttlMs === "number") ttlMs = options.ttlMs;
      if (options && typeof options.timeoutMs === "number") timeoutMs = options.timeoutMs;
    } catch (_e) {}

    var key = lower(email);
    var doc = await getCacheDoc(CONTACTS_CACHE_KEY);
    var cached = doc.entries[key];
    if (isFresh(cached, ttlMs) && typeof cached.value === "boolean") {
      return { ok: true, value: cached.value, cached: true };
    }

    var requestXml = buildResolveNamesRequest(email, "Contacts");
    var response = await withTimeout(pMakeEwsRequest(requestXml), timeoutMs, { ok: false, error: "Timeout" });
    if (!response || !response.ok) {
      return { ok: false, value: null, error: (response && response.error) || "EWS failed." };
    }

    var docXml = parseXml(response.value);
    if (!docXml) return { ok: false, value: null, error: "Failed to parse EWS response." };
    if (!isEwsSuccess(docXml)) {
      return {
        ok: false,
        value: null,
        error: getEwsResponseCode(docXml) || extractEwsMessageText(docXml) || "EWS error.",
      };
    }

    var parsed = parseResolveNames(docXml);
    var found = Array.isArray(parsed.resolutions) && parsed.resolutions.length > 0;
    doc.entries[key] = { ts: nowMs(), value: found };
    await setCacheDoc(CONTACTS_CACHE_KEY, doc);
    return { ok: true, value: found, cached: false };
  }

  ns.ews = {
    canUseEws: canUseEws,
    _pMakeEwsRequest: pMakeEwsRequest,
    expandDlCached: expandDlCached,
    resolveInContactsCached: resolveInContactsCached,
  };
})();

