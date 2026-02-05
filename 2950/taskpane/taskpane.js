"use strict";
(function () {
  function $(id) {
    return document.getElementById(id);
  }

  function normalizeString(value) {
    return value == null ? "" : String(value).trim();
  }

  function lower(value) {
    return value == null ? "" : String(value).toLowerCase();
  }

  function startsWith(value, prefix) {
    return String(value || "").slice(0, String(prefix || "").length) === String(prefix || "");
  }

  function endsWith(value, suffix) {
    var s = String(value || "");
    var suf = String(suffix || "");
    if (!suf) return false;
    return s.slice(s.length - suf.length) === suf;
  }

  function domainSuffixFromAddress(address) {
    var a = normalizeString(address);
    if (!a) return "";
    var at = a.lastIndexOf("@");
    if (at < 0) return "";
    return lower(a.slice(at)); // includes "@"
  }

  function normalizeDomainSuffix(value) {
    var v = lower(normalizeString(value));
    if (!v) return "";
    if (v.indexOf("@") === 0) return v;
    if (v.indexOf("@") >= 0) return v.slice(v.lastIndexOf("@"));
    return "@" + v;
  }

  function buildInternalDomains(settings, senderDomainSuffix) {
    var list = [];
    try {
      if (settings && Array.isArray(settings.internalDomains)) {
        for (var i = 0; i < settings.internalDomains.length; i++) {
          var row = settings.internalDomains[i];
          if (!row) continue;
          var d = normalizeString(row.domain != null ? row.domain : row.Domain);
          var suf = normalizeDomainSuffix(d);
          if (suf) list.push(suf);
        }
      }
    } catch (_e) {}
    if (senderDomainSuffix) list.push(senderDomainSuffix);
    var seen = {};
    var out = [];
    for (var j = 0; j < list.length; j++) {
      var k = lower(list[j]);
      if (!k || seen[k]) continue;
      seen[k] = true;
      out.push(k);
    }
    return out;
  }

  function isInternalAddress(address, internalDomains) {
    var a = lower(normalizeString(address));
    if (!a) return false;
    if (!Array.isArray(internalDomains) || internalDomains.length === 0) return false;
    for (var i = 0; i < internalDomains.length; i++) {
      var suf = internalDomains[i];
      if (!suf) continue;
      if (endsWith(a, suf)) return true;
    }
    return false;
  }

  function isDistributionListRecipient(recipient) {
    if (!recipient) return false;
    try {
      var rt = recipient.recipientType != null ? recipient.recipientType : recipient.RecipientType;
      if (rt == null) return false;
      var s = lower(rt);
      return s.indexOf("distribution") >= 0 || s.indexOf("group") >= 0;
    } catch (_e) {
      return false;
    }
  }

  async function enrichSnapshotWithEws(snapshot, settings) {
    if (!snapshot || !settings) return snapshot;

    var g = (settings && settings.general) || {};
    var wantDl = !!g.enableGetContactGroupMembers || !!g.enableGetExchangeDistributionListMembers;
    var wantContacts =
      !!g.isAutoCheckRegisteredInContacts ||
      !!g.isWarningIfRecipientsIsNotRegistered ||
      !!g.isProhibitsSendingMailIfRecipientsIsNotRegistered;

    if (!wantDl && !wantContacts) return snapshot;

    var ews = MailChecker && MailChecker.ews;
    var canEws = !!(ews && typeof ews.canUseEws === "function" && ews.canUseEws());

    var resolved = { expandedGroups: [], contacts: null, contactsLookupFailed: false };

    if (!canEws) {
      if (wantContacts) resolved.contactsLookupFailed = true;
      snapshot.resolved = resolved;
      return snapshot;
    }

    var internalDomains = buildInternalDomains(settings, domainSuffixFromAddress(snapshot.senderEmailAddress));

    function collectDlCandidates(list, field, outList, seen) {
      for (var i = 0; i < list.length; i++) {
        var r = list[i];
        if (!r) continue;
        var email = normalizeString(r.emailAddress);
        if (!email || email.indexOf("@") < 0) continue;
        var key = lower(email);
        if (seen[key]) continue;
        seen[key] = true;
        if (!isDistributionListRecipient(r)) continue;
        if (internalDomains.length > 0 && !isInternalAddress(email, internalDomains)) continue;
        outList.push({ emailAddress: email, displayName: normalizeString(r.displayName), field: field });
      }
    }

    if (wantDl && typeof ews.expandDlCached === "function") {
      try {
        var candidates = [];
        var seenDl = {};
        collectDlCandidates((snapshot.recipients && snapshot.recipients.to) || [], "To", candidates, seenDl);
        collectDlCandidates((snapshot.recipients && snapshot.recipients.cc) || [], "Cc", candidates, seenDl);
        collectDlCandidates((snapshot.recipients && snapshot.recipients.bcc) || [], "Bcc", candidates, seenDl);

        for (var c = 0; c < candidates.length && c < 5; c++) {
          var cand = candidates[c];
          var r1 = await ews.expandDlCached(cand.emailAddress, { timeoutMs: 2500 });
          if (!r1 || !r1.ok || !Array.isArray(r1.members) || r1.members.length === 0) continue;

          var members = r1.members
            .map(function (m) {
              return {
                emailAddress: normalizeString(m && m.emailAddress),
                displayName: normalizeString(m && m.displayName),
              };
            })
            .filter(function (m) {
              return m.emailAddress && m.emailAddress.indexOf("@") >= 0;
            });

          if (members.length === 0) continue;

          resolved.expandedGroups.push({
            emailAddress: cand.emailAddress,
            displayName: cand.displayName,
            field: cand.field,
            members: members,
          });
        }
      } catch (_e2) {}
    }

    if (wantContacts && typeof ews.resolveInContactsCached === "function") {
      try {
        var emails = [];
        function addRecipients(list) {
          for (var i = 0; i < list.length; i++) {
            var r = list[i];
            if (!r) continue;
            var email = normalizeString(r.emailAddress);
            if (!email || email.indexOf("@") < 0) continue;
            if (internalDomains.length > 0 && isInternalAddress(email, internalDomains)) continue;
            emails.push(email);
          }
        }

        addRecipients((snapshot.recipients && snapshot.recipients.to) || []);
        addRecipients((snapshot.recipients && snapshot.recipients.cc) || []);
        addRecipients((snapshot.recipients && snapshot.recipients.bcc) || []);

        for (var g2 = 0; g2 < resolved.expandedGroups.length; g2++) {
          var grp = resolved.expandedGroups[g2];
          var members = (grp && grp.members) || [];
          for (var m2 = 0; m2 < members.length; m2++) {
            var em = normalizeString(members[m2] && members[m2].emailAddress);
            if (!em || em.indexOf("@") < 0) continue;
            if (internalDomains.length > 0 && isInternalAddress(em, internalDomains)) continue;
            emails.push(em);
          }
        }

        var seen = {};
        var uniq = [];
        for (var u = 0; u < emails.length; u++) {
          var k2 = lower(emails[u]);
          if (!k2 || seen[k2]) continue;
          seen[k2] = true;
          uniq.push(emails[u]);
        }

        var contacts = {};
        for (var q = 0; q < uniq.length && q < 40; q++) {
          var email2 = uniq[q];
          var r2 = await ews.resolveInContactsCached(email2, { timeoutMs: 2000 });
          if (r2 && r2.ok && typeof r2.value === "boolean") {
            contacts[lower(email2)] = r2.value;
          } else if (r2 && !r2.ok) {
            resolved.contactsLookupFailed = true;
          }
        }

        resolved.contacts = contacts;
      } catch (_e3) {
        resolved.contactsLookupFailed = true;
        resolved.contacts = null;
      }
    }

    snapshot.resolved = resolved;
    return snapshot;
  }

  function esc(text) {
    return String(text || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function setStatus(text) {
    var el = $("status");
    if (!el) return;
    el.textContent = String(text || "");
  }

  function pickLocale(settings) {
    var lang = "";
    try {
      lang = normalizeString(settings && settings.general && settings.general.languageCode);
    } catch (_e) {}
    if (lang) return lang;
    try {
      return normalizeString(Office.context && Office.context.displayLanguage) || "en-US";
    } catch (_e2) {
      return "en-US";
    }
  }

  function t(locale, jaText, enText) {
    return startsWith(lower(locale), "ja") ? jaText : enText;
  }

  function parseLines(text) {
    return String(text || "")
      .split(/\r?\n/)
      .map(function (l) {
        return l.trim();
      })
      .filter(Boolean);
  }

  function parseWhitelist(text) {
    var lines = parseLines(text);
    var out = [];
    for (var i = 0; i < lines.length; i++) {
      var parts = lines[i].split(",");
      var whiteName = normalizeString(parts[0]);
      if (!whiteName) continue;
      var skipRaw = lower(normalizeString(parts[1]));
      var isSkip = skipRaw === "y" || skipRaw === "yes" || skipRaw === "true" || skipRaw === "1";
      out.push({ whiteName: whiteName, isSkipConfirmation: isSkip });
    }
    return out;
  }

  function whitelistToText(list) {
    if (!Array.isArray(list)) return "";
    return list
      .map(function (w) {
        if (!w || !w.whiteName) return "";
        return w.whiteName + (w.isSkipConfirmation ? ",Y" : ",N");
      })
      .filter(Boolean)
      .join("\n");
  }

  function domainsToText(list) {
    if (!Array.isArray(list)) return "";
    return list
      .map(function (d) {
        return d && d.domain ? d.domain : "";
      })
      .filter(Boolean)
      .join("\n");
  }

  function parseDomains(text) {
    return parseLines(text).map(function (d) {
      return { domain: d };
    });
  }

  function renderTabs() {
    var buttons = Array.prototype.slice.call(document.querySelectorAll(".tab"));
    buttons.forEach(function (btn) {
      btn.addEventListener("click", function () {
        var tab = btn.getAttribute("data-tab");
        buttons.forEach(function (b) {
          b.classList.toggle("active", b === btn);
          b.setAttribute("aria-selected", b === btn ? "true" : "false");
        });
        Array.prototype.slice.call(document.querySelectorAll(".tab-panel")).forEach(function (panel) {
          panel.classList.toggle("active", panel.id === "tab-" + tab);
        });
      });
    });
  }

  function getCheckboxValue(id) {
    var el = $(id);
    return !!(el && el.checked);
  }

  function setCheckboxValue(id, value) {
    var el = $(id);
    if (!el) return;
    el.checked = !!value;
  }

  function getTextValue(id) {
    var el = $(id);
    return el ? normalizeString(el.value) : "";
  }

  function setTextValue(id, value) {
    var el = $(id);
    if (!el) return;
    el.value = value == null ? "" : String(value);
  }

  function getNumberValue(id, fallback) {
    var el = $(id);
    if (!el) return fallback;
    var n = parseInt(String(el.value || ""), 10);
    return isFinite(n) ? n : fallback;
  }

  function setNumberValue(id, value) {
    var el = $(id);
    if (!el) return;
    el.value = value == null ? "" : String(value);
  }

  function renderSettingsForm(settings) {
    var g = (settings && settings.general) || {};
    var ex = (settings && settings.externalDomains) || {};
    var fb = (settings && settings.forceAutoChangeRecipientsToBcc) || {};
    var att = (settings && settings.attachmentsSetting) || {};
    var sec = (settings && settings.securityForReceivedMail) || {};

    setCheckboxValue("g-enableForgottenToAttachAlert", g.enableForgottenToAttachAlert);
    setCheckboxValue(
      "g-isDoNotConfirmationIfAllRecipientsAreSameDomain",
      g.isDoNotConfirmationIfAllRecipientsAreSameDomain
    );
    setCheckboxValue("g-isShowConfirmationToMultipleDomain", g.isShowConfirmationToMultipleDomain);
    setCheckboxValue("g-isDoDoNotConfirmationIfAllWhite", g.isDoDoNotConfirmationIfAllWhite);
    setCheckboxValue("g-isAutoCheckIfAllRecipientsAreSameDomain", g.isAutoCheckIfAllRecipientsAreSameDomain);
    setCheckboxValue("g-isEnableRecipientsAreSortedByDomain", g.isEnableRecipientsAreSortedByDomain);
    setCheckboxValue("g-isAutoAddSenderToCc", g.isAutoAddSenderToCc);
    setCheckboxValue("g-isAutoAddSenderToBcc", g.isAutoAddSenderToBcc);
    setCheckboxValue("g-isAutoCheckAttachments", g.isAutoCheckAttachments);
    setCheckboxValue(
      "g-isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles",
      g.isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles
    );

    setCheckboxValue("g-enableGetExchangeDistributionListMembers", g.enableGetExchangeDistributionListMembers);
    setCheckboxValue("g-exchangeDistributionListMembersAreWhite", g.exchangeDistributionListMembersAreWhite);
    setCheckboxValue("g-enableGetContactGroupMembers", g.enableGetContactGroupMembers);
    setCheckboxValue("g-contactGroupMembersAreWhite", g.contactGroupMembersAreWhite);
    setCheckboxValue("g-isAutoCheckRegisteredInContacts", g.isAutoCheckRegisteredInContacts);
    setCheckboxValue("g-isWarningIfRecipientsIsNotRegistered", g.isWarningIfRecipientsIsNotRegistered);
    setCheckboxValue(
      "g-isProhibitsSendingMailIfRecipientsIsNotRegistered",
      g.isProhibitsSendingMailIfRecipientsIsNotRegistered
    );

    setNumberValue("ex-target", ex.targetToAndCcExternalDomainsNum);
    setCheckboxValue("ex-warn", ex.isWarningWhenLargeNumberOfExternalDomains);
    setCheckboxValue("ex-prohibit", ex.isProhibitedWhenLargeNumberOfExternalDomains);
    setCheckboxValue("ex-autoBcc", ex.isAutoChangeToBccWhenLargeNumberOfExternalDomains);

    setCheckboxValue("fb-enable", fb.isForceAutoChangeRecipientsToBcc);
    setTextValue("fb-toRecipient", fb.toRecipient);
    setCheckboxValue("fb-includeInternal", fb.isIncludeInternalDomain);

    setCheckboxValue("att-warnEncryptedZip", att.isWarningWhenEncryptedZipIsAttached);
    setCheckboxValue("att-prohibitEncryptedZip", att.isProhibitedWhenEncryptedZipIsAttached);
    setCheckboxValue("att-prohibitAll", att.isAttachmentsProhibited);
    setCheckboxValue("att-warnRealFile", att.isWarningWhenAttachedRealFile);

    setCheckboxValue("sec-enable", sec.isEnableSecurityForReceivedMail);
    setCheckboxValue("sec-subjectKeywords", sec.isEnableAlertKeywordOfSubjectWhenOpeningMailsData);
    setCheckboxValue("sec-headerAnalysis", sec.isEnableMailHeaderAnalysis);
    setCheckboxValue("sec-warnSpf", sec.isShowWarningWhenSpfFails);
    setCheckboxValue("sec-warnDkim", sec.isShowWarningWhenDkimFails);
    setCheckboxValue("sec-attachments", sec.isEnableWarningFeatureWhenOpeningAttachments);
    setCheckboxValue("sec-warnBeforeOpen", sec.isWarnBeforeOpeningAttachments);
    setCheckboxValue("sec-warnEncryptedZip", sec.isWarnBeforeOpeningEncryptedZip);
    setCheckboxValue("sec-warnZipLnk", sec.isWarnLinkFileInTheZip);
    setCheckboxValue("sec-warnZipOne", sec.isWarnOneFileInTheZip);
    setCheckboxValue("sec-warnZipMacroOffice", sec.isWarnOfficeFileWithMacroInTheZip);
    setCheckboxValue("sec-warnMacroOffice", sec.isWarnBeforeOpeningAttachmentsThatContainMacros);
    setCheckboxValue("sec-spoofing", sec.isShowWarningWhenSpoofingRisk);
    setCheckboxValue("sec-dmarcNotImpl", sec.isShowWarningWhenDmarcNotImplemented);

    setTextValue("list-internalDomains", domainsToText(settings && settings.internalDomains));
    setTextValue("list-whitelist", whitelistToText(settings && settings.whitelist));
  }

  function applySettingsFromForm(settings) {
    var next = JSON.parse(JSON.stringify(settings || {}));

    next.general = next.general || {};
    next.general.enableForgottenToAttachAlert = getCheckboxValue("g-enableForgottenToAttachAlert");
    next.general.isDoNotConfirmationIfAllRecipientsAreSameDomain = getCheckboxValue(
      "g-isDoNotConfirmationIfAllRecipientsAreSameDomain"
    );
    next.general.isShowConfirmationToMultipleDomain = getCheckboxValue("g-isShowConfirmationToMultipleDomain");
    next.general.isDoDoNotConfirmationIfAllWhite = getCheckboxValue("g-isDoDoNotConfirmationIfAllWhite");
    next.general.isAutoCheckIfAllRecipientsAreSameDomain = getCheckboxValue("g-isAutoCheckIfAllRecipientsAreSameDomain");
    next.general.isEnableRecipientsAreSortedByDomain = getCheckboxValue("g-isEnableRecipientsAreSortedByDomain");
    next.general.isAutoAddSenderToCc = getCheckboxValue("g-isAutoAddSenderToCc");
    next.general.isAutoAddSenderToBcc = getCheckboxValue("g-isAutoAddSenderToBcc");
    next.general.isAutoCheckAttachments = getCheckboxValue("g-isAutoCheckAttachments");
    next.general.isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles = getCheckboxValue(
      "g-isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles"
    );

    next.general.enableGetExchangeDistributionListMembers = getCheckboxValue("g-enableGetExchangeDistributionListMembers");
    next.general.exchangeDistributionListMembersAreWhite = getCheckboxValue("g-exchangeDistributionListMembersAreWhite");
    next.general.enableGetContactGroupMembers = getCheckboxValue("g-enableGetContactGroupMembers");
    next.general.contactGroupMembersAreWhite = getCheckboxValue("g-contactGroupMembersAreWhite");
    next.general.isAutoCheckRegisteredInContacts = getCheckboxValue("g-isAutoCheckRegisteredInContacts");
    next.general.isWarningIfRecipientsIsNotRegistered = getCheckboxValue("g-isWarningIfRecipientsIsNotRegistered");
    next.general.isProhibitsSendingMailIfRecipientsIsNotRegistered = getCheckboxValue(
      "g-isProhibitsSendingMailIfRecipientsIsNotRegistered"
    );

    next.externalDomains = next.externalDomains || {};
    next.externalDomains.targetToAndCcExternalDomainsNum = getNumberValue("ex-target", 10);
    next.externalDomains.isWarningWhenLargeNumberOfExternalDomains = getCheckboxValue("ex-warn");
    next.externalDomains.isProhibitedWhenLargeNumberOfExternalDomains = getCheckboxValue("ex-prohibit");
    next.externalDomains.isAutoChangeToBccWhenLargeNumberOfExternalDomains = getCheckboxValue("ex-autoBcc");

    next.forceAutoChangeRecipientsToBcc = next.forceAutoChangeRecipientsToBcc || {};
    next.forceAutoChangeRecipientsToBcc.isForceAutoChangeRecipientsToBcc = getCheckboxValue("fb-enable");
    next.forceAutoChangeRecipientsToBcc.toRecipient = getTextValue("fb-toRecipient");
    next.forceAutoChangeRecipientsToBcc.isIncludeInternalDomain = getCheckboxValue("fb-includeInternal");

    next.attachmentsSetting = next.attachmentsSetting || {};
    next.attachmentsSetting.isWarningWhenEncryptedZipIsAttached = getCheckboxValue("att-warnEncryptedZip");
    next.attachmentsSetting.isProhibitedWhenEncryptedZipIsAttached = getCheckboxValue("att-prohibitEncryptedZip");
    next.attachmentsSetting.isAttachmentsProhibited = getCheckboxValue("att-prohibitAll");
    next.attachmentsSetting.isWarningWhenAttachedRealFile = getCheckboxValue("att-warnRealFile");

    next.securityForReceivedMail = next.securityForReceivedMail || {};
    next.securityForReceivedMail.isEnableSecurityForReceivedMail = getCheckboxValue("sec-enable");
    next.securityForReceivedMail.isEnableAlertKeywordOfSubjectWhenOpeningMailsData = getCheckboxValue("sec-subjectKeywords");
    next.securityForReceivedMail.isEnableMailHeaderAnalysis = getCheckboxValue("sec-headerAnalysis");
    next.securityForReceivedMail.isShowWarningWhenSpfFails = getCheckboxValue("sec-warnSpf");
    next.securityForReceivedMail.isShowWarningWhenDkimFails = getCheckboxValue("sec-warnDkim");
    next.securityForReceivedMail.isEnableWarningFeatureWhenOpeningAttachments = getCheckboxValue("sec-attachments");
    next.securityForReceivedMail.isWarnBeforeOpeningAttachments = getCheckboxValue("sec-warnBeforeOpen");
    next.securityForReceivedMail.isWarnBeforeOpeningEncryptedZip = getCheckboxValue("sec-warnEncryptedZip");
    next.securityForReceivedMail.isWarnLinkFileInTheZip = getCheckboxValue("sec-warnZipLnk");
    next.securityForReceivedMail.isWarnOneFileInTheZip = getCheckboxValue("sec-warnZipOne");
    next.securityForReceivedMail.isWarnOfficeFileWithMacroInTheZip = getCheckboxValue("sec-warnZipMacroOffice");
    next.securityForReceivedMail.isWarnBeforeOpeningAttachmentsThatContainMacros = getCheckboxValue("sec-warnMacroOffice");
    next.securityForReceivedMail.isShowWarningWhenSpoofingRisk = getCheckboxValue("sec-spoofing");
    next.securityForReceivedMail.isShowWarningWhenDmarcNotImplemented = getCheckboxValue("sec-dmarcNotImpl");

    next.internalDomains = parseDomains(getTextValue("list-internalDomains"));
    next.whitelist = parseWhitelist(getTextValue("list-whitelist"));

    return next;
  }

  function downloadText(filename, text, mimeType) {
    try {
      var type = normalizeString(mimeType) || "text/plain;charset=utf-8";
      var blob = new Blob([String(text || "")], { type: type });
      if (window.navigator && typeof window.navigator.msSaveBlob === "function") {
        window.navigator.msSaveBlob(blob, filename);
        return;
      }
      var a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      setTimeout(function () {
        try {
          URL.revokeObjectURL(a.href);
          document.body.removeChild(a);
        } catch (_e) {}
      }, 0);
    } catch (_e2) {}
  }

  function parseBool(value) {
    if (value == null) return null;
    var v = lower(normalizeString(value));
    if (!v) return null;
    if (v === "yes" || v === "y" || v === "true" || v === "1") return true;
    if (v === "no" || v === "n" || v === "false" || v === "0") return false;
    return null;
  }

  function boolToYesNo(value) {
    return value ? "Yes" : "No";
  }

  function parseIntOrNull(value) {
    if (value == null) return null;
    var s = normalizeString(value);
    if (!s) return null;
    var n = parseInt(s, 10);
    return isFinite(n) ? n : null;
  }

  function parseCsv(text) {
    var s = String(text || "");
    var rows = [];
    var row = [];
    var field = "";
    var inQuotes = false;

    for (var i = 0; i < s.length; i++) {
      var ch = s.charAt(i);

      if (inQuotes) {
        if (ch === "\"") {
          if (s.charAt(i + 1) === "\"") {
            field += "\"";
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          field += ch;
        }
        continue;
      }

      if (ch === "\"") {
        inQuotes = true;
        continue;
      }

      if (ch === ",") {
        row.push(field);
        field = "";
        continue;
      }

      if (ch === "\r") {
        if (s.charAt(i + 1) === "\n") i++;
        row.push(field);
        field = "";
        if (!(row.length === 1 && !normalizeString(row[0]))) rows.push(row);
        row = [];
        continue;
      }

      if (ch === "\n") {
        row.push(field);
        field = "";
        if (!(row.length === 1 && !normalizeString(row[0]))) rows.push(row);
        row = [];
        continue;
      }

      field += ch;
    }

    row.push(field);
    if (!(row.length === 1 && !normalizeString(row[0]))) rows.push(row);

    return rows;
  }

  function csvEscape(value) {
    var s = value == null ? "" : String(value);
    if (/[\",\r\n]/.test(s)) return "\"" + s.replace(/\"/g, "\"\"") + "\"";
    return s;
  }

  function toCsv(rows) {
    if (!Array.isArray(rows) || rows.length === 0) return "";
    return (
      rows
        .map(function (row) {
          var r = Array.isArray(row) ? row : [];
          return r
            .map(function (cell) {
              return csvEscape(cell);
            })
            .join(",");
        })
        .join("\r\n") + "\r\n"
    );
  }

  function getBaseName(fileName) {
    var n = String(fileName || "");
    n = n.replace(/^.*[\\/]/, "");
    return n;
  }

  function applyOutlookOkanCsvFiles(settings, fileEntries) {
    var next = JSON.parse(JSON.stringify(settings || {}));
    var summary = [];

    function setIf(rows, idx, setter) {
      if (!rows || rows.length === 0) return;
      var v = rows[0] && rows[0].length > idx ? rows[0][idx] : null;
      if (v == null) return;
      if (!normalizeString(v)) return;
      setter(v);
    }

    function applyList(rows, mapper) {
      var out = [];
      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        if (!row) continue;
        var item = mapper(row);
        if (item) out.push(item);
      }
      return out;
    }

    next.general = next.general || {};
    next.externalDomains = next.externalDomains || {};
    next.forceAutoChangeRecipientsToBcc = next.forceAutoChangeRecipientsToBcc || {};
    next.attachmentsSetting = next.attachmentsSetting || {};
    next.autoAddMessage = next.autoAddMessage || {};
    next.securityForReceivedMail = next.securityForReceivedMail || {};

    var ignored = [];

    for (var f = 0; f < fileEntries.length; f++) {
      var fe = fileEntries[f];
      if (!fe || !fe.name) continue;
      var base = lower(getBaseName(fe.name));
      var rows = parseCsv(String(fe.text || ""));

      switch (base) {
        case "generalsetting.csv": {
          if (rows.length === 0) break;
          var r0 = rows[0];
          function b(i, key) {
            if (r0.length <= i) return;
            var v = parseBool(r0[i]);
            if (v == null) return;
            next.general[key] = v;
          }
          function s(i, key) {
            if (r0.length <= i) return;
            var v = normalizeString(r0[i]);
            if (v === "") return;
            next.general[key] = v;
          }

          b(0, "isDoNotConfirmationIfAllRecipientsAreSameDomain");
          b(1, "isDoDoNotConfirmationIfAllWhite");
          b(2, "isAutoCheckIfAllRecipientsAreSameDomain");
          s(3, "languageCode");
          b(4, "isShowConfirmationToMultipleDomain");
          b(5, "enableForgottenToAttachAlert");
          b(6, "enableGetContactGroupMembers");
          b(7, "enableGetExchangeDistributionListMembers");
          b(8, "contactGroupMembersAreWhite");
          b(9, "exchangeDistributionListMembersAreWhite");
          b(10, "isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles");
          b(11, "isDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain");
          b(12, "isDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain");
          b(13, "isDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain");
          b(14, "isEnableRecipientsAreSortedByDomain");
          b(15, "isAutoAddSenderToBcc");
          b(16, "isAutoCheckRegisteredInContacts");
          b(17, "isAutoCheckRegisteredInContactsAndMemberOfContactLists");
          b(18, "isCheckNameAndDomainsFromRecipients");
          b(19, "isWarningIfRecipientsIsNotRegistered");
          b(20, "isProhibitsSendingMailIfRecipientsIsNotRegistered");
          b(21, "isShowConfirmationAtSendMeetingRequest");
          b(22, "isAutoAddSenderToCc");
          b(23, "isCheckNameAndDomainsIncludeSubject");
          b(24, "isCheckNameAndDomainsFromSubject");
          b(25, "isShowConfirmationAtSendTaskRequest");
          b(26, "isAutoCheckAttachments");
          b(27, "isCheckKeywordAndRecipientsIncludeSubject");

          summary.push("Imported GeneralSetting.csv");
          break;
        }
        case "internaldomainlist.csv": {
          next.internalDomains = applyList(rows, function (r) {
            var domain = normalizeString(r[0]);
            return domain ? { domain: domain } : null;
          });
          summary.push("Imported InternalDomainList.csv (" + String(next.internalDomains.length) + ")");
          break;
        }
        case "whitelist.csv": {
          next.whitelist = applyList(rows, function (r) {
            var whiteName = normalizeString(r[0]);
            if (!whiteName) return null;
            var skip = parseBool(r[1]);
            return { whiteName: whiteName, isSkipConfirmation: skip === true };
          });
          summary.push("Imported Whitelist.csv (" + String(next.whitelist.length) + ")");
          break;
        }
        case "alertaddresslist.csv": {
          next.alertAddresses = applyList(rows, function (r) {
            var target = normalizeString(r[0]);
            if (!target) return null;
            var isCanNotSend = parseBool(r[1]);
            return { targetAddress: target, isCanNotSend: isCanNotSend === true, message: normalizeString(r[2]) };
          });
          summary.push("Imported AlertAddressList.csv (" + String(next.alertAddresses.length) + ")");
          break;
        }
        case "alertkeywordandmessagelist.csv": {
          next.alertKeywordsBody = applyList(rows, function (r) {
            var kw = normalizeString(r[0]);
            if (!kw) return null;
            var isCanNotSend = parseBool(r[2]);
            return { alertKeyword: kw, message: normalizeString(r[1]), isCanNotSend: isCanNotSend === true };
          });
          summary.push("Imported AlertKeywordAndMessageList.csv (" + String(next.alertKeywordsBody.length) + ")");
          break;
        }
        case "alertkeywordandmessagelistforsubject.csv": {
          next.alertKeywordsSubject = applyList(rows, function (r) {
            var kw = normalizeString(r[0]);
            if (!kw) return null;
            var isCanNotSend = parseBool(r[2]);
            return { alertKeyword: kw, message: normalizeString(r[1]), isCanNotSend: isCanNotSend === true };
          });
          summary.push("Imported AlertKeywordAndMessageListForSubject.csv (" + String(next.alertKeywordsSubject.length) + ")");
          break;
        }
        case "autoccbcckeywordlist.csv": {
          next.autoCcBccKeyword = applyList(rows, function (r) {
            var keyword = normalizeString(r[0]);
            var ccOrBcc = normalizeString(r[1]);
            var addr = normalizeString(r[2]);
            if (!keyword || !addr) return null;
            return { keyword: keyword, ccOrBcc: ccOrBcc === "Cc" ? "Cc" : "Bcc", autoAddAddress: addr };
          });
          summary.push("Imported AutoCcBccKeywordList.csv (" + String(next.autoCcBccKeyword.length) + ")");
          break;
        }
        case "autoccbccrecipientlist.csv": {
          next.autoCcBccRecipient = applyList(rows, function (r) {
            var targetRecipient = normalizeString(r[0]);
            var ccOrBcc = normalizeString(r[1]);
            var addr = normalizeString(r[2]);
            if (!targetRecipient || !addr) return null;
            return { targetRecipient: targetRecipient, ccOrBcc: ccOrBcc === "Cc" ? "Cc" : "Bcc", autoAddAddress: addr };
          });
          summary.push("Imported AutoCcBccRecipientList.csv (" + String(next.autoCcBccRecipient.length) + ")");
          break;
        }
        case "autoccbccattachedfilelist.csv": {
          next.autoCcBccAttachedFile = applyList(rows, function (r) {
            var ccOrBcc = normalizeString(r[0]);
            var addr = normalizeString(r[1]);
            if (!addr) return null;
            return { ccOrBcc: ccOrBcc === "Cc" ? "Cc" : "Bcc", autoAddAddress: addr };
          });
          summary.push("Imported AutoCcBccAttachedFileList.csv (" + String(next.autoCcBccAttachedFile.length) + ")");
          break;
        }
        case "nameanddomains.csv": {
          next.nameAndDomains = applyList(rows, function (r) {
            var name = normalizeString(r[0]);
            var domain = normalizeString(r[1]);
            if (!name || !domain) return null;
            return { name: name, domain: domain };
          });
          summary.push("Imported NameAndDomains.csv (" + String(next.nameAndDomains.length) + ")");
          break;
        }
        case "keywordandrecipientslist.csv":
        {
          next.keywordAndRecipients = applyList(rows, function (r) {
            var keyword = normalizeString(r[0]);
            var recipient = normalizeString(r[1]);
            if (!keyword || !recipient) return null;
            return { keyword: keyword, recipient: recipient };
          });
          summary.push("Imported KeywordAndRecipientsList.csv (" + String(next.keywordAndRecipients.length) + ")");
          break;
        }
        case "externaldomainswarningandautochangetobccsetting.csv": {
          if (rows.length === 0) break;
          var r1 = rows[0];
          var num = parseIntOrNull(r1[0]);
          if (num != null) next.externalDomains.targetToAndCcExternalDomainsNum = num;
          var b1 = parseBool(r1[1]);
          var b2 = parseBool(r1[2]);
          var b3 = parseBool(r1[3]);
          if (b1 != null) next.externalDomains.isWarningWhenLargeNumberOfExternalDomains = b1;
          if (b2 != null) next.externalDomains.isProhibitedWhenLargeNumberOfExternalDomains = b2;
          if (b3 != null) next.externalDomains.isAutoChangeToBccWhenLargeNumberOfExternalDomains = b3;
          summary.push("Imported ExternalDomainsWarningAndAutoChangeToBccSetting.csv");
          break;
        }
        case "forceautochangerecipientstobcc.csv": {
          if (rows.length === 0) break;
          var r2 = rows[0];
          var b = parseBool(r2[0]);
          if (b != null) next.forceAutoChangeRecipientsToBcc.isForceAutoChangeRecipientsToBcc = b;
          next.forceAutoChangeRecipientsToBcc.toRecipient = normalizeString(r2[1]);
          var b4 = parseBool(r2[2]);
          if (b4 != null) next.forceAutoChangeRecipientsToBcc.isIncludeInternalDomain = b4;
          summary.push("Imported ForceAutoChangeRecipientsToBcc.csv");
          break;
        }
        case "attachmentssetting.csv": {
          if (rows.length === 0) break;
          var r3 = rows[0];
          function ab(i, key) {
            var v = parseBool(r3[i]);
            if (v == null) return;
            next.attachmentsSetting[key] = v;
          }
          ab(0, "isWarningWhenEncryptedZipIsAttached");
          ab(1, "isProhibitedWhenEncryptedZipIsAttached");
          ab(2, "isEnableAllAttachedFilesAreDetectEncryptedZip");
          ab(3, "isAttachmentsProhibited");
          ab(4, "isWarningWhenAttachedRealFile");
          ab(5, "isEnableOpenAttachedFiles");
          if (r3.length > 6) next.attachmentsSetting.targetAttachmentFileExtensionOfOpen = normalizeString(r3[6]);
          ab(7, "isMustOpenBeforeCheckTheAttachedFiles");
          ab(8, "isIgnoreMustOpenBeforeCheckTheAttachedFilesIfInternalDomain");
          summary.push("Imported AttachmentsSetting.csv");
          break;
        }
        case "attachmentprohibitedrecipients.csv": {
          next.attachmentProhibitedRecipients = applyList(rows, function (r) {
            var recipient = normalizeString(r[0]);
            return recipient ? { recipient: recipient } : null;
          });
          summary.push("Imported AttachmentProhibitedRecipients.csv (" + String(next.attachmentProhibitedRecipients.length) + ")");
          break;
        }
        case "attachmentalertrecipients.csv": {
          next.attachmentAlertRecipients = applyList(rows, function (r) {
            var recipient = normalizeString(r[0]);
            if (!recipient) return null;
            return { recipient: recipient, message: normalizeString(r[1]) };
          });
          summary.push("Imported AttachmentAlertRecipients.csv (" + String(next.attachmentAlertRecipients.length) + ")");
          break;
        }
        case "recipientsandattachmentsname.csv": {
          next.recipientsAndAttachmentsName = applyList(rows, function (r) {
            var attachmentsName = normalizeString(r[0]);
            var recipient = normalizeString(r[1]);
            if (!attachmentsName || !recipient) return null;
            return { attachmentsName: attachmentsName, recipient: recipient };
          });
          summary.push("Imported RecipientsAndAttachmentsName.csv (" + String(next.recipientsAndAttachmentsName.length) + ")");
          break;
        }
        case "autodeleterecipientlist.csv": {
          next.autoDeleteRecipients = applyList(rows, function (r) {
            var recipient = normalizeString(r[0]);
            return recipient ? { recipient: recipient } : null;
          });
          summary.push("Imported AutoDeleteRecipientList.csv (" + String(next.autoDeleteRecipients.length) + ")");
          break;
        }
        case "autoaddmessage.csv": {
          if (rows.length === 0) break;
          var r4 = rows[0];
          var b5 = parseBool(r4[0]);
          var b6 = parseBool(r4[1]);
          if (b5 != null) next.autoAddMessage.isAddToStart = b5;
          if (b6 != null) next.autoAddMessage.isAddToEnd = b6;
          next.autoAddMessage.messageOfAddToStart = normalizeString(r4[2]);
          next.autoAddMessage.messageOfAddToEnd = normalizeString(r4[3]);
          summary.push("Imported AutoAddMessage.csv");
          break;
        }
        case "securityforreceivedmail.csv": {
          if (rows.length === 0) break;
          var r5 = rows[0];
          function sb(i, key) {
            var v = parseBool(r5[i]);
            if (v == null) return;
            next.securityForReceivedMail[key] = v;
          }
          sb(0, "isEnableSecurityForReceivedMail");
          sb(1, "isEnableAlertKeywordOfSubjectWhenOpeningMailsData");
          sb(2, "isEnableMailHeaderAnalysis");
          sb(3, "isShowWarningWhenSpfFails");
          sb(4, "isShowWarningWhenDkimFails");
          sb(5, "isEnableWarningFeatureWhenOpeningAttachments");
          sb(6, "isWarnBeforeOpeningAttachments");
          sb(7, "isWarnBeforeOpeningEncryptedZip");
          sb(8, "isWarnLinkFileInTheZip");
          sb(9, "isWarnOneFileInTheZip");
          sb(10, "isWarnOfficeFileWithMacroInTheZip");
          sb(11, "isWarnBeforeOpeningAttachmentsThatContainMacros");
          sb(12, "isShowWarningWhenSpoofingRisk");
          sb(13, "isShowWarningWhenDmarcNotImplemented");
          summary.push("Imported SecurityForReceivedMail.csv");
          break;
        }
        case "alertkeywordofsubjectwhenopeningmaillist.csv": {
          next.alertKeywordOfSubjectWhenOpeningMail = applyList(rows, function (r) {
            var kw = normalizeString(r[0]);
            if (!kw) return null;
            return { alertKeyword: kw, message: normalizeString(r[1]) };
          });
          summary.push(
            "Imported AlertKeywordOfSubjectWhenOpeningMailList.csv (" +
              String(next.alertKeywordOfSubjectWhenOpeningMail.length) +
              ")"
          );
          break;
        }
        case "deferreddeliveryminutes.csv": {
          ignored.push(getBaseName(fe.name));
          break;
        }
        default:
          // Unknown CSV: ignore
          break;
      }
    }

    try {
      if (MailChecker && MailChecker.settings && typeof MailChecker.settings._normalize === "function") {
        next = MailChecker.settings._normalize(next);
      }
    } catch (_e2) {}

    if (ignored.length > 0) {
      summary.push("Ignored: " + ignored.join(", "));
    }

    return { settings: next, summary: summary };
  }

  function exportOutlookOkanCsvFiles(settings) {
    var s = settings || {};
    var g = (s.general || {});
    var out = [];

    function pushFile(name, rows) {
      out.push({ name: name, text: toCsv(rows) });
    }

    pushFile("GeneralSetting.csv", [
      [
        boolToYesNo(!!g.isDoNotConfirmationIfAllRecipientsAreSameDomain),
        boolToYesNo(!!g.isDoDoNotConfirmationIfAllWhite),
        boolToYesNo(!!g.isAutoCheckIfAllRecipientsAreSameDomain),
        normalizeString(g.languageCode),
        boolToYesNo(!!g.isShowConfirmationToMultipleDomain),
        boolToYesNo(g.enableForgottenToAttachAlert !== false),
        boolToYesNo(!!g.enableGetContactGroupMembers),
        boolToYesNo(!!g.enableGetExchangeDistributionListMembers),
        boolToYesNo(g.contactGroupMembersAreWhite !== false),
        boolToYesNo(g.exchangeDistributionListMembersAreWhite !== false),
        boolToYesNo(!!g.isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles),
        boolToYesNo(!!g.isDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain),
        boolToYesNo(!!g.isDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain),
        boolToYesNo(!!g.isDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain),
        boolToYesNo(!!g.isEnableRecipientsAreSortedByDomain),
        boolToYesNo(!!g.isAutoAddSenderToBcc),
        boolToYesNo(!!g.isAutoCheckRegisteredInContacts),
        boolToYesNo(!!g.isAutoCheckRegisteredInContactsAndMemberOfContactLists),
        boolToYesNo(!!g.isCheckNameAndDomainsFromRecipients),
        boolToYesNo(!!g.isWarningIfRecipientsIsNotRegistered),
        boolToYesNo(!!g.isProhibitsSendingMailIfRecipientsIsNotRegistered),
        boolToYesNo(!!g.isShowConfirmationAtSendMeetingRequest),
        boolToYesNo(!!g.isAutoAddSenderToCc),
        boolToYesNo(!!g.isCheckNameAndDomainsIncludeSubject),
        boolToYesNo(!!g.isCheckNameAndDomainsFromSubject),
        boolToYesNo(!!g.isShowConfirmationAtSendTaskRequest),
        boolToYesNo(!!g.isAutoCheckAttachments),
        boolToYesNo(!!g.isCheckKeywordAndRecipientsIncludeSubject),
      ],
    ]);

    if (Array.isArray(s.internalDomains) && s.internalDomains.length > 0) {
      pushFile(
        "InternalDomainList.csv",
        s.internalDomains
          .map(function (d) {
            return d && d.domain ? [d.domain] : null;
          })
          .filter(Boolean)
      );
    }

    if (Array.isArray(s.whitelist) && s.whitelist.length > 0) {
      pushFile(
        "Whitelist.csv",
        s.whitelist
          .map(function (w) {
            if (!w || !w.whiteName) return null;
            return [w.whiteName, boolToYesNo(!!w.isSkipConfirmation)];
          })
          .filter(Boolean)
      );
    }

    if (Array.isArray(s.alertAddresses) && s.alertAddresses.length > 0) {
      pushFile(
        "AlertAddressList.csv",
        s.alertAddresses
          .map(function (a) {
            if (!a || !a.targetAddress) return null;
            return [a.targetAddress, boolToYesNo(!!a.isCanNotSend), normalizeString(a.message)];
          })
          .filter(Boolean)
      );
    }

    if (Array.isArray(s.alertKeywordsBody) && s.alertKeywordsBody.length > 0) {
      pushFile(
        "AlertKeywordAndMessageList.csv",
        s.alertKeywordsBody
          .map(function (a) {
            if (!a || !a.alertKeyword) return null;
            return [a.alertKeyword, normalizeString(a.message), boolToYesNo(!!a.isCanNotSend)];
          })
          .filter(Boolean)
      );
    }

    if (Array.isArray(s.alertKeywordsSubject) && s.alertKeywordsSubject.length > 0) {
      pushFile(
        "AlertKeywordAndMessageListForSubject.csv",
        s.alertKeywordsSubject
          .map(function (a) {
            if (!a || !a.alertKeyword) return null;
            return [a.alertKeyword, normalizeString(a.message), boolToYesNo(!!a.isCanNotSend)];
          })
          .filter(Boolean)
      );
    }

    if (Array.isArray(s.autoCcBccKeyword) && s.autoCcBccKeyword.length > 0) {
      pushFile(
        "AutoCcBccKeywordList.csv",
        s.autoCcBccKeyword
          .map(function (r) {
            if (!r || !r.keyword || !r.autoAddAddress) return null;
            return [r.keyword, r.ccOrBcc === "Cc" ? "Cc" : "Bcc", r.autoAddAddress];
          })
          .filter(Boolean)
      );
    }

    if (Array.isArray(s.autoCcBccRecipient) && s.autoCcBccRecipient.length > 0) {
      pushFile(
        "AutoCcBccRecipientList.csv",
        s.autoCcBccRecipient
          .map(function (r) {
            if (!r || !r.targetRecipient || !r.autoAddAddress) return null;
            return [r.targetRecipient, r.ccOrBcc === "Cc" ? "Cc" : "Bcc", r.autoAddAddress];
          })
          .filter(Boolean)
      );
    }

    if (Array.isArray(s.autoCcBccAttachedFile) && s.autoCcBccAttachedFile.length > 0) {
      pushFile(
        "AutoCcBccAttachedFileList.csv",
        s.autoCcBccAttachedFile
          .map(function (r) {
            if (!r || !r.autoAddAddress) return null;
            return [r.ccOrBcc === "Cc" ? "Cc" : "Bcc", r.autoAddAddress];
          })
          .filter(Boolean)
      );
    }

    if (Array.isArray(s.nameAndDomains) && s.nameAndDomains.length > 0) {
      pushFile(
        "NameAndDomains.csv",
        s.nameAndDomains
          .map(function (r) {
            if (!r || !r.name || !r.domain) return null;
            return [r.name, r.domain];
          })
          .filter(Boolean)
      );
    }

    if (Array.isArray(s.keywordAndRecipients) && s.keywordAndRecipients.length > 0) {
      pushFile(
        "KeywordAndRecipientsList.csv",
        s.keywordAndRecipients
          .map(function (r) {
            if (!r || !r.keyword || !r.recipient) return null;
            return [r.keyword, r.recipient];
          })
          .filter(Boolean)
      );
    }

    pushFile("ExternalDomainsWarningAndAutoChangeToBccSetting.csv", [
      [
        String((s.externalDomains && s.externalDomains.targetToAndCcExternalDomainsNum) || 10),
        boolToYesNo(!!(s.externalDomains && s.externalDomains.isWarningWhenLargeNumberOfExternalDomains)),
        boolToYesNo(!!(s.externalDomains && s.externalDomains.isProhibitedWhenLargeNumberOfExternalDomains)),
        boolToYesNo(!!(s.externalDomains && s.externalDomains.isAutoChangeToBccWhenLargeNumberOfExternalDomains)),
      ],
    ]);

    pushFile("ForceAutoChangeRecipientsToBcc.csv", [
      [
        boolToYesNo(!!(s.forceAutoChangeRecipientsToBcc && s.forceAutoChangeRecipientsToBcc.isForceAutoChangeRecipientsToBcc)),
        normalizeString(s.forceAutoChangeRecipientsToBcc && s.forceAutoChangeRecipientsToBcc.toRecipient),
        boolToYesNo(!!(s.forceAutoChangeRecipientsToBcc && s.forceAutoChangeRecipientsToBcc.isIncludeInternalDomain)),
      ],
    ]);

    var att = s.attachmentsSetting || {};
    pushFile("AttachmentsSetting.csv", [
      [
        boolToYesNo(!!att.isWarningWhenEncryptedZipIsAttached),
        boolToYesNo(!!att.isProhibitedWhenEncryptedZipIsAttached),
        boolToYesNo(!!att.isEnableAllAttachedFilesAreDetectEncryptedZip),
        boolToYesNo(!!att.isAttachmentsProhibited),
        boolToYesNo(!!att.isWarningWhenAttachedRealFile),
        boolToYesNo(!!att.isEnableOpenAttachedFiles),
        normalizeString(att.targetAttachmentFileExtensionOfOpen),
        boolToYesNo(!!att.isMustOpenBeforeCheckTheAttachedFiles),
        boolToYesNo(!!att.isIgnoreMustOpenBeforeCheckTheAttachedFilesIfInternalDomain),
      ],
    ]);

    if (Array.isArray(s.attachmentProhibitedRecipients) && s.attachmentProhibitedRecipients.length > 0) {
      pushFile(
        "AttachmentProhibitedRecipients.csv",
        s.attachmentProhibitedRecipients
          .map(function (r) {
            if (!r || !r.recipient) return null;
            return [r.recipient];
          })
          .filter(Boolean)
      );
    }

    if (Array.isArray(s.attachmentAlertRecipients) && s.attachmentAlertRecipients.length > 0) {
      pushFile(
        "AttachmentAlertRecipients.csv",
        s.attachmentAlertRecipients
          .map(function (r) {
            if (!r || !r.recipient) return null;
            return [r.recipient, normalizeString(r.message)];
          })
          .filter(Boolean)
      );
    }

    if (Array.isArray(s.recipientsAndAttachmentsName) && s.recipientsAndAttachmentsName.length > 0) {
      pushFile(
        "RecipientsAndAttachmentsName.csv",
        s.recipientsAndAttachmentsName
          .map(function (r) {
            if (!r || !r.attachmentsName || !r.recipient) return null;
            return [r.attachmentsName, r.recipient];
          })
          .filter(Boolean)
      );
    }

    if (Array.isArray(s.autoDeleteRecipients) && s.autoDeleteRecipients.length > 0) {
      pushFile(
        "AutoDeleteRecipientList.csv",
        s.autoDeleteRecipients
          .map(function (r) {
            if (!r || !r.recipient) return null;
            return [r.recipient];
          })
          .filter(Boolean)
      );
    }

    var am = s.autoAddMessage || {};
    pushFile("AutoAddMessage.csv", [
      [
        boolToYesNo(!!am.isAddToStart),
        boolToYesNo(!!am.isAddToEnd),
        normalizeString(am.messageOfAddToStart),
        normalizeString(am.messageOfAddToEnd),
      ],
    ]);

    var sec = s.securityForReceivedMail || {};
    pushFile("SecurityForReceivedMail.csv", [
      [
        boolToYesNo(!!sec.isEnableSecurityForReceivedMail),
        boolToYesNo(!!sec.isEnableAlertKeywordOfSubjectWhenOpeningMailsData),
        boolToYesNo(!!sec.isEnableMailHeaderAnalysis),
        boolToYesNo(!!sec.isShowWarningWhenSpfFails),
        boolToYesNo(!!sec.isShowWarningWhenDkimFails),
        boolToYesNo(!!sec.isEnableWarningFeatureWhenOpeningAttachments),
        boolToYesNo(!!sec.isWarnBeforeOpeningAttachments),
        boolToYesNo(!!sec.isWarnBeforeOpeningEncryptedZip),
        boolToYesNo(!!sec.isWarnLinkFileInTheZip),
        boolToYesNo(!!sec.isWarnOneFileInTheZip),
        boolToYesNo(!!sec.isWarnOfficeFileWithMacroInTheZip),
        boolToYesNo(!!sec.isWarnBeforeOpeningAttachmentsThatContainMacros),
        boolToYesNo(!!sec.isShowWarningWhenSpoofingRisk),
        boolToYesNo(!!sec.isShowWarningWhenDmarcNotImplemented),
      ],
    ]);

    if (Array.isArray(s.alertKeywordOfSubjectWhenOpeningMail) && s.alertKeywordOfSubjectWhenOpeningMail.length > 0) {
      pushFile(
        "AlertKeywordOfSubjectWhenOpeningMailList.csv",
        s.alertKeywordOfSubjectWhenOpeningMail
          .map(function (r) {
            if (!r || !r.alertKeyword) return null;
            return [r.alertKeyword, normalizeString(r.message)];
          })
          .filter(Boolean)
      );
    }

    return out;
  }

  function pGetRecipients(recipientsObj) {
    return new Promise(function (resolve) {
      try {
        if (!recipientsObj || typeof recipientsObj.getAsync !== "function") return resolve([]);
        recipientsObj.getAsync(function (result) {
          try {
            if (!result || result.status !== Office.AsyncResultStatus.Succeeded) return resolve([]);
            resolve(Array.isArray(result.value) ? result.value : []);
          } catch (_e) {
            resolve([]);
          }
        });
      } catch (_e2) {
        resolve([]);
      }
    });
  }

  function pSetRecipients(recipientsObj, emails) {
    return new Promise(function (resolve) {
      try {
        if (!recipientsObj || typeof recipientsObj.setAsync !== "function") return resolve(false);
        var list = Array.isArray(emails)
          ? emails
              .map(function (e) {
                var v = normalizeString(e);
                return v ? { emailAddress: v } : null;
              })
              .filter(Boolean)
          : [];
        recipientsObj.setAsync(list, function (result) {
          try {
            resolve(!!result && result.status === Office.AsyncResultStatus.Succeeded);
          } catch (_e) {
            resolve(false);
          }
        });
      } catch (_e2) {
        resolve(false);
      }
    });
  }

  function pGetSubject(item) {
    return new Promise(function (resolve) {
      try {
        if (!item || !item.subject || typeof item.subject.getAsync !== "function") return resolve("");
        item.subject.getAsync(function (result) {
          try {
            resolve(result && result.status === Office.AsyncResultStatus.Succeeded ? normalizeString(result.value) : "");
          } catch (_e) {
            resolve("");
          }
        });
      } catch (_e2) {
        resolve("");
      }
    });
  }

  function pGetBodyText(item) {
    return new Promise(function (resolve) {
      try {
        if (!item || !item.body || typeof item.body.getAsync !== "function") return resolve("");
        item.body.getAsync(Office.CoercionType.Text, function (result) {
          try {
            resolve(result && result.status === Office.AsyncResultStatus.Succeeded ? normalizeString(result.value) : "");
          } catch (_e) {
            resolve("");
          }
        });
      } catch (_e2) {
        resolve("");
      }
    });
  }

  function pGetAllInternetHeaders(item) {
    return new Promise(function (resolve) {
      try {
        if (!item || typeof item.getAllInternetHeadersAsync !== "function") return resolve("");
        item.getAllInternetHeadersAsync(function (result) {
          try {
            if (!result || result.status !== Office.AsyncResultStatus.Succeeded) return resolve("");
            resolve(String(result.value || ""));
          } catch (_e) {
            resolve("");
          }
        });
      } catch (_e2) {
        resolve("");
      }
    });
  }

  function pGetAttachmentContent(item, attachmentId) {
    return new Promise(function (resolve) {
      try {
        if (!item || typeof item.getAttachmentContentAsync !== "function") return resolve(null);
        item.getAttachmentContentAsync(String(attachmentId || ""), function (result) {
          try {
            if (!result || result.status !== Office.AsyncResultStatus.Succeeded) return resolve(null);
            resolve(result.value || null);
          } catch (_e) {
            resolve(null);
          }
        });
      } catch (_e2) {
        resolve(null);
      }
    });
  }

  function base64ToBytes(base64) {
    try {
      var bin = atob(String(base64 || ""));
      var len = bin.length;
      var bytes = new Uint8Array(len);
      for (var i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i) & 0xff;
      return bytes;
    } catch (_e) {
      return null;
    }
  }

  function fileExt(name) {
    var n = normalizeString(name);
    var dot = n.lastIndexOf(".");
    if (dot < 0) return "";
    return lower(n.slice(dot));
  }

  function applyAutoAddMessagePreviewToText(bodyText, autoAddMessage) {
    var out = String(bodyText || "");
    if (!autoAddMessage) return out;
    if (autoAddMessage.isAddToStart && normalizeString(autoAddMessage.messageOfAddToStart)) {
      out = normalizeString(autoAddMessage.messageOfAddToStart) + "\n\n" + out;
    }
    if (autoAddMessage.isAddToEnd && normalizeString(autoAddMessage.messageOfAddToEnd)) {
      out = out + "\n\n" + normalizeString(autoAddMessage.messageOfAddToEnd);
    }
    return out;
  }

  function renderSecurityResult(scan) {
    var el = $("security-summary");
    if (!el) return;

    if (!scan) {
      el.innerHTML = "<h2>Security</h2><p class=\"hint\">(no data)</p>";
      return;
    }

    var locale = scan.locale || "en-US";
    var parts = [];
    parts.push("<h2>" + esc(t(locale, "受信メールセキュリティ", "Received Mail Security")) + "</h2>");

    if (!scan.enabled) {
      parts.push("<p class=\"hint\">" + esc(t(locale, "無効です (設定で有効化できます)", "Disabled (enable in settings).")) + "</p>");
      el.innerHTML = parts.join("");
      return;
    }

    if (scan.subjectAlerts && scan.subjectAlerts.length) {
      parts.push("<p><strong>" + esc(t(locale, "件名警告", "Subject alerts")) + ":</strong></p>");
      parts.push(
        "<ul>" +
          scan.subjectAlerts
            .map(function (m) {
              return "<li>" + esc(m) + "</li>";
            })
            .join("") +
          "</ul>"
      );
    }

    if (scan.headerWarnings && scan.headerWarnings.length) {
      parts.push("<p><strong>" + esc(t(locale, "ヘッダ警告", "Header warnings")) + ":</strong></p>");
      parts.push(
        "<ul>" +
          scan.headerWarnings
            .map(function (m) {
              return "<li>" + esc(m) + "</li>";
            })
            .join("") +
          "</ul>"
      );
    }

    if (scan.headerAnalysis) {
      var rows = Object.keys(scan.headerAnalysis).map(function (k) {
        return "<tr><td><code>" + esc(k) + "</code></td><td>" + esc(String(scan.headerAnalysis[k])) + "</td></tr>";
      });
      parts.push("<details><summary>" + esc(t(locale, "ヘッダ解析結果", "Header analysis details")) + "</summary>");
      parts.push("<div style=\"overflow:auto\"><table><tbody>" + rows.join("") + "</tbody></table></div>");
      parts.push("</details>");
    }

    if (scan.attachmentWarnings && scan.attachmentWarnings.length) {
      parts.push("<p><strong>" + esc(t(locale, "添付ファイル警告", "Attachment warnings")) + ":</strong></p>");
      parts.push(
        "<ul>" +
          scan.attachmentWarnings
            .map(function (m) {
              return "<li>" + esc(m) + "</li>";
            })
            .join("") +
          "</ul>"
      );
    }

    if (
      (!scan.subjectAlerts || scan.subjectAlerts.length === 0) &&
      (!scan.headerWarnings || scan.headerWarnings.length === 0) &&
      (!scan.attachmentWarnings || scan.attachmentWarnings.length === 0)
    ) {
      parts.push("<p class=\"hint\">" + esc(t(locale, "警告はありません。", "No warnings.")) + "</p>");
    }

    el.innerHTML = parts.join("");
  }

  async function runSecurityScan(item, settings) {
    var locale = pickLocale(settings);
    var sec = (settings && settings.securityForReceivedMail) || {};

    var scan = {
      locale: locale,
      enabled: !!sec.isEnableSecurityForReceivedMail,
      subject: "",
      subjectAlerts: [],
      headerAnalysis: null,
      headerWarnings: [],
      attachmentWarnings: [],
    };

    if (!scan.enabled) return scan;

    scan.subject = await pGetSubject(item);

    // Subject keyword alerts
    if (sec.isEnableAlertKeywordOfSubjectWhenOpeningMailsData && Array.isArray(settings.alertKeywordOfSubjectWhenOpeningMail)) {
      for (var i = 0; i < settings.alertKeywordOfSubjectWhenOpeningMail.length; i++) {
        var row = settings.alertKeywordOfSubjectWhenOpeningMail[i];
        if (!row || !row.alertKeyword) continue;
        if (scan.subject.indexOf(row.alertKeyword) < 0) continue;
        scan.subjectAlerts.push(normalizeString(row.message) || (t(locale, "件名に警告キーワードがあります", "Warning keyword in subject") + ": [" + row.alertKeyword + "]"));
      }
    }

    // Header analysis
    if (sec.isEnableMailHeaderAnalysis && MailChecker.readSecurity) {
      var headers = await pGetAllInternetHeaders(item);
      if (headers) {
        var analysis = MailChecker.readSecurity.validateEmailHeader(headers);
        scan.headerAnalysis = analysis;
        if (analysis) {
          var isInternalMail =
            analysis.SPF === "NONE" &&
            analysis.DKIM === "NONE" &&
            analysis.DMARC === "NONE" &&
            analysis.Internal === "TRUE";

          if (!isInternalMail) {
            if (sec.isShowWarningWhenSpoofingRisk) {
              if (sec.isShowWarningWhenDmarcNotImplemented) {
                if (analysis.DMARC !== "PASS") {
                  scan.headerWarnings.push(t(locale, "なりすましの可能性があります(DMARC未達)", "Possible spoofing risk (DMARC not PASS)."));
                }
              } else {
                var selfDmarc = MailChecker.readSecurity.determineDmarcResult(
                  analysis.SPF,
                  analysis["SPF Alignment"],
                  analysis.DKIM,
                  analysis["DKIM Alignment"]
                );
                if (analysis.DMARC !== "PASS" && analysis.DMARC !== "BESTGUESSPASS" && selfDmarc === "FAIL") {
                  scan.headerWarnings.push(t(locale, "なりすましの可能性があります(SPF/DKIM)", "Possible spoofing risk (SPF/DKIM)."));
                }
              }
            } else {
              if (sec.isShowWarningWhenSpfFails) {
                if (analysis.SPF === "FAIL" || analysis.SPF === "NONE") {
                  scan.headerWarnings.push(t(locale, "SPF検証に失敗しました", "SPF validation failed."));
                }
              }
              if (sec.isShowWarningWhenDkimFails) {
                if (analysis.DKIM === "FAIL") {
                  scan.headerWarnings.push(t(locale, "DKIM検証に失敗しました", "DKIM validation failed."));
                }
              }
            }
          }
        }
      } else {
        scan.headerWarnings.push(t(locale, "ヘッダ取得ができませんでした。", "Could not read internet headers."));
      }
    }

    // Attachment checks (best-effort)
    if (sec.isEnableWarningFeatureWhenOpeningAttachments) {
      var attachments = [];
      try {
        attachments = Array.isArray(item.attachments) ? item.attachments : [];
      } catch (_e) {
        attachments = [];
      }

      if (sec.isWarnBeforeOpeningAttachments && attachments.length > 0) {
        scan.attachmentWarnings.push(
          t(locale, "添付ファイルを開く前に内容を確認してください。", "Review attachments before opening.")
        );
      }

      var needsZip =
        sec.isWarnBeforeOpeningEncryptedZip ||
        sec.isWarnLinkFileInTheZip ||
        sec.isWarnOneFileInTheZip ||
        sec.isWarnOfficeFileWithMacroInTheZip;

      for (var a = 0; a < attachments.length; a++) {
        var att = attachments[a];
        if (!att) continue;
        var name = att.name || att.fileName || "";
        var ext = fileExt(name);

        if (needsZip && ext === ".zip") {
          var content = await pGetAttachmentContent(item, att.id);
          if (content && content.content && MailChecker.readSecurity && MailChecker.readSecurity.parseZipCentralDirectory) {
            // Content format differs across clients; only Base64 is handled here.
            var bytes = base64ToBytes(content.content);
            if (bytes && bytes.length) {
              var zip = MailChecker.readSecurity.parseZipCentralDirectory(bytes);
              if (sec.isWarnBeforeOpeningEncryptedZip && zip.isEncrypted) {
                scan.attachmentWarnings.push(t(locale, "暗号化ZIPの可能性: ", "Possible encrypted ZIP: ") + name);
              }
              if (sec.isWarnLinkFileInTheZip && zip.includeExtensions && zip.includeExtensions.indexOf(".lnk") >= 0) {
                scan.attachmentWarnings.push(t(locale, "ZIP内に.lnkがあります: ", "ZIP contains .lnk: ") + name);
              }
              if (sec.isWarnOneFileInTheZip && zip.includeExtensions && zip.includeExtensions.indexOf(".one") >= 0) {
                scan.attachmentWarnings.push(t(locale, "ZIP内に.oneがあります: ", "ZIP contains .one: ") + name);
              }
              if (
                sec.isWarnOfficeFileWithMacroInTheZip &&
                zip.includeExtensions &&
                (zip.includeExtensions.indexOf(".docm") >= 0 ||
                  zip.includeExtensions.indexOf(".xlsm") >= 0 ||
                  zip.includeExtensions.indexOf(".pptm") >= 0)
              ) {
                scan.attachmentWarnings.push(t(locale, "ZIP内にマクロ付きOfficeがあります: ", "ZIP contains macro Office file: ") + name);
              }
            } else {
              scan.attachmentWarnings.push(t(locale, "ZIP解析に失敗: ", "ZIP analysis failed: ") + name);
            }
          } else {
            scan.attachmentWarnings.push(t(locale, "ZIP解析に未対応: ", "ZIP analysis not available: ") + name);
          }
        }

        if (sec.isWarnBeforeOpeningAttachmentsThatContainMacros) {
          if (ext === ".docm" || ext === ".xlsm" || ext === ".pptm") {
            scan.attachmentWarnings.push(t(locale, "マクロ付きOfficeの可能性: ", "Macro-enabled Office file: ") + name);
          }
        }
      }
    }

    return scan;
  }

  async function buildSnapshotFromItem(item, settings) {
    var itemType = "";
    try {
      itemType = normalizeString(item && item.itemType);
    } catch (_e) {}

    var senderEmail = "";
    try {
      senderEmail = normalizeString(
        Office.context &&
          Office.context.mailbox &&
          Office.context.mailbox.userProfile &&
          Office.context.mailbox.userProfile.emailAddress
      );
    } catch (_e2) {}

    var to = [];
    var cc = [];
    var bcc = [];

    if (lower(itemType) === "appointment") {
      to = await pGetRecipients(item.requiredAttendees);
      cc = await pGetRecipients(item.optionalAttendees);
      bcc = [];
    } else {
      to = await pGetRecipients(item.to);
      cc = await pGetRecipients(item.cc);
      bcc = await pGetRecipients(item.bcc);
    }

    var subject = await pGetSubject(item);
    var bodyRaw = await pGetBodyText(item);
    var bodyText = applyAutoAddMessagePreviewToText(bodyRaw, settings && settings.autoAddMessage);

    var attachments = [];
    try {
      attachments = Array.isArray(item.attachments) ? item.attachments : [];
    } catch (_e3) {
      attachments = [];
    }

    return {
      displayLanguage: pickLocale(settings),
      senderEmailAddress: senderEmail,
      itemType: itemType,
      subject: subject,
      bodyText: bodyText,
      recipients: { to: to, cc: cc, bcc: bcc },
      attachments: attachments,
    };
  }

  function renderCheckResult(result) {
    var locale = (result && result.locale) || "en-US";
    var cl = result && result.checkList ? result.checkList : null;
    if (!cl) return;

    var summary = $("check-summary");
    var recipients = $("check-recipients");
    var attachments = $("check-attachments");
    var alerts = $("check-alerts");

    var decision = cl.isCanNotSendMail ? "Blocked" : result.showConfirmation ? "Confirm" : "OK";
    var pillClass = cl.isCanNotSendMail ? "danger" : result.showConfirmation ? "warn" : "ok";

    summary.innerHTML =
      "<div class=\"row\">" +
      "<span class=\"pill " +
      pillClass +
      "\">" +
      esc(decision) +
      "</span>" +
      "<span class=\"pill\">" +
      esc(t(locale, "外部ドメイン数", "External domains")) +
      ": " +
      esc(String(cl.recipientExternalDomainNumAll || 0)) +
      "</span>" +
      "</div>" +
      (cl.isCanNotSendMail
        ? "<p><strong>" + esc(t(locale, "理由", "Reason")) + ":</strong> " + esc(cl.canNotSendMailMessage) + "</p>"
        : "");

    function listRecipients(title, list) {
      var items = (list || []).map(function (r) {
        var flags = [];
        if (r.isExternal) flags.push("<span class=\"pill warn\">External</span>");
        if (r.isWhite) flags.push("<span class=\"pill ok\">White</span>");
        if (r.isSkip) flags.push("<span class=\"pill\">Skip</span>");
        if (r.isExpanded) flags.push("<span class=\"pill\">Expanded</span>");
        if (r.isRegisteredInContacts === true) flags.push("<span class=\"pill ok\">In Contacts</span>");
        if (r.isRegisteredInContacts === false) flags.push("<span class=\"pill warn\">Not in Contacts</span>");
        return "<li>" + esc(r.mailAddress) + " " + flags.join(" ") + "</li>";
      });
      return "<h2>" + esc(title) + "</h2>" + (items.length ? "<ul>" + items.join("") + "</ul>" : "<p class=\"hint\">(none)</p>");
    }

    recipients.innerHTML =
      listRecipients("To", cl.toAddresses) + listRecipients("Cc", cl.ccAddresses) + listRecipients("Bcc", cl.bccAddresses);

    attachments.innerHTML =
      "<h2>" +
      esc(t(locale, "添付ファイル", "Attachments")) +
      "</h2>" +
      (cl.attachments && cl.attachments.length
        ? "<ul>" +
          cl.attachments
            .map(function (a) {
              var flags = [];
              if (a.isTooBig) flags.push("<span class=\"pill warn\">Large</span>");
              if (a.isDangerous) flags.push("<span class=\"pill danger\">Danger</span>");
              if (a.isEncrypted) flags.push("<span class=\"pill warn\">ZIP</span>");
              return (
                "<li>" +
                esc(a.fileName) +
                " <span class=\"pill\">" +
                esc(a.fileSize) +
                "</span> " +
                flags.join(" ") +
                "</li>"
              );
            })
            .join("") +
          "</ul>"
        : "<p class=\"hint\">(none)</p>");

    alerts.innerHTML =
      "<h2>" +
      esc(t(locale, "警告", "Alerts")) +
      "</h2>" +
      (cl.alerts && cl.alerts.length
        ? "<ul>" +
          cl.alerts
            .map(function (a) {
              var cls = a.isImportant && !a.isChecked ? "pill warn" : "pill";
              return "<li><span class=\"" + cls + "\">" + esc(a.isImportant ? "!" : "i") + "</span> " + esc(a.alertMessage) + "</li>";
            })
            .join("") +
          "</ul>"
        : "<p class=\"hint\">(none)</p>");
  }

  async function applyMutationsToItem(item, snapshot, mutations) {
    if (!mutations) return;
    if (lower(snapshot.itemType) === "appointment") {
      await pSetRecipients(item.requiredAttendees, mutations.to || []);
      await pSetRecipients(item.optionalAttendees, mutations.cc || []);
      return;
    }
    await pSetRecipients(item.to, mutations.to || []);
    await pSetRecipients(item.cc, mutations.cc || []);
    await pSetRecipients(item.bcc, mutations.bcc || []);
  }

  function setupHandlers(state) {
    $("btn-run").addEventListener("click", async function () {
      try {
        setStatus("Running checks...");
        $("security-summary").innerHTML = "";
        var item = Office.context && Office.context.mailbox && Office.context.mailbox.item;
        if (!item) throw new Error("No item context");

        state.settings = await MailChecker.settings.load();
        var snapshot = await buildSnapshotFromItem(item, state.settings);
        try {
          setStatus("Resolving contacts / distribution lists...");
          await enrichSnapshotWithEws(snapshot, state.settings);
        } catch (_e4) {}
        state.lastResult = MailChecker.engine.evaluate(snapshot, state.settings);
        renderCheckResult(state.lastResult);
        $("btn-apply").disabled = !state.lastResult || !state.lastResult.mutations;
        setStatus("Done.");
      } catch (e) {
        setStatus("Failed: " + (e && e.message ? e.message : String(e)));
      }
    });

    $("btn-apply").addEventListener("click", async function () {
      try {
        var item = Office.context && Office.context.mailbox && Office.context.mailbox.item;
        if (!item) throw new Error("No item context");
        if (!state.lastResult) throw new Error("Run check first");

        var snapshot = await buildSnapshotFromItem(item, state.settings);
        await applyMutationsToItem(item, snapshot, state.lastResult.mutations);
        setStatus("Applied recipient changes.");
      } catch (e) {
        setStatus("Failed: " + (e && e.message ? e.message : String(e)));
      }
    });

    $("btn-security").addEventListener("click", async function () {
      try {
        setStatus("Running security scan...");
        var item = Office.context && Office.context.mailbox && Office.context.mailbox.item;
        if (!item) throw new Error("No item context");

        state.settings = await MailChecker.settings.load();
        var scan = await runSecurityScan(item, state.settings);
        renderSecurityResult(scan);
        setStatus("Done.");
      } catch (e) {
        setStatus("Failed: " + (e && e.message ? e.message : String(e)));
      }
    });

    $("btn-save").addEventListener("click", async function () {
      try {
        var current = state.settings || (await MailChecker.settings.load());
        var next = applySettingsFromForm(current);
        await MailChecker.settings.save(next);
        state.settings = await MailChecker.settings.load();
        renderSettingsForm(state.settings);
        setTextValue("json-editor", JSON.stringify(state.settings, null, 2));
        setStatus("Saved.");
      } catch (e) {
        setStatus("Save failed: " + (e && e.message ? e.message : String(e)));
      }
    });

    $("btn-reset").addEventListener("click", async function () {
      try {
        await MailChecker.settings.reset();
        state.settings = await MailChecker.settings.load();
        renderSettingsForm(state.settings);
        setTextValue("json-editor", JSON.stringify(state.settings, null, 2));
        setStatus("Reset.");
      } catch (e) {
        setStatus("Reset failed: " + (e && e.message ? e.message : String(e)));
      }
    });

    $("btn-load-editor").addEventListener("click", function () {
      var btn = document.querySelector('.tab[data-tab="advanced"]');
      if (btn) btn.click();
    });

    $("btn-export-json").addEventListener("click", async function () {
      try {
        var s = state.settings || (await MailChecker.settings.load());
        downloadText("mailchecker.settings.json", JSON.stringify(s, null, 2), "application/json;charset=utf-8");
        setStatus("Downloaded JSON.");
      } catch (e) {
        setStatus("Export failed: " + (e && e.message ? e.message : String(e)));
      }
    });

    $("file-import-json").addEventListener("change", function (ev) {
      var file = ev && ev.target && ev.target.files && ev.target.files[0];
      if (!file) return;
      var reader = new FileReader();
      reader.onload = async function () {
        try {
          var text = String(reader.result || "");
          var obj = JSON.parse(text);
          await MailChecker.settings.save(obj);
          state.settings = await MailChecker.settings.load();
          renderSettingsForm(state.settings);
          setTextValue("json-editor", JSON.stringify(state.settings, null, 2));
          setStatus("Imported JSON.");
        } catch (e) {
          setStatus("Import failed: " + (e && e.message ? e.message : String(e)));
        }
      };
      reader.readAsText(file);
      ev.target.value = "";
    });

    $("btn-apply-json").addEventListener("click", async function () {
      try {
        var text = getTextValue("json-editor");
        var obj = JSON.parse(text);
        await MailChecker.settings.save(obj);
        state.settings = await MailChecker.settings.load();
        renderSettingsForm(state.settings);
        setStatus("Applied JSON.");
      } catch (e) {
        setStatus("Apply failed: " + (e && e.message ? e.message : String(e)));
      }
    });

    function pReadFileAsText(file) {
      return new Promise(function (resolve, reject) {
        try {
          var reader = new FileReader();
          reader.onload = function () {
            resolve(String(reader.result || ""));
          };
          reader.onerror = function () {
            reject(new Error("Failed to read file"));
          };
          reader.readAsText(file);
        } catch (e) {
          reject(e);
        }
      });
    }

    $("btn-export-csv").addEventListener("click", async function () {
      try {
        var s2 = state.settings || (await MailChecker.settings.load());
        var files = exportOutlookOkanCsvFiles(s2);
        for (var i = 0; i < files.length; i++) {
          downloadText(files[i].name, files[i].text, "text/csv;charset=utf-8");
        }
        setStatus("Downloaded CSV files (" + String(files.length) + ").");
      } catch (e) {
        setStatus("Export failed: " + (e && e.message ? e.message : String(e)));
      }
    });

    $("file-import-csv").addEventListener("change", function (ev) {
      (async function () {
        try {
          var files = ev && ev.target && ev.target.files ? ev.target.files : null;
          if (!files || files.length === 0) return;

          setStatus("Reading CSV files...");

          var entries = [];
          for (var i = 0; i < files.length; i++) {
            var f = files[i];
            var text = await pReadFileAsText(f);
            entries.push({ name: f.name, text: text });
          }

          var baseSettings = state.settings || (await MailChecker.settings.load());
          state.csvImport = applyOutlookOkanCsvFiles(baseSettings, entries);

          var summary = (state.csvImport && state.csvImport.summary) || [];
          $("csv-summary").textContent = summary.length ? summary.join("\n") : "(no recognized CSV files)";

          setStatus("CSV loaded. Click Apply CSV to save.");
        } catch (e) {
          setStatus("CSV import failed: " + (e && e.message ? e.message : String(e)));
        }
      })();
      try {
        ev.target.value = "";
      } catch (_e5) {}
    });

    $("btn-apply-csv").addEventListener("click", async function () {
      try {
        if (!state.csvImport || !state.csvImport.settings) {
          setStatus("No CSV loaded.");
          return;
        }
        await MailChecker.settings.save(state.csvImport.settings);
        state.settings = await MailChecker.settings.load();
        renderSettingsForm(state.settings);
        setTextValue("json-editor", JSON.stringify(state.settings, null, 2));
        $("csv-summary").textContent = "";
        state.csvImport = null;
        setStatus("Applied CSV.");
      } catch (e) {
        setStatus("Apply failed: " + (e && e.message ? e.message : String(e)));
      }
    });
  }

  async function init() {
    renderTabs();

    var state = { settings: null, lastResult: null, csvImport: null };

    try {
      state.settings = await MailChecker.settings.load();
      renderSettingsForm(state.settings);
      setTextValue("json-editor", JSON.stringify(state.settings, null, 2));

      try {
        var item = Office.context && Office.context.mailbox && Office.context.mailbox.item;
        var canSecurity =
          !!item &&
          (typeof item.getAllInternetHeadersAsync === "function" || typeof item.getAttachmentContentAsync === "function");
        $("btn-security").disabled = !canSecurity;
      } catch (_e) {
        $("btn-security").disabled = true;
      }

      setStatus("Ready.");
    } catch (e) {
      setStatus("Failed to load settings: " + (e && e.message ? e.message : String(e)));
    }

    setupHandlers(state);
  }

  try {
    if (typeof Office !== "undefined" && Office.onReady) {
      Office.onReady(init);
    } else {
      document.addEventListener("DOMContentLoaded", init);
    }
  } catch (_e) {
    document.addEventListener("DOMContentLoaded", init);
  }
})();
