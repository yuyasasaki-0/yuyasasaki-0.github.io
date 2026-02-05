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

  function downloadText(filename, text) {
    try {
      var blob = new Blob([String(text || "")], { type: "application/json;charset=utf-8" });
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
        downloadText("mailchecker.settings.json", JSON.stringify(s, null, 2));
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
  }

  async function init() {
    renderTabs();

    var state = { settings: null, lastResult: null };

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
