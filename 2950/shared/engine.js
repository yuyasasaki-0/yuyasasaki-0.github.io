"use strict";
(function () {
  var root =
    typeof globalThis !== "undefined" ? globalThis : typeof window !== "undefined" ? window : undefined;
  if (!root) return;

  var ns = (root.MailChecker = root.MailChecker || {});

  function isObject(value) {
    return value != null && typeof value === "object" && !Array.isArray(value);
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

  function uniq(list) {
    if (!Array.isArray(list)) return [];
    var seen = {};
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var v = String(list[i] || "");
      if (!v) continue;
      var key = lower(v);
      if (seen[key]) continue;
      seen[key] = true;
      out.push(v);
    }
    return out;
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

  function addressMatchesTarget(address, targetAddressOrDomain) {
    var a = lower(normalizeString(address));
    var t = lower(normalizeString(targetAddressOrDomain));
    if (!a || !t) return false;

    // OutlookOkan semantics are mostly "EndsWith" and sometimes "Equals".
    if (a === t) return true;
    if (endsWith(a, t)) return true;

    // If target is a domain without '@', match domain suffix too.
    if (t.indexOf("@") < 0) {
      var dom = normalizeDomainSuffix(t);
      if (dom && endsWith(a, dom)) return true;
    }

    return false;
  }

  function isInternalAddress(address, internalDomainList) {
    var a = lower(normalizeString(address));
    if (!a) return false;
    for (var i = 0; i < internalDomainList.length; i++) {
      var suf = internalDomainList[i];
      if (!suf) continue;
      if (endsWith(a, suf)) return true;
    }
    return false;
  }

  function isWhitelisted(address, whitelist) {
    var a = lower(normalizeString(address));
    if (!a || !Array.isArray(whitelist) || whitelist.length === 0) return false;
    for (var i = 0; i < whitelist.length; i++) {
      var w = whitelist[i];
      if (!w || !w.whiteName) continue;
      var name = lower(normalizeString(w.whiteName));
      if (!name) continue;
      if (a === name) return true;
      if (endsWith(a, name)) return true;
    }
    return false;
  }

  function isSkipConfirmation(address, whitelist) {
    var a = lower(normalizeString(address));
    if (!a || !Array.isArray(whitelist) || whitelist.length === 0) return false;
    for (var i = 0; i < whitelist.length; i++) {
      var w = whitelist[i];
      if (!w || !w.whiteName) continue;
      var name = lower(normalizeString(w.whiteName));
      if (!name) continue;
      if (a.indexOf(name) >= 0) return !!w.isSkipConfirmation;
    }
    return false;
  }

  function formatRecipient(displayName, emailAddress) {
    var dn = normalizeString(displayName);
    var ea = normalizeString(emailAddress);
    if (!dn) return ea || "";
    if (!ea) return dn;
    if (lower(dn) === lower(ea)) return ea;
    return dn + " <" + ea + ">";
  }

  function normalizeRecipients(list) {
    if (!Array.isArray(list)) return [];
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var r = list[i];
      if (!r) continue;
      var email = normalizeString(r.emailAddress != null ? r.emailAddress : r.address);
      var displayName = normalizeString(r.displayName != null ? r.displayName : r.name);
      var key = lower(email);
      if (!key) continue;
      out.push({
        emailAddress: email,
        displayName: displayName,
        key: key,
        formatted: formatRecipient(displayName, email),
      });
    }
    return out;
  }

  function dedupeRecipients(to, cc, bcc) {
    function filter(list) {
      var seen = {};
      var out = [];
      for (var i = 0; i < list.length; i++) {
        var r = list[i];
        if (!r || !r.key) continue;
        if (seen[r.key]) continue;
        seen[r.key] = true;
        out.push(r);
      }
      return out;
    }

    return {
      to: filter(to),
      cc: filter(cc),
      bcc: filter(bcc),
    };
  }

  function dedupeRecipientList(list) {
    if (!Array.isArray(list)) return [];
    var seen = {};
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var r = list[i];
      if (!r || !r.key) continue;
      if (seen[r.key]) continue;
      seen[r.key] = true;
      out.push(r);
    }
    return out;
  }

  function normalizeExpandedGroups(snapshot, recipientsActual) {
    var resolved = snapshot && snapshot.resolved;
    var groups = [];
    try {
      groups = Array.isArray(resolved && resolved.expandedGroups)
        ? resolved.expandedGroups
        : Array.isArray(resolved && resolved.groups)
          ? resolved.groups
          : [];
    } catch (_e) {
      groups = [];
    }

    var out = { to: [], cc: [], bcc: [], all: [], errors: [] };
    if (!groups || groups.length === 0) return out;

    var present = {};
    try {
      if (recipientsActual) {
        var all = []
          .concat(recipientsActual.to || [])
          .concat(recipientsActual.cc || [])
          .concat(recipientsActual.bcc || []);
        for (var i = 0; i < all.length; i++) {
          if (all[i] && all[i].key) present[all[i].key] = true;
        }
      }
    } catch (_e2) {}

    var seenByField = {
      to: {},
      cc: {},
      bcc: {},
    };
    var seenAll = {};

    function normalizeMembers(list) {
      if (!Array.isArray(list)) return [];
      var temp = [];
      for (var i = 0; i < list.length; i++) {
        var m = list[i];
        if (!m) continue;
        if (typeof m === "string") temp.push({ emailAddress: m, displayName: "" });
        else temp.push(m);
      }
      return normalizeRecipients(temp);
    }

    for (var g = 0; g < groups.length; g++) {
      var group = groups[g];
      if (!group) continue;

      var groupEmail = normalizeString(group.emailAddress != null ? group.emailAddress : group.address);
      var groupKey = lower(groupEmail);
      if (groupKey && Object.keys(present).length > 0 && !present[groupKey]) continue;

      var label = normalizeString(group.displayName) || groupEmail;
      var field = lower(normalizeString(group.field != null ? group.field : group.type));
      var outField = field === "cc" ? "cc" : field === "bcc" ? "bcc" : "to";
      var fieldSeen = seenByField[outField];

      var members = [];
      try {
        members = normalizeMembers(group.members || group.Members || []);
      } catch (_e3) {
        members = [];
      }

      for (var m2 = 0; m2 < members.length; m2++) {
        var mem = members[m2];
        if (!mem || !mem.key) continue;
        var memberForField = mem;
        if (!fieldSeen[mem.key]) {
          fieldSeen[mem.key] = true;
          memberForField = {
            emailAddress: mem.emailAddress,
            displayName: mem.displayName,
            key: mem.key,
            formatted: mem.formatted,
          };

          if (label) {
            memberForField.expandedFrom = label;
            memberForField.formatted = memberForField.formatted + " [" + label + "]";
          }

          if (outField === "cc") out.cc.push(memberForField);
          else if (outField === "bcc") out.bcc.push(memberForField);
          else out.to.push(memberForField);
        }

        if (!seenAll[mem.key]) {
          seenAll[mem.key] = true;
          out.all.push({
            emailAddress: mem.emailAddress,
            displayName: mem.displayName,
            key: mem.key,
            formatted: mem.formatted,
          });
        }
      }
    }

    return out;
  }

  function toEmailList(list) {
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var r = list[i];
      if (!r || !r.emailAddress) continue;
      out.push(r.emailAddress);
    }
    return out;
  }

  function attachmentFileType(fileName) {
    var name = normalizeString(fileName);
    if (!name) return "";
    var dot = name.lastIndexOf(".");
    if (dot < 0) return "";
    return lower(name.slice(dot));
  }

  function formatBytes(bytes) {
    if (typeof bytes !== "number" || !isFinite(bytes) || bytes < 0) return "?";
    var kb = bytes / 1024;
    if (kb < 1024) return Math.round(kb) + "KB";
    var mb = kb / 1024;
    if (mb < 1024) return (Math.round(mb * 10) / 10).toFixed(1) + "MB";
    var gb = mb / 1024;
    return (Math.round(gb * 10) / 10).toFixed(1) + "GB";
  }

  function countRecipientExternalDomains(recipients, senderDomainSuffix, internalDomains, isToAndCcOnly) {
    var domainMap = {};
    var domains = [];

    function addDomain(address) {
      if (!address) return;
      var suf = domainSuffixFromAddress(address);
      if (!suf) return;
      if (domainMap[suf]) return;
      domainMap[suf] = true;
      domains.push(suf);
    }

    if (isToAndCcOnly) {
      for (var i = 0; i < recipients.to.length; i++) addDomain(recipients.to[i].emailAddress);
      for (var j = 0; j < recipients.cc.length; j++) addDomain(recipients.cc[j].emailAddress);
    } else {
      for (var k = 0; k < recipients.to.length; k++) addDomain(recipients.to[k].emailAddress);
      for (var l = 0; l < recipients.cc.length; l++) addDomain(recipients.cc[l].emailAddress);
      for (var m = 0; m < recipients.bcc.length; m++) addDomain(recipients.bcc[m].emailAddress);
    }

    var externalCount = domains.length;
    for (var x = 0; x < internalDomains.length; x++) {
      var internal = internalDomains[x];
      if (!internal) continue;
      var anyMatch = false;
      for (var y = 0; y < domains.length; y++) {
        if (endsWith(domains[y], internal)) {
          anyMatch = true;
          break;
        }
      }
      if (anyMatch && !endsWith(senderDomainSuffix, internal)) {
        externalCount--;
      }
    }

    if (domainMap[senderDomainSuffix]) externalCount--;
    return externalCount;
  }

  function isAllRecipientsInternal(checkList) {
    function allInternal(list) {
      for (var i = 0; i < list.length; i++) {
        if (list[i].isExternal) return false;
      }
      return true;
    }
    return (
      allInternal(checkList.toAddresses) &&
      allInternal(checkList.ccAddresses) &&
      allInternal(checkList.bccAddresses)
    );
  }

  function isAllRecipientsSkip(checkList) {
    function allSkip(list) {
      if (list.length === 0) return true;
      for (var i = 0; i < list.length; i++) {
        if (!list[i].isSkip) return false;
      }
      return true;
    }
    return allSkip(checkList.toAddresses) && allSkip(checkList.ccAddresses) && allSkip(checkList.bccAddresses);
  }

  function isAllChecked(checkList) {
    function allChecked(list) {
      for (var i = 0; i < list.length; i++) {
        if (!list[i].isChecked) return false;
      }
      return true;
    }

    return (
      allChecked(checkList.toAddresses) &&
      allChecked(checkList.ccAddresses) &&
      allChecked(checkList.bccAddresses) &&
      allChecked(checkList.alerts) &&
      allChecked(checkList.attachments)
    );
  }

  function computeShowConfirmation(checkList, settingsGeneral) {
    var g = settingsGeneral || {};

    if (checkList.recipientExternalDomainNumAll >= 2 && !!g.isShowConfirmationToMultipleDomain) return true;
    if (!!g.isDoNotConfirmationIfAllRecipientsAreSameDomain && isAllRecipientsInternal(checkList)) return false;
    if (isAllRecipientsSkip(checkList)) return false;
    if (!!g.isDoDoNotConfirmationIfAllWhite && isAllChecked(checkList)) return false;
    return true;
  }

  function pushAlert(checkList, alertMessage, isImportant, isWhite, isChecked) {
    checkList.alerts.push({
      alertMessage: String(alertMessage || ""),
      isImportant: !!isImportant,
      isWhite: !!isWhite,
      isChecked: !!isChecked,
    });
  }

  function t(locale, key) {
    var ja = startsWith(locale, "ja");
    switch (key) {
      case "forbid":
        return ja ? "送信禁止" : "Send blocked";
      case "forgotAttach":
        return ja ? "添付ファイルの添付漏れの可能性があります。" : "Possible missing attachment.";
      case "largeAttachment":
        return ja ? "大容量の添付ファイルです" : "Large attachment";
      case "dangerousExe":
        return ja ? "実行ファイル(.exe)が添付されています" : "Executable (.exe) attached";
      case "encryptedZip":
        return ja ? "暗号化ZIP(パスワード付きZIP)の可能性があります" : "Possible encrypted ZIP (password-protected ZIP)";
      case "attachmentsProhibited":
        return ja ? "添付ファイル付きメールの送信は禁止されています。" : "Sending with attachments is prohibited.";
      case "attachmentProhibitedRecipients":
        return ja ? "添付ファイル送付禁止の宛先が含まれます" : "Attachments prohibited for recipients";
      case "attachmentAlertRecipients":
        return ja ? "添付ファイル付きメールの宛先に注意してください" : "Caution: attachments with these recipients";
      case "recipientsAndAttachmentsName":
        return ja ? "添付ファイル名と宛先の紐づけに一致しません" : "Attachment name / recipient mapping mismatch";
      case "recommendLink":
        return ja ? "可能であればリンクとして添付することを推奨します。" : "Consider attaching as a link instead of a file.";
      case "keywordAndRecipients":
        return ja ? "本文/件名にキーワードがあるのに宛先が含まれません" : "Keyword present but required recipient missing";
      case "nameDomainMissingInBody":
        return ja ? "本文に紐づく名称が見つかりません" : "Linked name not found in body";
      case "nameDomainMissingInSubject":
        return ja ? "件名に紐づく名称が見つかりません" : "Linked name not found in subject";
      case "maybeIrrelevantRecipient":
        return ja ? "本文(件名)の名称と宛先が一致しない可能性があります" : "Recipients may not match names in content";
      case "externalDomainWarning":
        return ja ? "宛先(To/Cc)の外部ドメイン数が多いです" : "Many external domains in To/Cc";
      case "externalDomainProhibited":
        return ja ? "宛先(To/Cc)の外部ドメイン数が多いため送信禁止です" : "Send blocked: too many external domains in To/Cc";
      case "externalToBccChanged":
        return ja ? "外部宛先(To/Cc)をBccへ自動変換しました" : "Converted external To/Cc recipients to Bcc";
      case "forceToBccChanged":
        return ja ? "宛先を強制的にBccへ変換しました" : "Forced recipients to Bcc";
      case "autoAddSenderToTo":
        return ja ? "Toが空のため送信者をToへ追加しました" : "Added sender to To (To was empty)";
      case "removedRecipients":
        return ja ? "設定により宛先を削除しました" : "Removed recipients by rule";
      case "contactsNotRegisteredWarning":
        return ja ? "連絡先(アドレス帳)未登録の宛先です" : "Recipient not found in Contacts";
      case "contactsNotRegisteredProhibit":
        return ja ? "連絡先(アドレス帳)未登録の宛先があるため送信禁止です" : "Send blocked: recipient not found in Contacts";
      case "contactsLookupUnavailable":
        return ja ? "連絡先(アドレス帳)の確認ができませんでした" : "Couldn't verify recipients in Contacts";
      case "contactsLookupIncomplete":
        return ja ? "連絡先(アドレス帳)の確認が未完了です" : "Contacts verification incomplete";
      case "allRecipientsRemoved":
        return ja ? "宛先がすべて削除されたため送信できません。" : "All recipients were removed; cannot send.";
      case "autoAddedRecipient":
        return ja ? "宛先を自動追加しました" : "Auto-added recipient";
      case "addedTextStart":
        return ja ? "本文の先頭に文言を自動追加します" : "Will prepend text to body";
      case "addedTextEnd":
        return ja ? "本文の末尾に文言を自動追加します" : "Will append text to body";
      default:
        return key;
    }
  }

  function buildInternalDomains(settings, senderDomainSuffix) {
    var list = [];
    if (settings && Array.isArray(settings.internalDomains)) {
      for (var i = 0; i < settings.internalDomains.length; i++) {
        var row = settings.internalDomains[i];
        if (!row) continue;
        var d = normalizeString(row.domain != null ? row.domain : row.Domain);
        var suf = normalizeDomainSuffix(d);
        if (suf) list.push(suf);
      }
    }
    if (senderDomainSuffix) list.push(senderDomainSuffix);
    return uniq(list).map(function (d) {
      return lower(d);
    });
  }

  function shouldIgnoreAttachment(attachment, isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles) {
    if (!attachment) return true;

    var name = normalizeString(attachment.name != null ? attachment.name : attachment.FileName);
    if (!name) return true;

    var fileType = attachmentFileType(name);
    if (fileType === ".p7s" || fileType === "p7s") return true;

    if (!!isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles) {
      try {
        if (attachment.isInline === true) return true;
      } catch (_e) {}
    }

    return false;
  }

  function computeAttachments(checkList, snapshot, settings, locale) {
    var attachmentsSetting = (settings && settings.attachmentsSetting) || {};
    var general = (settings && settings.general) || {};
    var isIgnoreInline = !!general.isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles;
    var runtime = (settings && settings.runtime) || {};
    var largeBytes = typeof runtime.largeAttachmentBytes === "number" ? runtime.largeAttachmentBytes : 10485760;

    var raw = snapshot && Array.isArray(snapshot.attachments) ? snapshot.attachments : [];
    for (var i = 0; i < raw.length; i++) {
      var a = raw[i];
      if (shouldIgnoreAttachment(a, isIgnoreInline)) continue;

      var fileName = normalizeString(a.name != null ? a.name : a.fileName);
      var bytes = typeof a.size === "number" ? a.size : typeof a.fileSizeBytes === "number" ? a.fileSizeBytes : NaN;
      var fileType = attachmentFileType(fileName);

      var isTooBig = typeof bytes === "number" && isFinite(bytes) && bytes >= largeBytes;
      if (isTooBig) {
        pushAlert(checkList, t(locale, "largeAttachment") + " [" + fileName + "]", true, false, false);
      }

      var isDangerous = fileType === ".exe";
      if (isDangerous) {
        pushAlert(checkList, t(locale, "dangerousExe") + " [" + fileName + "]", true, false, false);
      }

      var isEncrypted = false;
      if (
        (!!attachmentsSetting.isWarningWhenEncryptedZipIsAttached ||
          !!attachmentsSetting.isProhibitedWhenEncryptedZipIsAttached) &&
        fileName
      ) {
        // Office.js compose events don't reliably allow reading attachment bytes.
        // For now we treat ".zip" as "possibly encrypted zip" (best-effort).
        if (fileType === ".zip" || fileType === "zip") {
          isEncrypted = true;
          pushAlert(checkList, t(locale, "encryptedZip") + " [" + fileName + "]", true, false, false);

          if (!!attachmentsSetting.isProhibitedWhenEncryptedZipIsAttached) {
            checkList.isCanNotSendMail = true;
            checkList.canNotSendMailMessage = t(locale, "encryptedZip") + " [" + fileName + "]";
          }
        }
      }

      var isChecked = false;
      if (!attachmentsSetting.isMustOpenBeforeCheckTheAttachedFiles) {
        isChecked = !!general.isAutoCheckAttachments;
      }

      checkList.attachments.push({
        fileName: fileName,
        fileType: fileType || "",
        fileSize: formatBytes(bytes),
        fileSizeBytes: bytes,
        isTooBig: isTooBig,
        isDangerous: isDangerous,
        isEncrypted: isEncrypted,
        isChecked: !!isChecked,
      });
    }

    return checkList;
  }

  function checkForgotAttach(checkList, settings, locale) {
    var g = (settings && settings.general) || {};
    if (checkList.attachments.length >= 1) return checkList;
    if (!g.enableForgottenToAttachAlert) return checkList;

    var body = lower(checkList.mailBody);
    if (!body) return checkList;

    // OutlookOkan uses a single localized keyword; we support common variants.
    var keywords = ["添付", "attached file", "attach", "attachment", "attached"];
    for (var i = 0; i < keywords.length; i++) {
      if (body.indexOf(lower(keywords[i])) >= 0) {
        pushAlert(checkList, t(locale, "forgotAttach"), true, false, false);
        break;
      }
    }

    return checkList;
  }

  function checkAlertKeywords(checkList, list, targetText) {
    if (!Array.isArray(list) || list.length === 0) return checkList;
    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      if (!row) continue;
      var keyword = normalizeString(row.alertKeyword != null ? row.alertKeyword : row.AlertKeyword);
      if (!keyword) continue;

      if (keyword !== "*" && String(targetText || "").indexOf(keyword) < 0) continue;

      var message = normalizeString(row.message != null ? row.message : row.Message);
      var isCanNotSend = !!(row.isCanNotSend != null ? row.isCanNotSend : row.IsCanNotSend);
      var alertMessage = message ? message : "Alert keyword [" + keyword + "]";

      pushAlert(checkList, alertMessage, true, false, false);

      if (isCanNotSend) {
        checkList.isCanNotSendMail = true;
        checkList.canNotSendMailMessage = alertMessage;
      }
    }
    return checkList;
  }

  function applyAutoDeleteRecipients(recipients, autoDeleteRecipients, checkList, locale) {
    if (!Array.isArray(autoDeleteRecipients) || autoDeleteRecipients.length === 0) return recipients;

    var patterns = [];
    for (var i = 0; i < autoDeleteRecipients.length; i++) {
      var row = autoDeleteRecipients[i];
      if (!row) continue;
      var p = normalizeString(row.recipient != null ? row.recipient : row.Recipient);
      if (!p) continue;
      patterns.push(lower(p));
    }
    if (patterns.length === 0) return recipients;

    var removedCount = 0;

    function filter(list) {
      var out = [];
      for (var j = 0; j < list.length; j++) {
        var r = list[j];
        if (!r || !r.key) continue;
        var address = r.key;
        var keep = true;

        for (var k = 0; k < patterns.length; k++) {
          var ptn = patterns[k];
          if (!ptn) continue;

          if (startsWith(ptn, "@") && endsWith(address, ptn)) {
            keep = false;
            break;
          }
          if (address === ptn) {
            keep = false;
            break;
          }
        }

        if (keep) out.push(r);
        else removedCount++;
      }
      return out;
    }

    var next = {
      to: filter(recipients.to),
      cc: filter(recipients.cc),
      bcc: filter(recipients.bcc),
    };

    if (removedCount > 0) {
      pushAlert(checkList, t(locale, "removedRecipients"), true, true, true);
    }

    if (next.to.length + next.cc.length + next.bcc.length === 0) {
      checkList.isCanNotSendMail = true;
      checkList.canNotSendMailMessage = t(locale, "allRecipientsRemoved");
    }

    return next;
  }

  function applyAutoCcBccRules(recipients, settings, checkList, locale, externalDomainCountAll, senderEmail, matchExtraKeys) {
    var g = (settings && settings.general) || {};

    var allowKeywordRules = !(
      externalDomainCountAll === 0 && !!g.isDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain
    );
    var allowAttachedFileRules = !(
      externalDomainCountAll === 0 && !!g.isDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain
    );

    var to = recipients.to.slice();
    var cc = recipients.cc.slice();
    var bcc = recipients.bcc.slice();

    var whitelistExtra = [];

    function hasRecipientSubstring(target) {
      var t2 = lower(normalizeString(target));
      if (!t2) return false;
      var all = to.concat(cc).concat(bcc);
      for (var i = 0; i < all.length; i++) {
        if (all[i].key && all[i].key.indexOf(t2) >= 0) return true;
      }
      if (Array.isArray(matchExtraKeys)) {
        for (var j = 0; j < matchExtraKeys.length; j++) {
          var k = matchExtraKeys[j];
          if (k && String(k).indexOf(t2) >= 0) return true;
        }
      }
      return false;
    }

    function addRecipient(field, email, reason) {
      var e = normalizeString(email);
      if (!e || e.indexOf("@") < 0) return;
      var key = lower(e);

      var fieldList = field === "Cc" ? cc : bcc;
      for (var i = 0; i < fieldList.length; i++) {
        if (fieldList[i].key === key) return;
      }

      var r = { emailAddress: e, displayName: "", key: key, formatted: e };
      fieldList.push(r);

      pushAlert(
        checkList,
        t(locale, "autoAddedRecipient") + " [" + field + "] [" + e + "] (" + reason + ")",
        false,
        true,
        true
      );

      whitelistExtra.push({ whiteName: e, isSkipConfirmation: false });
    }

    if (allowKeywordRules && Array.isArray(settings.autoCcBccKeyword) && settings.autoCcBccKeyword.length > 0) {
      for (var i = 0; i < settings.autoCcBccKeyword.length; i++) {
        var row = settings.autoCcBccKeyword[i];
        if (!row) continue;
        var keyword = normalizeString(row.keyword != null ? row.keyword : row.Keyword);
        var target = normalizeString(row.autoAddAddress != null ? row.autoAddAddress : row.AutoAddAddress);
        var ccOrBcc = normalizeString(row.ccOrBcc != null ? row.ccOrBcc : row.CcOrBcc);
        if (!keyword || !target) continue;
        if (String(checkList.mailBody || "").indexOf(keyword) < 0) continue;
        addRecipient(ccOrBcc === "Cc" ? "Cc" : "Bcc", target, "keyword: " + keyword);
      }
    }

    if (
      allowAttachedFileRules &&
      checkList.attachments.length > 0 &&
      Array.isArray(settings.autoCcBccAttachedFile) &&
      settings.autoCcBccAttachedFile.length > 0
    ) {
      for (var j = 0; j < settings.autoCcBccAttachedFile.length; j++) {
        var arow = settings.autoCcBccAttachedFile[j];
        if (!arow) continue;
        var add = normalizeString(arow.autoAddAddress != null ? arow.autoAddAddress : arow.AutoAddAddress);
        var ab = normalizeString(arow.ccOrBcc != null ? arow.ccOrBcc : arow.CcOrBcc);
        if (!add) continue;
        addRecipient(ab === "Cc" ? "Cc" : "Bcc", add, "attachment");
      }
    }

    if (Array.isArray(settings.autoCcBccRecipient) && settings.autoCcBccRecipient.length > 0) {
      for (var k = 0; k < settings.autoCcBccRecipient.length; k++) {
        var rrow = settings.autoCcBccRecipient[k];
        if (!rrow) continue;
        var targetRecipient = normalizeString(
          rrow.targetRecipient != null ? rrow.targetRecipient : rrow.TargetRecipient
        );
        var addAddr = normalizeString(rrow.autoAddAddress != null ? rrow.autoAddAddress : rrow.AutoAddAddress);
        var rb = normalizeString(rrow.ccOrBcc != null ? rrow.ccOrBcc : rrow.CcOrBcc);
        if (!targetRecipient || !addAddr) continue;
        if (!hasRecipientSubstring(targetRecipient)) continue;
        addRecipient(rb === "Cc" ? "Cc" : "Bcc", addAddr, "recipient: " + targetRecipient);
      }
    }

    if (senderEmail && senderEmail.indexOf("@") >= 0) {
      if (!!g.isAutoAddSenderToCc) addRecipient("Cc", senderEmail, "sender");
      if (!!g.isAutoAddSenderToBcc) addRecipient("Bcc", senderEmail, "sender");
    }

    return {
      recipients: dedupeRecipients(to, cc, bcc),
      whitelistExtra: whitelistExtra,
    };
  }

  function externalDomainsChangeToBccIfNeeded(recipients, settings, checkList, locale, internalDomains, senderEmail, senderDomainSuffix) {
    var ex = (settings && settings.externalDomains) || {};
    var force = (settings && settings.forceAutoChangeRecipientsToBcc) || {};

    var threshold = typeof ex.targetToAndCcExternalDomainsNum === "number" ? ex.targetToAndCcExternalDomainsNum : 10;
    var externalDomainNumToAndCc = countRecipientExternalDomains(recipients, senderDomainSuffix, internalDomains, true);

    var shouldForce = !!force.isForceAutoChangeRecipientsToBcc;
    var shouldAuto =
      !!ex.isAutoChangeToBccWhenLargeNumberOfExternalDomains &&
      !ex.isProhibitedWhenLargeNumberOfExternalDomains &&
      threshold <= externalDomainNumToAndCc;

    if (!shouldForce && !shouldAuto) return recipients;

    var internalForConvert = internalDomains.slice();
    if (shouldForce && !!force.isIncludeInternalDomain) {
      internalForConvert = []; // include internal domain in conversion
    }

    var to = [];
    var cc = [];
    var bcc = recipients.bcc.slice();

    function move(list, outList) {
      for (var i = 0; i < list.length; i++) {
        var r = list[i];
        if (!r) continue;
        var isInternal = internalForConvert.length > 0 && isInternalAddress(r.emailAddress, internalForConvert);
        if (isInternal) outList.push(r);
        else bcc.push(r);
      }
    }

    move(recipients.to, to);
    move(recipients.cc, cc);

    if (shouldForce) {
      pushAlert(checkList, t(locale, "forceToBccChanged") + " [" + threshold + "]", false, false, true);
    } else {
      pushAlert(checkList, t(locale, "externalToBccChanged") + " [" + threshold + "]", true, false, false);
    }

    var toRecipient = normalizeString(force.toRecipient);
    var addTo = toRecipient || senderEmail;
    if (to.length === 0 && addTo) {
      to.push({ emailAddress: addTo, displayName: "", key: lower(addTo), formatted: addTo });
      pushAlert(checkList, t(locale, "autoAddSenderToTo"), true, false, false);
    }

    return dedupeRecipients(to, cc, bcc);
  }

  function applyRecipientChecks(checkList, recipients, settings, locale, internalDomains, effectiveWhitelist, expandedByField) {
    var alerts = Array.isArray(settings.alertAddresses) ? settings.alertAddresses : [];

    function addAddress(outList, recipient, seen) {
      var email = recipient.emailAddress;
      var key = lower(email);
      if (key && seen[key]) return;
      if (key) seen[key] = true;
      var formatted = recipient.formatted || email;

      var external = !isInternalAddress(email, internalDomains);
      var white = isWhitelisted(email, effectiveWhitelist);
      var skip = white ? isSkipConfirmation(email, effectiveWhitelist) : false;
      var expandedFrom = normalizeString(recipient.expandedFrom);

      outList.push({
        mailAddress: formatted,
        emailAddress: email,
        isExternal: external,
        isWhite: white,
        isSkip: skip,
        isChecked: white,
        isExpanded: !!expandedFrom,
        expandedFrom: expandedFrom,
      });

      for (var i = 0; i < alerts.length; i++) {
        var row = alerts[i];
        if (!row) continue;
        var target = normalizeString(row.targetAddress != null ? row.targetAddress : row.TargetAddress);
        if (!target) continue;
        if (!addressMatchesTarget(email, target)) continue;

        var isCanNotSend = !!(row.isCanNotSend != null ? row.isCanNotSend : row.IsCanNotSend);
        var message = normalizeString(row.message != null ? row.message : row.Message);

        if (isCanNotSend) {
          checkList.isCanNotSendMail = true;
          checkList.canNotSendMailMessage = t(locale, "forbid") + ": " + formatted;
          continue;
        }

        pushAlert(checkList, (message ? message : "Alert address") + " [" + formatted + "]", true, false, false);
      }
    }

    var seenTo = {};
    var seenCc = {};
    var seenBcc = {};

    for (var i = 0; i < recipients.to.length; i++) addAddress(checkList.toAddresses, recipients.to[i], seenTo);
    for (var j = 0; j < recipients.cc.length; j++) addAddress(checkList.ccAddresses, recipients.cc[j], seenCc);
    for (var k = 0; k < recipients.bcc.length; k++) addAddress(checkList.bccAddresses, recipients.bcc[k], seenBcc);

    var ex = expandedByField || {};
    var exTo = Array.isArray(ex.to) ? ex.to : [];
    var exCc = Array.isArray(ex.cc) ? ex.cc : [];
    var exBcc = Array.isArray(ex.bcc) ? ex.bcc : [];

    for (var e1 = 0; e1 < exTo.length; e1++) addAddress(checkList.toAddresses, exTo[e1], seenTo);
    for (var e2 = 0; e2 < exCc.length; e2++) addAddress(checkList.ccAddresses, exCc[e2], seenCc);
    for (var e3 = 0; e3 < exBcc.length; e3++) addAddress(checkList.bccAddresses, exBcc[e3], seenBcc);

    return checkList;
  }

  function applyContactsChecks(checkList, settings, locale, contactsInfo) {
    var g = (settings && settings.general) || {};

    var autoCheck = !!g.isAutoCheckRegisteredInContacts;
    var warn = !!g.isWarningIfRecipientsIsNotRegistered;
    var prohibit = !!g.isProhibitsSendingMailIfRecipientsIsNotRegistered;

    if (!autoCheck && !warn && !prohibit) return checkList;

    var info = contactsInfo && contactsInfo.resolved && isObject(contactsInfo.resolved) ? contactsInfo.resolved : {};
    var map = info && isObject(info.contacts) ? info.contacts : null;
    var lookupFailed = !!(info && info.contactsLookupFailed);

    if (!map) {
      if (warn || prohibit) {
        pushAlert(checkList, t(locale, "contactsLookupUnavailable"), true, false, false);
      }
      return checkList;
    }

    var unknown = [];

    function handle(list) {
      for (var i = 0; i < list.length; i++) {
        var addr = list[i];
        if (!addr || !addr.emailAddress) continue;
        if (!addr.isExternal) continue; // internal domain is excluded
        var k = lower(addr.emailAddress);
        var v = map[k];

        if (v === true) {
          addr.isRegisteredInContacts = true;
          if (autoCheck) addr.isChecked = true;
          continue;
        }

        if (v === false) {
          addr.isRegisteredInContacts = false;

          if (prohibit) {
            checkList.isCanNotSendMail = true;
            checkList.canNotSendMailMessage = t(locale, "contactsNotRegisteredProhibit") + " [" + addr.mailAddress + "]";
            return false;
          }

          if (warn) {
            pushAlert(
              checkList,
              t(locale, "contactsNotRegisteredWarning") + " [" + addr.mailAddress + "]",
              true,
              false,
              false
            );
          }
          continue;
        }

        addr.isRegisteredInContacts = null;
        unknown.push(addr.mailAddress);
      }
      return true;
    }

    if (!handle(checkList.toAddresses)) return checkList;
    if (!handle(checkList.ccAddresses)) return checkList;
    if (!handle(checkList.bccAddresses)) return checkList;

    if (unknown.length > 0) {
      pushAlert(
        checkList,
        t(locale, "contactsLookupIncomplete") +
          " (" +
          String(unknown.length) +
          "): " +
          unknown.slice(0, 3).join(", ") +
          (unknown.length > 3 ? "..." : ""),
        true,
        false,
        false
      );
    }

    if (lookupFailed) {
      pushAlert(checkList, t(locale, "contactsLookupUnavailable"), true, false, false);
    }

    return checkList;
  }

  function checkRecipientsAndAttachments(checkList, settings, locale) {
    var attachmentsSetting = (settings && settings.attachmentsSetting) || {};

    if (checkList.attachments.length <= 0) return checkList;

    if (!!attachmentsSetting.isAttachmentsProhibited) {
      checkList.isCanNotSendMail = true;
      checkList.canNotSendMailMessage = t(locale, "attachmentsProhibited");
      return checkList;
    }

    var prohibitedList = Array.isArray(settings.attachmentProhibitedRecipients) ? settings.attachmentProhibitedRecipients : [];
    if (prohibitedList.length > 0) {
      var prohibitedRecipients = [];
      var isProhibited = false;

      function scan(list, recipientPattern) {
        var pat = lower(recipientPattern);
        for (var i = 0; i < list.length; i++) {
          var addr = list[i];
          if (!addr || !addr.emailAddress) continue;
          if (lower(addr.emailAddress).indexOf(pat) >= 0) {
            prohibitedRecipients.push(addr.mailAddress);
            isProhibited = true;
          }
        }
      }

      for (var p = 0; p < prohibitedList.length; p++) {
        var pr = prohibitedList[p];
        if (!pr) continue;
        var pat2 = normalizeString(pr.recipient != null ? pr.recipient : pr.Recipient);
        if (!pat2) continue;
        scan(checkList.toAddresses, pat2);
        scan(checkList.ccAddresses, pat2);
        scan(checkList.bccAddresses, pat2);
      }

      if (isProhibited) {
        checkList.isCanNotSendMail = true;
        checkList.canNotSendMailMessage =
          t(locale, "attachmentProhibitedRecipients") + ": " + uniq(prohibitedRecipients).join(" ");
        return checkList;
      }
    }

    var alertRecipients = Array.isArray(settings.attachmentAlertRecipients) ? settings.attachmentAlertRecipients : [];
    if (alertRecipients.length > 0) {
      for (var a = 0; a < alertRecipients.length; a++) {
        var ar = alertRecipients[a];
        if (!ar) continue;
        var target = normalizeString(ar.recipient != null ? ar.recipient : ar.Recipient);
        if (!target) continue;
        var msg = normalizeString(ar.message != null ? ar.message : ar.Message);
        var text = msg ? msg : t(locale, "attachmentAlertRecipients");
        var tLower = lower(target);

        function warn(list) {
          for (var i = 0; i < list.length; i++) {
            if (lower(list[i].emailAddress).indexOf(tLower) >= 0) {
              pushAlert(checkList, text + " [" + list[i].mailAddress + "]", true, false, false);
            }
          }
        }

        warn(checkList.toAddresses);
        warn(checkList.ccAddresses);
        warn(checkList.bccAddresses);
      }
    }

    var mapList = Array.isArray(settings.recipientsAndAttachmentsName) ? settings.recipientsAndAttachmentsName : [];
    if (mapList.length > 0) {
      for (var m = 0; m < mapList.length; m++) {
        var map = mapList[m];
        if (!map) continue;
        var attachmentsName = normalizeString(map.attachmentsName != null ? map.attachmentsName : map.AttachmentsName);
        var recipient = normalizeString(map.recipient != null ? map.recipient : map.Recipient);
        if (!attachmentsName || !recipient) continue;
        var recipientLower = lower(recipient);

        for (var iAtt = 0; iAtt < checkList.attachments.length; iAtt++) {
          var att = checkList.attachments[iAtt];
          if (!att || !att.fileName) continue;
          if (att.fileName.indexOf(attachmentsName) < 0) continue;

          function check(list) {
            for (var i = 0; i < list.length; i++) {
              var addr = list[i];
              if (!addr || !addr.isExternal) continue;
              if (lower(addr.emailAddress).indexOf(recipientLower) >= 0) continue;
              pushAlert(
                checkList,
                t(locale, "recipientsAndAttachmentsName") + ": " + addr.mailAddress + " / " + att.fileName,
                true,
                true,
                false
              );
            }
          }

          check(checkList.toAddresses);
          check(checkList.ccAddresses);
          check(checkList.bccAddresses);
        }
      }
    }

    if (!!attachmentsSetting.isWarningWhenAttachedRealFile) {
      pushAlert(checkList, t(locale, "recommendLink"), false, true, false);
    }

    return checkList;
  }

  function checkNameAndDomains(checkList, recipientsAll, settings, locale) {
    var g = (settings && settings.general) || {};
    var list = Array.isArray(settings.nameAndDomains) ? settings.nameAndDomains : [];
    if (list.length === 0) return checkList;

    var cleaned = [];
    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      if (!row) continue;
      var name = normalizeString(row.name != null ? row.name : row.Name);
      var domain = normalizeString(row.domain != null ? row.domain : row.Domain);
      if (!name || !domain) continue;
      cleaned.push({ name: name, domain: domain });
    }
    if (cleaned.length === 0) return checkList;

    var includeSubject = !!g.isCheckNameAndDomainsIncludeSubject;

    // 1) If requested, warn when linked name is not found for recipients in that domain.
    if (!!g.isCheckNameAndDomainsFromRecipients || (includeSubject && !!g.isCheckNameAndDomainsFromSubject)) {
      var candidates = [];
      for (var c = 0; c < cleaned.length; c++) {
        for (var r = 0; r < recipientsAll.length; r++) {
          var addr = recipientsAll[r];
          if (!addr || !addr.emailAddress) continue;
          if (addressMatchesTarget(addr.emailAddress, cleaned[c].domain)) {
            candidates.push([addr.formatted, cleaned[c].name]);
          }
        }
      }

      if (candidates.length > 0 && !!g.isCheckNameAndDomainsFromRecipients) {
        var missCount = 0;
        for (var i2 = 0; i2 < candidates.length; i2++) {
          var pair = candidates[i2];
          if (String(checkList.mailBody || "").indexOf(pair[1]) < 0 && pair[0].indexOf(checkList.senderDomain) < 0) {
            missCount++;
          }
        }
        if (missCount >= candidates.length) {
          for (var i3 = 0; i3 < candidates.length; i3++) {
            var pair2 = candidates[i3];
            if (String(checkList.mailBody || "").indexOf(pair2[1]) < 0 && pair2[0].indexOf(checkList.senderDomain) < 0) {
              pushAlert(
                checkList,
                pair2[0] + " : " + t(locale, "nameDomainMissingInBody") + " (" + pair2[1] + ")",
                true,
                false,
                false
              );
            }
          }
        }
      }

      if (candidates.length > 0 && includeSubject && !!g.isCheckNameAndDomainsFromSubject) {
        var missSub = 0;
        for (var i4 = 0; i4 < candidates.length; i4++) {
          var pair3 = candidates[i4];
          if (String(checkList.subject || "").indexOf(pair3[1]) < 0 && pair3[0].indexOf(checkList.senderDomain) < 0) {
            missSub++;
          }
        }
        if (missSub >= candidates.length) {
          for (var i5 = 0; i5 < candidates.length; i5++) {
            var pair4 = candidates[i5];
            if (String(checkList.subject || "").indexOf(pair4[1]) < 0 && pair4[0].indexOf(checkList.senderDomain) < 0) {
              pushAlert(
                checkList,
                pair4[0] + " : " + t(locale, "nameDomainMissingInSubject") + " (" + pair4[1] + ")",
                true,
                false,
                false
              );
            }
          }
        }
      }
    }

    // 2) If any names are found in text, warn for recipients outside candidate domains.
    var targetText = String(checkList.mailBody || "");
    if (includeSubject) targetText += String(checkList.subject || "");

    var candidateDomains = [];
    for (var j = 0; j < cleaned.length; j++) {
      if (targetText.indexOf(cleaned[j].name) >= 0) candidateDomains.push(cleaned[j].domain);
    }
    if (candidateDomains.length === 0) return checkList;

    for (var r2 = 0; r2 < recipientsAll.length; r2++) {
      var rec = recipientsAll[r2];
      if (!rec || !rec.emailAddress) continue;
      var ok = false;
      for (var d = 0; d < candidateDomains.length; d++) {
        if (addressMatchesTarget(rec.emailAddress, candidateDomains[d])) {
          ok = true;
          break;
        }
      }
      if (ok) continue;
      if (rec.key.indexOf(checkList.senderDomain) >= 0) continue;

      pushAlert(checkList, rec.formatted + " : " + t(locale, "maybeIrrelevantRecipient"), true, false, false);
    }

    return checkList;
  }

  function checkKeywordAndRecipients(checkList, recipientsAll, settings, locale) {
    var g = (settings && settings.general) || {};
    var list = Array.isArray(settings.keywordAndRecipients) ? settings.keywordAndRecipients : [];
    if (list.length === 0) return checkList;

    var cleaned = [];
    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      if (!row) continue;
      var keyword = normalizeString(row.keyword != null ? row.keyword : row.Keyword);
      var recipient = normalizeString(row.recipient != null ? row.recipient : row.Recipient);
      if (!keyword || !recipient) continue;
      cleaned.push({ keyword: keyword, recipient: recipient });
    }
    if (cleaned.length === 0) return checkList;

    var targetText = String(checkList.mailBody || "");
    if (!!g.isCheckKeywordAndRecipientsIncludeSubject) targetText += String(checkList.subject || "");

    for (var j = 0; j < cleaned.length; j++) {
      var rule = cleaned[j];
      if (targetText.indexOf(rule.keyword) < 0) continue;

      var has = false;
      for (var r = 0; r < recipientsAll.length; r++) {
        if (addressMatchesTarget(recipientsAll[r].emailAddress, rule.recipient)) {
          has = true;
          break;
        }
      }
      if (!has) {
        pushAlert(
          checkList,
          t(locale, "keywordAndRecipients") + ": " + rule.keyword + " -> " + rule.recipient,
          true,
          false,
          false
        );
      }
    }

    return checkList;
  }

  function externalDomainsWarningIfNeeded(checkList, settings, locale, externalDomainNumToAndCc) {
    var ex = (settings && settings.externalDomains) || {};
    var force = (settings && settings.forceAutoChangeRecipientsToBcc) || {};
    if (!!force.isForceAutoChangeRecipientsToBcc) return checkList;

    var threshold = typeof ex.targetToAndCcExternalDomainsNum === "number" ? ex.targetToAndCcExternalDomainsNum : 10;
    if (threshold > externalDomainNumToAndCc) return checkList;

    if (!!ex.isProhibitedWhenLargeNumberOfExternalDomains) {
      checkList.isCanNotSendMail = true;
      checkList.canNotSendMailMessage = t(locale, "externalDomainProhibited") + " [" + threshold + "]";
      return checkList;
    }

    if (!!ex.isWarningWhenLargeNumberOfExternalDomains && !ex.isAutoChangeToBccWhenLargeNumberOfExternalDomains) {
      pushAlert(checkList, t(locale, "externalDomainWarning") + " [" + threshold + "]", true, false, false);
    }

    return checkList;
  }

  function applyAutoAddMessagePreview(checkList, settings, locale) {
    var autoAdd = (settings && settings.autoAddMessage) || {};
    if (!autoAdd || (!autoAdd.isAddToStart && !autoAdd.isAddToEnd)) return checkList;
    if (checkList.mailBody == null) return checkList;

    if (!!autoAdd.isAddToStart && normalizeString(autoAdd.messageOfAddToStart)) {
      pushAlert(checkList, t(locale, "addedTextStart"), false, false, true);
    }
    if (!!autoAdd.isAddToEnd && normalizeString(autoAdd.messageOfAddToEnd)) {
      pushAlert(checkList, t(locale, "addedTextEnd"), false, false, true);
    }
    return checkList;
  }

  function defaultLocale(settings, snapshot) {
    var general = (settings && settings.general) || {};
    var lang = normalizeString(general.languageCode);
    if (lang) return lang;
    var dl = normalizeString(snapshot && snapshot.displayLanguage);
    if (dl) return dl;
    return "en-US";
  }

  function evaluate(snapshot, settings) {
    if (!isObject(snapshot)) snapshot = {};
    if (!isObject(settings)) settings = {};

    var locale = lower(defaultLocale(settings, snapshot));

    var senderEmail = normalizeString(
      snapshot.senderEmailAddress ||
        (snapshot.sender && snapshot.sender.emailAddress) ||
        (snapshot.sender && snapshot.sender.address) ||
        ""
    );
    var senderDomainSuffix = domainSuffixFromAddress(senderEmail);

    var internalDomains = buildInternalDomains(settings, senderDomainSuffix);

    var to = normalizeRecipients(snapshot.to || (snapshot.recipients && snapshot.recipients.to) || []);
    var cc = normalizeRecipients(snapshot.cc || (snapshot.recipients && snapshot.recipients.cc) || []);
    var bcc = normalizeRecipients(snapshot.bcc || (snapshot.recipients && snapshot.recipients.bcc) || []);
    var recipients = dedupeRecipients(to, cc, bcc);

    var checkList = {
      alerts: [],
      toAddresses: [],
      ccAddresses: [],
      bccAddresses: [],
      attachments: [],
      sender: senderEmail,
      senderDomain: senderDomainSuffix || "",
      recipientExternalDomainNumAll: 0,
      subject: normalizeString(snapshot.subject),
      mailType: normalizeString(snapshot.itemType) || "Message",
      mailBody: normalizeString(snapshot.bodyText),
      isCanNotSendMail: false,
      canNotSendMailMessage: "",
    };

    var g = (settings && settings.general) || {};
    var itemType = lower(checkList.mailType);
    if (itemType === "appointment" && !g.isShowConfirmationAtSendMeetingRequest) {
      return {
        checkList: checkList,
        mutations: { to: toEmailList(recipients.to), cc: toEmailList(recipients.cc), bcc: toEmailList(recipients.bcc) },
        showConfirmation: false,
        locale: locale,
      };
    }

    recipients = applyAutoDeleteRecipients(recipients, settings.autoDeleteRecipients, checkList, locale);
    if (checkList.isCanNotSendMail) {
      return {
        checkList: checkList,
        mutations: { to: toEmailList(recipients.to), cc: toEmailList(recipients.cc), bcc: toEmailList(recipients.bcc) },
        showConfirmation: true,
        locale: locale,
      };
    }

    var expanded = normalizeExpandedGroups(snapshot, recipients);

    computeAttachments(checkList, snapshot, settings, locale);
    checkForgotAttach(checkList, settings, locale);
    checkAlertKeywords(checkList, settings.alertKeywordsBody, checkList.mailBody);
    checkAlertKeywords(checkList, settings.alertKeywordsSubject, checkList.subject);

    applyAutoAddMessagePreview(checkList, settings, locale);

    var recipientsForCounts = {
      to: recipients.to.concat(expanded.to),
      cc: recipients.cc.concat(expanded.cc),
      bcc: recipients.bcc.concat(expanded.bcc),
    };

    var externalDomainCountAll = countRecipientExternalDomains(
      recipientsForCounts,
      senderDomainSuffix,
      internalDomains,
      false
    );

    var expandedKeys = [];
    for (var ek = 0; ek < expanded.all.length; ek++) {
      if (expanded.all[ek] && expanded.all[ek].key) expandedKeys.push(expanded.all[ek].key);
    }

    var autoAdd = applyAutoCcBccRules(
      recipients,
      settings,
      checkList,
      locale,
      externalDomainCountAll,
      senderEmail,
      expandedKeys
    );
    recipients = autoAdd.recipients;

    recipients = externalDomainsChangeToBccIfNeeded(
      recipients,
      settings,
      checkList,
      locale,
      internalDomains,
      senderEmail,
      senderDomainSuffix
    );

    var effectiveWhitelist = []
      .concat(Array.isArray(settings.whitelist) ? settings.whitelist : [])
      .concat(Array.isArray(autoAdd.whitelistExtra) ? autoAdd.whitelistExtra : []);

    if (!!g.contactGroupMembersAreWhite || !!g.exchangeDistributionListMembersAreWhite) {
      for (var wl = 0; wl < expanded.all.length; wl++) {
        if (!expanded.all[wl] || !expanded.all[wl].emailAddress) continue;
        effectiveWhitelist.push({ whiteName: expanded.all[wl].emailAddress, isSkipConfirmation: false });
      }
    }

    applyRecipientChecks(checkList, recipients, settings, locale, internalDomains, effectiveWhitelist, expanded);

    if (!!g.isAutoCheckIfAllRecipientsAreSameDomain) {
      var i;
      for (i = 0; i < checkList.toAddresses.length; i++) if (!checkList.toAddresses[i].isExternal) checkList.toAddresses[i].isChecked = true;
      for (i = 0; i < checkList.ccAddresses.length; i++) if (!checkList.ccAddresses[i].isExternal) checkList.ccAddresses[i].isChecked = true;
      for (i = 0; i < checkList.bccAddresses.length; i++) if (!checkList.bccAddresses[i].isExternal) checkList.bccAddresses[i].isChecked = true;
    }

    applyContactsChecks(checkList, settings, locale, snapshot);
    if (checkList.isCanNotSendMail) {
      return {
        checkList: checkList,
        mutations: { to: toEmailList(recipients.to), cc: toEmailList(recipients.cc), bcc: toEmailList(recipients.bcc) },
        showConfirmation: true,
        locale: locale,
      };
    }

    checkRecipientsAndAttachments(checkList, settings, locale);

    var allRecipients = dedupeRecipientList(recipients.to.concat(recipients.cc).concat(recipients.bcc).concat(expanded.all));
    checkNameAndDomains(checkList, allRecipients, settings, locale);
    checkKeywordAndRecipients(checkList, allRecipients, settings, locale);

    var recipientsForCounts2 = {
      to: recipients.to.concat(expanded.to),
      cc: recipients.cc.concat(expanded.cc),
      bcc: recipients.bcc.concat(expanded.bcc),
    };

    var externalDomainNumToAndCc = countRecipientExternalDomains(recipientsForCounts2, senderDomainSuffix, internalDomains, true);
    checkList.recipientExternalDomainNumAll = countRecipientExternalDomains(
      recipientsForCounts2,
      senderDomainSuffix,
      internalDomains,
      false
    );
    externalDomainsWarningIfNeeded(checkList, settings, locale, externalDomainNumToAndCc);

    var showConfirmation = computeShowConfirmation(checkList, g);

    if (!!g.isEnableRecipientsAreSortedByDomain) {
      function sortByDomain(a, b) {
        var da = domainSuffixFromAddress(a.emailAddress);
        var db = domainSuffixFromAddress(b.emailAddress);
        if (da < db) return -1;
        if (da > db) return 1;
        return a.emailAddress < b.emailAddress ? -1 : a.emailAddress > b.emailAddress ? 1 : 0;
      }
      checkList.toAddresses = checkList.toAddresses.slice().sort(sortByDomain);
      checkList.ccAddresses = checkList.ccAddresses.slice().sort(sortByDomain);
      checkList.bccAddresses = checkList.bccAddresses.slice().sort(sortByDomain);
    }

    return {
      checkList: checkList,
      mutations: { to: toEmailList(recipients.to), cc: toEmailList(recipients.cc), bcc: toEmailList(recipients.bcc) },
      showConfirmation: showConfirmation,
      locale: locale,
    };
  }

  ns.engine = {
    evaluate: evaluate,
    _domainSuffixFromAddress: domainSuffixFromAddress,
    _countRecipientExternalDomains: countRecipientExternalDomains,
  };
})();
