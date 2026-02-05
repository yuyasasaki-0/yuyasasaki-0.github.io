"use strict";
(function () {
  var root =
    typeof globalThis !== "undefined" ? globalThis : typeof window !== "undefined" ? window : undefined;
  if (!root) return;

  var ns = (root.MailChecker = root.MailChecker || {});

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
    // uniq + lower
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

  async function enrichSnapshotWithEws(snapshot, settings, timeoutMs) {
    if (!snapshot || !settings) return snapshot;

    var g = (settings && settings.general) || {};
    var wantDl =
      !!g.enableGetContactGroupMembers ||
      !!g.enableGetExchangeDistributionListMembers;
    var wantContacts =
      !!g.isAutoCheckRegisteredInContacts ||
      !!g.isWarningIfRecipientsIsNotRegistered ||
      !!g.isProhibitsSendingMailIfRecipientsIsNotRegistered;

    if (!wantDl && !wantContacts) return snapshot;

    var ews = ns.ews;
    var canEws = !!(ews && typeof ews.canUseEws === "function" && ews.canUseEws());

    var resolved = { expandedGroups: [], contacts: null, contactsLookupFailed: false };

    if (!canEws) {
      if (wantContacts) resolved.contactsLookupFailed = true;
      snapshot.resolved = resolved;
      return snapshot;
    }

    var internalDomains = buildInternalDomains(settings, domainSuffixFromAddress(snapshot.senderEmailAddress));

    var deadline = 0;
    try {
      var maxBudget = Math.min(2200, Math.max(600, (timeoutMs || 0) - 1800));
      deadline = Date.now() + maxBudget;
    } catch (_e2) {
      deadline = 0;
    }

    function timeLeft() {
      try {
        return deadline ? deadline - Date.now() : 0;
      } catch (_e3) {
        return 0;
      }
    }

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
        collectDlCandidates(snapshot.recipients && snapshot.recipients.to ? snapshot.recipients.to : [], "To", candidates, seenDl);
        collectDlCandidates(snapshot.recipients && snapshot.recipients.cc ? snapshot.recipients.cc : [], "Cc", candidates, seenDl);
        collectDlCandidates(snapshot.recipients && snapshot.recipients.bcc ? snapshot.recipients.bcc : [], "Bcc", candidates, seenDl);

        for (var c = 0; c < candidates.length && c < 3; c++) {
          if (timeLeft() < 300) break;
          var cand = candidates[c];
          var perTimeout = Math.max(300, Math.min(900, timeLeft()));
          var r1 = await ews.expandDlCached(cand.emailAddress, { timeoutMs: perTimeout });
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
      } catch (_e4) {}
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
        for (var q = 0; q < uniq.length; q++) {
          if (timeLeft() < 250) break;
          var email2 = uniq[q];
          var perTimeout2 = Math.max(250, Math.min(700, timeLeft()));
          var r2 = await ews.resolveInContactsCached(email2, { timeoutMs: perTimeout2 });
          if (r2 && r2.ok && typeof r2.value === "boolean") {
            contacts[lower(email2)] = r2.value;
          } else if (r2 && !r2.ok) {
            resolved.contactsLookupFailed = true;
          }
        }

        resolved.contacts = contacts;
      } catch (_e5) {
        resolved.contactsLookupFailed = true;
        resolved.contacts = null;
      }
    }

    snapshot.resolved = resolved;
    return snapshot;
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
            if (!result || result.status !== Office.AsyncResultStatus.Succeeded) return resolve("");
            resolve(normalizeString(result.value));
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
            if (!result || result.status !== Office.AsyncResultStatus.Succeeded) return resolve("");
            resolve(normalizeString(result.value));
          } catch (_e) {
            resolve("");
          }
        });
      } catch (_e2) {
        resolve("");
      }
    });
  }

  function escapeHtml(text) {
    return String(text || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function pPrependBody(item, text, asHtml) {
    return new Promise(function (resolve) {
      try {
        if (!item || !item.body || typeof item.body.prependAsync !== "function") return resolve(false);
        var options = { coercionType: asHtml ? Office.CoercionType.Html : Office.CoercionType.Text };
        item.body.prependAsync(text, options, function (result) {
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

  function pAppendBody(item, text, asHtml) {
    return new Promise(function (resolve) {
      try {
        if (!item || !item.body || typeof item.body.appendAsync !== "function") return resolve(false);
        var options = { coercionType: asHtml ? Office.CoercionType.Html : Office.CoercionType.Text };
        item.body.appendAsync(text, options, function (result) {
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

  function pGetBodyType(item) {
    return new Promise(function (resolve) {
      try {
        if (!item || !item.body || typeof item.body.getTypeAsync !== "function") return resolve("text");
        item.body.getTypeAsync(function (result) {
          try {
            if (!result || result.status !== Office.AsyncResultStatus.Succeeded) return resolve("text");
            var t = String(result.value || "").toLowerCase();
            resolve(t);
          } catch (_e) {
            resolve("text");
          }
        });
      } catch (_e2) {
        resolve("text");
      }
    });
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

  async function buildSnapshot(item, settings, timeoutMs) {
    var displayLanguage = "";
    try {
      displayLanguage = normalizeString(Office.context && Office.context.displayLanguage);
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

    var itemType = "";
    try {
      itemType = normalizeString(item && item.itemType);
    } catch (_e3) {}

    // Recipients
    var to = [];
    var cc = [];
    var bcc = [];

    if (lower(itemType) === "appointment") {
      var required = await withTimeout(pGetRecipients(item.requiredAttendees), timeoutMs, []);
      var optional = await withTimeout(pGetRecipients(item.optionalAttendees), timeoutMs, []);
      to = required;
      cc = optional;
      bcc = [];
    } else {
      to = await withTimeout(pGetRecipients(item.to), timeoutMs, []);
      cc = await withTimeout(pGetRecipients(item.cc), timeoutMs, []);
      bcc = await withTimeout(pGetRecipients(item.bcc), timeoutMs, []);
    }

    var subject = await withTimeout(pGetSubject(item), timeoutMs, "");
    var bodyTextRaw = await withTimeout(pGetBodyText(item), timeoutMs, "");
    var bodyTextForChecks = bodyTextRaw;

    // Add auto-add message text to the evaluation text (OutlookOkan preview behavior)
    try {
      bodyTextForChecks = applyAutoAddMessagePreviewToText(bodyTextRaw, settings && settings.autoAddMessage);
    } catch (_e4) {}

    // Attachments metadata
    var attachments = [];
    try {
      attachments = Array.isArray(item.attachments) ? item.attachments : [];
    } catch (_e5) {
      attachments = [];
    }

    var snap = {
      displayLanguage: displayLanguage,
      senderEmailAddress: senderEmail,
      itemType: itemType,
      subject: subject,
      bodyText: bodyTextForChecks,
      bodyTextRaw: bodyTextRaw,
      recipients: { to: to, cc: cc, bcc: bcc },
      attachments: attachments,
    };

    try {
      await enrichSnapshotWithEws(snap, settings, timeoutMs);
    } catch (_e6) {}

    return snap;
  }

  function sameEmailLists(a, b) {
    if (!Array.isArray(a) || !Array.isArray(b)) return false;
    if (a.length !== b.length) return false;
    for (var i = 0; i < a.length; i++) {
      if (lower(a[i]) !== lower(b[i])) return false;
    }
    return true;
  }

  async function applyRecipientMutations(item, itemType, mutations, timeoutMs) {
    if (!mutations) return;

    var toEmails = Array.isArray(mutations.to) ? mutations.to : [];
    var ccEmails = Array.isArray(mutations.cc) ? mutations.cc : [];
    var bccEmails = Array.isArray(mutations.bcc) ? mutations.bcc : [];

    if (lower(itemType) === "appointment") {
      await withTimeout(pSetRecipients(item.requiredAttendees, toEmails), timeoutMs, false);
      await withTimeout(pSetRecipients(item.optionalAttendees, ccEmails), timeoutMs, false);
      return;
    }

    await withTimeout(pSetRecipients(item.to, toEmails), timeoutMs, false);
    await withTimeout(pSetRecipients(item.cc, ccEmails), timeoutMs, false);
    await withTimeout(pSetRecipients(item.bcc, bccEmails), timeoutMs, false);
  }

  async function applyAutoAddMessageToBody(item, settings, snapshot, timeoutMs) {
    if (!item || !settings || !settings.autoAddMessage) return;
    if (lower(snapshot.itemType) === "appointment") return;

    var autoAdd = settings.autoAddMessage;
    var startMsg = normalizeString(autoAdd.messageOfAddToStart);
    var endMsg = normalizeString(autoAdd.messageOfAddToEnd);
    var bodyText = normalizeString(snapshot.bodyTextRaw != null ? snapshot.bodyTextRaw : snapshot.bodyText);

    var bodyType = await withTimeout(pGetBodyType(item), timeoutMs, "text");
    var isHtml = String(bodyType || "").toLowerCase() === "html";

    if (autoAdd.isAddToStart && startMsg) {
      var already = bodyText.indexOf(startMsg) === 0;
      if (!already) {
        var textToInsert = isHtml ? "<div>" + escapeHtml(startMsg).replace(/\r?\n/g, "<br>") + "</div><br>" : startMsg + "\n\n";
        await withTimeout(pPrependBody(item, textToInsert, isHtml), timeoutMs, false);
      }
    }

    if (autoAdd.isAddToEnd && endMsg) {
      var alreadyEnd = bodyText.lastIndexOf(endMsg) === bodyText.length - endMsg.length;
      if (!alreadyEnd) {
        var textToAppend = isHtml ? "<br><div>" + escapeHtml(endMsg).replace(/\r?\n/g, "<br>") + "</div>" : "\n\n" + endMsg;
        await withTimeout(pAppendBody(item, textToAppend, isHtml), timeoutMs, false);
      }
    }
  }

  function buildSmartAlertMessage(result) {
    var locale = (result && result.locale) || "en-US";
    var ja = startsWith(lower(locale), "ja");
    var cl = result && result.checkList ? result.checkList : null;
    if (!cl) return ja ? "確認が必要です。" : "Review required.";

    if (cl.isCanNotSendMail) {
      return normalizeString(cl.canNotSendMailMessage) || (ja ? "送信禁止です。" : "Send blocked.");
    }

    var lines = [];
    lines.push(ja ? "送信前に確認してください。" : "Please review before sending.");

    try {
      var ext = [];
      function addExt(list, label) {
        for (var i = 0; i < list.length; i++) {
          if (list[i].isExternal) ext.push(label + ": " + list[i].mailAddress);
        }
      }
      addExt(cl.toAddresses || [], "To");
      addExt(cl.ccAddresses || [], "Cc");
      addExt(cl.bccAddresses || [], "Bcc");
      if (ext.length > 0) {
        lines.push((ja ? "外部宛先" : "External recipients") + ": " + ext.length);
        lines = lines.concat(ext.slice(0, 6));
        if (ext.length > 6) lines.push("...");
      }
    } catch (_e) {}

    try {
      var atts = cl.attachments || [];
      if (atts.length > 0) {
        lines.push((ja ? "添付" : "Attachments") + ": " + atts.length);
        for (var i2 = 0; i2 < Math.min(6, atts.length); i2++) {
          lines.push("- " + atts[i2].fileName);
        }
        if (atts.length > 6) lines.push("...");
      }
    } catch (_e2) {}

    try {
      var alerts = (cl.alerts || []).filter(function (a) {
        return a && a.isImportant && !a.isChecked;
      });
      if (alerts.length > 0) {
        lines.push(ja ? "警告:" : "Warnings:");
        for (var i3 = 0; i3 < Math.min(5, alerts.length); i3++) {
          lines.push("- " + alerts[i3].alertMessage);
        }
        if (alerts.length > 5) lines.push("...");
      }
    } catch (_e3) {}

    lines.push(ja ? "送信しますか？" : "Send anyway?");
    return lines.join("\n");
  }

  function completeAllow(event) {
    try {
      event.completed({ allowEvent: true });
    } catch (_e) {}
  }

  function completeBlock(event, message, promptUser) {
    try {
      var opts = { allowEvent: false, errorMessage: String(message || "") };
      if (promptUser && Office && Office.MailboxEnums && Office.MailboxEnums.SendModeOverride) {
        opts.sendModeOverride = Office.MailboxEnums.SendModeOverride.PromptUser;
      }
      event.completed(opts);
    } catch (_e) {
      try {
        event.completed({ allowEvent: false });
      } catch (_e2) {}
    }
  }

  async function messageOnSent(event) {
    if (!event || typeof event.completed !== "function") return;

    var item = null;
    try {
      item = Office.context && Office.context.mailbox && Office.context.mailbox.item;
    } catch (_e) {}
    if (!item) return completeAllow(event);

    var settings = null;
    try {
      if (ns.settings && typeof ns.settings.load === "function") {
        settings = await ns.settings.load();
      }
    } catch (_e2) {
      settings = null;
    }
    if (!settings && ns.settings && typeof ns.settings.defaults === "function") {
      try {
        settings = ns.settings.defaults();
      } catch (_e3) {}
    }
    if (!settings) settings = {};

    var timeoutMs = 4500;
    try {
      if (settings.runtime && typeof settings.runtime.sendEventTimeoutMs === "number") {
        timeoutMs = settings.runtime.sendEventTimeoutMs;
      }
    } catch (_e4) {}

    try {
      var snapshot = await buildSnapshot(item, settings, timeoutMs);
      var result = ns.engine && typeof ns.engine.evaluate === "function" ? ns.engine.evaluate(snapshot, settings) : null;
      if (!result || !result.checkList) {
        return completeBlock(
          event,
          "Mail Checker failed to evaluate the message. Please try again.",
          true
        );
      }

      // Apply mutations (recipients)
      await applyRecipientMutations(item, snapshot.itemType, result.mutations, timeoutMs);

      // Apply auto-add message to actual body (best-effort)
      await applyAutoAddMessageToBody(item, settings, snapshot, timeoutMs);

      if (result.checkList.isCanNotSendMail) {
        return completeBlock(event, result.checkList.canNotSendMailMessage || "Send blocked.", false);
      }

      if (result.showConfirmation) {
        return completeBlock(event, buildSmartAlertMessage(result), true);
      }

      return completeAllow(event);
    } catch (_err) {
      return completeBlock(
        event,
        "Mail Checker failed to run checks due to an unexpected error. Send anyway?",
        true
      );
    }
  }

  function registerHandlers() {
    try {
      if (typeof Office === "undefined") return;
      if (Office.actions && typeof Office.actions.associate === "function") {
        Office.actions.associate("messageOnSent", messageOnSent);
      }
    } catch (_e) {}
  }

  try {
    if (typeof Office !== "undefined" && Office.onReady && typeof Office.onReady === "function") {
      Office.onReady(registerHandlers);
    } else {
      registerHandlers();
    }
  } catch (_e) {
    registerHandlers();
  }
})();
