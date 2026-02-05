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

    return {
      displayLanguage: displayLanguage,
      senderEmailAddress: senderEmail,
      itemType: itemType,
      subject: subject,
      bodyText: bodyTextForChecks,
      bodyTextRaw: bodyTextRaw,
      recipients: { to: to, cc: cc, bcc: bcc },
      attachments: attachments,
    };
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
