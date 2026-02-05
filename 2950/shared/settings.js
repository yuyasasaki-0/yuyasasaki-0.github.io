"use strict";
(function () {
  var root =
    typeof globalThis !== "undefined" ? globalThis : typeof window !== "undefined" ? window : undefined;
  if (!root) return;

  var ns = (root.MailChecker = root.MailChecker || {});
  var storage = ns.storage;

  var SETTINGS_KEY = "mailchecker.settings.v1";
  var SCHEMA_VERSION = 1;

  function deepClone(obj) {
    return obj == null ? obj : JSON.parse(JSON.stringify(obj));
  }

  function isObject(value) {
    return value != null && typeof value === "object" && !Array.isArray(value);
  }

  function normalizeString(value) {
    return value == null ? "" : String(value).trim();
  }

  function normalizeBool(value, fallback) {
    if (typeof value === "boolean") return value;
    if (typeof value === "number") return value !== 0;
    if (typeof value === "string") {
      var v = value.trim().toLowerCase();
      if (v === "yes" || v === "y" || v === "true" || v === "1") return true;
      if (v === "no" || v === "n" || v === "false" || v === "0") return false;
    }
    return !!fallback;
  }

  function normalizeCcOrBcc(value) {
    var v = normalizeString(value).toLowerCase();
    if (v === "cc") return "Cc";
    if (v === "bcc") return "Bcc";
    return "Bcc";
  }

  function mergeDefaults(target, defaults) {
    if (!isObject(target)) target = {};
    if (!isObject(defaults)) return target;

    Object.keys(defaults).forEach(function (key) {
      var d = defaults[key];
      var t = target[key];
      if (t == null) {
        target[key] = deepClone(d);
        return;
      }
      if (Array.isArray(d)) {
        if (!Array.isArray(t)) target[key] = deepClone(d);
        return;
      }
      if (isObject(d)) {
        target[key] = mergeDefaults(isObject(t) ? t : {}, d);
        return;
      }
    });

    return target;
  }

  function defaultSettings() {
    return {
      schemaVersion: SCHEMA_VERSION,
      general: {
        languageCode: "",
        enableForgottenToAttachAlert: true,
        isDoNotConfirmationIfAllRecipientsAreSameDomain: false,
        isDoDoNotConfirmationIfAllWhite: false,
        isAutoCheckIfAllRecipientsAreSameDomain: false,
        isShowConfirmationToMultipleDomain: false,
        enableGetContactGroupMembers: false,
        enableGetExchangeDistributionListMembers: false,
        contactGroupMembersAreWhite: true,
        exchangeDistributionListMembersAreWhite: true,
        isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles: false,
        isDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain: false,
        isDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain: false,
        isDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain: false,
        isEnableRecipientsAreSortedByDomain: false,
        isAutoAddSenderToBcc: false,
        isAutoAddSenderToCc: false,
        isAutoCheckAttachments: false,
        isCheckNameAndDomainsFromRecipients: false,
        isCheckNameAndDomainsIncludeSubject: false,
        isCheckNameAndDomainsFromSubject: false,
        isCheckKeywordAndRecipientsIncludeSubject: false,
        isAutoCheckRegisteredInContacts: false,
        isAutoCheckRegisteredInContactsAndMemberOfContactLists: false,
        isWarningIfRecipientsIsNotRegistered: false,
        isProhibitsSendingMailIfRecipientsIsNotRegistered: false,
        isShowConfirmationAtSendMeetingRequest: false,
        isShowConfirmationAtSendTaskRequest: false,
      },
      whitelist: [],
      internalDomains: [],
      alertAddresses: [],
      alertKeywordsBody: [],
      alertKeywordsSubject: [],
      autoCcBccKeyword: [],
      autoCcBccRecipient: [],
      autoCcBccAttachedFile: [],
      nameAndDomains: [],
      keywordAndRecipients: [],
      externalDomains: {
        targetToAndCcExternalDomainsNum: 10,
        isWarningWhenLargeNumberOfExternalDomains: true,
        isProhibitedWhenLargeNumberOfExternalDomains: false,
        isAutoChangeToBccWhenLargeNumberOfExternalDomains: false,
      },
      forceAutoChangeRecipientsToBcc: {
        isForceAutoChangeRecipientsToBcc: false,
        toRecipient: "",
        isIncludeInternalDomain: false,
      },
      attachmentsSetting: {
        isWarningWhenEncryptedZipIsAttached: false,
        isProhibitedWhenEncryptedZipIsAttached: false,
        isEnableAllAttachedFilesAreDetectEncryptedZip: false,
        isAttachmentsProhibited: false,
        isWarningWhenAttachedRealFile: false,
        isEnableOpenAttachedFiles: false,
        targetAttachmentFileExtensionOfOpen:
          ".pdf,.txt,.csv,.rtf,.htm,.html,.doc,.docx,.xls,.xlm,.xlsm,.xlsx,.ppt,.pptx,.bmp,.gif,.jpg,.jpeg,.png,.tif,.pub,.vsd,.vsdx",
        isMustOpenBeforeCheckTheAttachedFiles: false,
        isIgnoreMustOpenBeforeCheckTheAttachedFilesIfInternalDomain: false,
      },
      recipientsAndAttachmentsName: [],
      attachmentProhibitedRecipients: [],
      attachmentAlertRecipients: [],
      autoDeleteRecipients: [],
      autoAddMessage: {
        isAddToStart: false,
        isAddToEnd: false,
        messageOfAddToStart: "",
        messageOfAddToEnd: "",
      },
      securityForReceivedMail: {
        isEnableSecurityForReceivedMail: false,
        isEnableAlertKeywordOfSubjectWhenOpeningMailsData: false,
        isEnableMailHeaderAnalysis: false,
        isShowWarningWhenSpfFails: false,
        isShowWarningWhenDkimFails: false,
        isEnableWarningFeatureWhenOpeningAttachments: false,
        isWarnBeforeOpeningAttachments: false,
        isWarnBeforeOpeningEncryptedZip: false,
        isWarnLinkFileInTheZip: false,
        isWarnOneFileInTheZip: false,
        isWarnOfficeFileWithMacroInTheZip: false,
        isWarnBeforeOpeningAttachmentsThatContainMacros: false,
        isShowWarningWhenSpoofingRisk: false,
        isShowWarningWhenDmarcNotImplemented: false,
      },
      alertKeywordOfSubjectWhenOpeningMail: [],

      // New add-in specific knobs (no CSV equivalent)
      runtime: {
        largeAttachmentBytes: 10485760, // 10MB (same as OutlookOkan)
        sendEventTimeoutMs: 4500,
      },
    };
  }

  function normalizeWhitelist(list) {
    if (!Array.isArray(list)) return [];
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      if (!isObject(row)) continue;
      var whiteName = normalizeString(row.whiteName != null ? row.whiteName : row.WhiteName);
      if (!whiteName) continue;
      out.push({
        whiteName: whiteName,
        isSkipConfirmation: normalizeBool(
          row.isSkipConfirmation != null ? row.isSkipConfirmation : row.IsSkipConfirmation,
          false
        ),
      });
    }
    return out;
  }

  function normalizeInternalDomains(list) {
    if (!Array.isArray(list)) return [];
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      if (!isObject(row)) continue;
      var domain = normalizeString(row.domain != null ? row.domain : row.Domain);
      if (!domain) continue;
      out.push({ domain: domain });
    }
    return out;
  }

  function normalizeAlertAddresses(list) {
    if (!Array.isArray(list)) return [];
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      if (!isObject(row)) continue;
      var targetAddress = normalizeString(row.targetAddress != null ? row.targetAddress : row.TargetAddress);
      if (!targetAddress) continue;
      out.push({
        targetAddress: targetAddress,
        isCanNotSend: normalizeBool(row.isCanNotSend != null ? row.isCanNotSend : row.IsCanNotSend, false),
        message: normalizeString(row.message != null ? row.message : row.Message),
      });
    }
    return out;
  }

  function normalizeAlertKeywords(list) {
    if (!Array.isArray(list)) return [];
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      if (!isObject(row)) continue;
      var alertKeyword = normalizeString(row.alertKeyword != null ? row.alertKeyword : row.AlertKeyword);
      if (!alertKeyword) continue;
      out.push({
        alertKeyword: alertKeyword,
        message: normalizeString(row.message != null ? row.message : row.Message),
        isCanNotSend: normalizeBool(row.isCanNotSend != null ? row.isCanNotSend : row.IsCanNotSend, false),
      });
    }
    return out;
  }

  function normalizeAutoCcBccKeyword(list) {
    if (!Array.isArray(list)) return [];
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      if (!isObject(row)) continue;
      var keyword = normalizeString(row.keyword != null ? row.keyword : row.Keyword);
      var autoAddAddress = normalizeString(row.autoAddAddress != null ? row.autoAddAddress : row.AutoAddAddress);
      if (!keyword || !autoAddAddress) continue;
      out.push({
        keyword: keyword,
        ccOrBcc: normalizeCcOrBcc(row.ccOrBcc != null ? row.ccOrBcc : row.CcOrBcc),
        autoAddAddress: autoAddAddress,
      });
    }
    return out;
  }

  function normalizeAutoCcBccRecipient(list) {
    if (!Array.isArray(list)) return [];
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      if (!isObject(row)) continue;
      var targetRecipient = normalizeString(row.targetRecipient != null ? row.targetRecipient : row.TargetRecipient);
      var autoAddAddress = normalizeString(row.autoAddAddress != null ? row.autoAddAddress : row.AutoAddAddress);
      if (!targetRecipient || !autoAddAddress) continue;
      out.push({
        targetRecipient: targetRecipient,
        ccOrBcc: normalizeCcOrBcc(row.ccOrBcc != null ? row.ccOrBcc : row.CcOrBcc),
        autoAddAddress: autoAddAddress,
      });
    }
    return out;
  }

  function normalizeAutoCcBccAttachedFile(list) {
    if (!Array.isArray(list)) return [];
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      if (!isObject(row)) continue;
      var autoAddAddress = normalizeString(row.autoAddAddress != null ? row.autoAddAddress : row.AutoAddAddress);
      if (!autoAddAddress) continue;
      out.push({
        ccOrBcc: normalizeCcOrBcc(row.ccOrBcc != null ? row.ccOrBcc : row.CcOrBcc),
        autoAddAddress: autoAddAddress,
      });
    }
    return out;
  }

  function normalizeNameAndDomains(list) {
    if (!Array.isArray(list)) return [];
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      if (!isObject(row)) continue;
      var name = normalizeString(row.name != null ? row.name : row.Name);
      var domain = normalizeString(row.domain != null ? row.domain : row.Domain);
      if (!name || !domain) continue;
      out.push({ name: name, domain: domain });
    }
    return out;
  }

  function normalizeKeywordAndRecipients(list) {
    if (!Array.isArray(list)) return [];
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      if (!isObject(row)) continue;
      var keyword = normalizeString(row.keyword != null ? row.keyword : row.Keyword);
      var recipient = normalizeString(row.recipient != null ? row.recipient : row.Recipient);
      if (!keyword || !recipient) continue;
      out.push({ keyword: keyword, recipient: recipient });
    }
    return out;
  }

  function normalizeRecipientsAndAttachmentsName(list) {
    if (!Array.isArray(list)) return [];
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      if (!isObject(row)) continue;
      var attachmentsName = normalizeString(
        row.attachmentsName != null ? row.attachmentsName : row.AttachmentsName
      );
      var recipient = normalizeString(row.recipient != null ? row.recipient : row.Recipient);
      if (!attachmentsName || !recipient) continue;
      out.push({ attachmentsName: attachmentsName, recipient: recipient });
    }
    return out;
  }

  function normalizeAttachmentProhibitedRecipients(list) {
    if (!Array.isArray(list)) return [];
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      if (!isObject(row)) continue;
      var recipient = normalizeString(row.recipient != null ? row.recipient : row.Recipient);
      if (!recipient) continue;
      out.push({ recipient: recipient });
    }
    return out;
  }

  function normalizeAttachmentAlertRecipients(list) {
    if (!Array.isArray(list)) return [];
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      if (!isObject(row)) continue;
      var recipient = normalizeString(row.recipient != null ? row.recipient : row.Recipient);
      if (!recipient) continue;
      out.push({
        recipient: recipient,
        message: normalizeString(row.message != null ? row.message : row.Message),
      });
    }
    return out;
  }

  function normalizeAutoDeleteRecipients(list) {
    if (!Array.isArray(list)) return [];
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      if (!isObject(row)) continue;
      var recipient = normalizeString(row.recipient != null ? row.recipient : row.Recipient);
      if (!recipient) continue;
      out.push({ recipient: recipient });
    }
    return out;
  }

  function normalizeAlertKeywordOfSubjectWhenOpeningMail(list) {
    if (!Array.isArray(list)) return [];
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var row = list[i];
      if (!isObject(row)) continue;
      var alertKeyword = normalizeString(row.alertKeyword != null ? row.alertKeyword : row.AlertKeyword);
      if (!alertKeyword) continue;
      out.push({
        alertKeyword: alertKeyword,
        message: normalizeString(row.message != null ? row.message : row.Message),
      });
    }
    return out;
  }

  function normalizeSettings(settings) {
    var defaults = defaultSettings();
    var merged = mergeDefaults(isObject(settings) ? settings : {}, defaults);

    merged.schemaVersion = SCHEMA_VERSION;

    merged.whitelist = normalizeWhitelist(merged.whitelist);
    merged.internalDomains = normalizeInternalDomains(merged.internalDomains);
    merged.alertAddresses = normalizeAlertAddresses(merged.alertAddresses);
    merged.alertKeywordsBody = normalizeAlertKeywords(merged.alertKeywordsBody);
    merged.alertKeywordsSubject = normalizeAlertKeywords(merged.alertKeywordsSubject);
    merged.autoCcBccKeyword = normalizeAutoCcBccKeyword(merged.autoCcBccKeyword);
    merged.autoCcBccRecipient = normalizeAutoCcBccRecipient(merged.autoCcBccRecipient);
    merged.autoCcBccAttachedFile = normalizeAutoCcBccAttachedFile(merged.autoCcBccAttachedFile);
    merged.nameAndDomains = normalizeNameAndDomains(merged.nameAndDomains);
    merged.keywordAndRecipients = normalizeKeywordAndRecipients(merged.keywordAndRecipients);
    merged.recipientsAndAttachmentsName = normalizeRecipientsAndAttachmentsName(merged.recipientsAndAttachmentsName);
    merged.attachmentProhibitedRecipients = normalizeAttachmentProhibitedRecipients(merged.attachmentProhibitedRecipients);
    merged.attachmentAlertRecipients = normalizeAttachmentAlertRecipients(merged.attachmentAlertRecipients);
    merged.autoDeleteRecipients = normalizeAutoDeleteRecipients(merged.autoDeleteRecipients);
    merged.alertKeywordOfSubjectWhenOpeningMail = normalizeAlertKeywordOfSubjectWhenOpeningMail(
      merged.alertKeywordOfSubjectWhenOpeningMail
    );

    return merged;
  }

  async function loadSettings() {
    if (!storage) return normalizeSettings(null);
    var data = await storage.getJson(SETTINGS_KEY);
    return normalizeSettings(data);
  }

  async function saveSettings(settings) {
    if (!storage) return;
    await storage.setJson(SETTINGS_KEY, normalizeSettings(settings));
  }

  async function resetSettings() {
    if (!storage) return;
    await storage.removeItem(SETTINGS_KEY);
  }

  ns.settings = {
    KEY: SETTINGS_KEY,
    SCHEMA_VERSION: SCHEMA_VERSION,
    defaults: defaultSettings,
    load: loadSettings,
    save: saveSettings,
    reset: resetSettings,
    _normalize: normalizeSettings,
  };
})();
