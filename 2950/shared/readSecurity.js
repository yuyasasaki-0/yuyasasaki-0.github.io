"use strict";
(function () {
  var root =
    typeof globalThis !== "undefined" ? globalThis : typeof window !== "undefined" ? window : undefined;
  if (!root) return;

  var ns = (root.MailChecker = root.MailChecker || {});

  function normalizeString(value) {
    return value == null ? "" : String(value).trim();
  }

  function upper(value) {
    return value == null ? "" : String(value).toUpperCase();
  }

  function extractMainDomain(domain) {
    var parts = String(domain || "").split(".");
    if (parts.length > 2) return parts.slice(parts.length - 3).join(".");
    return String(domain || "");
  }

  function isInternalMail(emailHeader) {
    if (!emailHeader) return false;

    var received = String(emailHeader).match(/^Received:.*$/gim) || [];
    if (received.length > 3) return false;

    var domainRegex = /from\s([^\s]+)/i;
    var prev = null;
    for (var i = 0; i < received.length; i++) {
      var m = received[i].match(domainRegex);
      if (!m || !m[1]) continue;
      var cur = extractMainDomain(String(m[1]));
      if (prev && cur && prev.toLowerCase() === cur.toLowerCase()) return true;
      prev = cur;
    }
    return false;
  }

  function validateEmailHeader(emailHeader) {
    if (!emailHeader) return null;

    var header = String(emailHeader);

    var results = {
      "From Domain": "NONE",
      "ReturnPath Domain": "NONE",
      SPF: "NONE",
      "SPF IP": "NONE",
      "SPF Alignment": "NONE",
      DKIM: "NONE",
      "DKIM Domain": "NONE",
      "DKIM Alignment": "NONE",
      DMARC: "NONE",
      Internal: "FALSE",
    };

    if (isInternalMail(header)) results.Internal = "TRUE";

    var fromDomain = "";
    try {
      var fromHeaderMatch = header.match(/^From:\s*.*(?:\r?\n\s+.*)*/im);
      if (fromHeaderMatch && fromHeaderMatch[0]) {
        var fromHeader = fromHeaderMatch[0];
        var domainMatch = fromHeader.match(/<.*?@([^\s>]+)>/i);
        if (!domainMatch) domainMatch = fromHeader.match(/[^<\s]+@([^\s>]+)/i);
        if (domainMatch && domainMatch[1]) {
          fromDomain = String(domainMatch[1]);
          results["From Domain"] = fromDomain;
        }
      }
    } catch (_e) {}

    // SPF
    try {
      var spfMatch = header.match(
        /Received-SPF:\s*(pass|fail|softfail|neutral|temperror|permerror|none)[\s\S]*?\b(?:does\s+not\s+)?designate[s]?\s+([^ ]+)\s+as/i
      );
      if (spfMatch) {
        results.SPF = upper(spfMatch[1]);
        results["SPF IP"] = String(spfMatch[2] || "");
      }
    } catch (_e2) {}

    // SPF alignment
    try {
      var returnPathMatch = header.match(/Return-Path:\s*.*@([^\s>]+)/i);
      if (returnPathMatch && returnPathMatch[1] && fromDomain) {
        var returnPathDomain = String(returnPathMatch[1]);
        results["ReturnPath Domain"] = returnPathDomain;
        var a = returnPathDomain.toLowerCase();
        var b = fromDomain.toLowerCase();
        results["SPF Alignment"] = a === b || a.indexOf(b) >= 0 || b.indexOf(a) >= 0 ? "PASS" : "FAIL";
      }
    } catch (_e3) {}

    // DKIM
    try {
      var dkimMatch = header.match(
        /Authentication-Results:[\s\S]*?dkim=(pass|policy|fail|softfail|hardfail|neutral|temperror|permerror|none)[\s\S]*?header\.d=([^;( )]+)/i
      );
      if (dkimMatch) results.DKIM = upper(dkimMatch[1]);
    } catch (_e4) {}

    // DKIM alignment
    try {
      var dkimSignatureRegex = /DKIM-Signature:[\s\S]*?d=([^;( )]+)/gi;
      var domains = [];
      var alignPass = false;
      var match;
      while ((match = dkimSignatureRegex.exec(header))) {
        var d = match[1];
        if (!d) continue;
        domains.push(d);
        if (fromDomain) {
          var da = String(d).toLowerCase();
          var fa = String(fromDomain).toLowerCase();
          if (da === fa || da.indexOf(fa) >= 0 || fa.indexOf(da) >= 0) alignPass = true;
        }
      }
      results["DKIM Domain"] = domains.join(", ");
      results["DKIM Alignment"] = alignPass ? "PASS" : "FAIL";
    } catch (_e5) {}

    // DMARC
    try {
      var dmarcMatch = header.match(/Authentication-Results:[\s\S]*?dmarc=(pass|bestguesspass|softfail|fail|none)/i);
      if (dmarcMatch) results.DMARC = upper(dmarcMatch[1]);
    } catch (_e6) {}

    return results;
  }

  function determineDmarcResult(spfResult, spfAlignmentResult, dkimResult, dkimAlignmentResult) {
    function normalize(value) {
      var v = upper(normalizeString(value));
      return v === "NONE" ? "FAIL" : v || "FAIL";
    }

    var key =
      normalize(spfResult) +
      "_" +
      normalize(spfAlignmentResult) +
      "_" +
      normalize(dkimResult) +
      "_" +
      normalize(dkimAlignmentResult);

    var map = {
      PASS_PASS_PASS_PASS: "PASS",
      PASS_PASS_PASS_FAIL: "PASS",
      PASS_PASS_FAIL_PASS: "PASS",
      PASS_PASS_FAIL_FAIL: "PASS",
      PASS_FAIL_PASS_PASS: "PASS",
      FAIL_PASS_PASS_PASS: "PASS",
      FAIL_FAIL_PASS_PASS: "PASS",

      PASS_FAIL_PASS_FAIL: "FAIL",
      PASS_FAIL_FAIL_PASS: "FAIL",
      PASS_FAIL_FAIL_FAIL: "FAIL",
      FAIL_PASS_PASS_FAIL: "FAIL",
      FAIL_PASS_FAIL_PASS: "FAIL",
      FAIL_PASS_FAIL_FAIL: "FAIL",
      FAIL_FAIL_PASS_FAIL: "FAIL",
      FAIL_FAIL_FAIL_PASS: "FAIL",
      FAIL_FAIL_FAIL_FAIL: "FAIL",
    };

    return map[key] || "FAIL";
  }

  function readUInt16LE(bytes, offset) {
    return bytes[offset] | (bytes[offset + 1] << 8);
  }

  function readUInt32LE(bytes, offset) {
    return (
      (bytes[offset] |
        (bytes[offset + 1] << 8) |
        (bytes[offset + 2] << 16) |
        (bytes[offset + 3] << 24)) >>>
      0
    );
  }

  function decodeBytes(bytes) {
    try {
      if (typeof TextDecoder !== "undefined") {
        return new TextDecoder("utf-8", { fatal: false }).decode(bytes);
      }
    } catch (_e) {}
    var out = "";
    for (var i = 0; i < bytes.length; i++) out += String.fromCharCode(bytes[i]);
    return out;
  }

  function parseZipCentralDirectory(bytes) {
    if (!bytes || typeof bytes.length !== "number") return { isZip: false };
    var len = bytes.length;
    if (len < 22) return { isZip: false };

    // Find EOCD
    var min = Math.max(0, len - 65557);
    var eocd = -1;
    for (var i = len - 22; i >= min; i--) {
      if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x05 && bytes[i + 3] === 0x06) {
        eocd = i;
        break;
      }
    }
    if (eocd < 0) return { isZip: false };

    var cdSize = readUInt32LE(bytes, eocd + 12);
    var cdOffset = readUInt32LE(bytes, eocd + 16);
    if (cdOffset + cdSize > len) return { isZip: true, isEncrypted: false, includeExtensions: [], fileNames: [] };

    var pos = cdOffset;
    var end = cdOffset + cdSize;
    var isEncrypted = false;
    var extMap = {};
    var includeExtensions = [];
    var fileNames = [];

    while (pos + 46 <= end && pos + 46 <= len) {
      if (readUInt32LE(bytes, pos) !== 0x02014b50) break;

      var flags = readUInt16LE(bytes, pos + 8);
      if ((flags & 0x0001) !== 0) isEncrypted = true;

      var fileNameLen = readUInt16LE(bytes, pos + 28);
      var extraLen = readUInt16LE(bytes, pos + 30);
      var commentLen = readUInt16LE(bytes, pos + 32);

      var nameBytes = bytes.slice(pos + 46, pos + 46 + fileNameLen);
      var name = decodeBytes(nameBytes);
      fileNames.push(name);

      var dot = name.lastIndexOf(".");
      if (dot >= 0) {
        var ext = name.slice(dot).toLowerCase();
        if (!extMap[ext]) {
          extMap[ext] = true;
          includeExtensions.push(ext);
        }
      }

      pos = pos + 46 + fileNameLen + extraLen + commentLen;
      if (fileNames.length > 500) break;
    }

    return {
      isZip: true,
      isEncrypted: isEncrypted,
      includeExtensions: includeExtensions,
      fileNames: fileNames,
    };
  }

  ns.readSecurity = {
    validateEmailHeader: validateEmailHeader,
    determineDmarcResult: determineDmarcResult,
    parseZipCentralDirectory: parseZipCentralDirectory,
  };
})();

