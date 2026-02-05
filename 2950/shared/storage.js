"use strict";
(function () {
  var root =
    typeof globalThis !== "undefined" ? globalThis : typeof window !== "undefined" ? window : undefined;
  if (!root) return;

  var ns = (root.MailChecker = root.MailChecker || {});

  function canUseOfficeRuntimeStorage() {
    try {
      return (
        typeof OfficeRuntime !== "undefined" &&
        OfficeRuntime &&
        OfficeRuntime.storage &&
        typeof OfficeRuntime.storage.getItem === "function" &&
        typeof OfficeRuntime.storage.setItem === "function"
      );
    } catch (_e) {
      return false;
    }
  }

  function canUseLocalStorage() {
    try {
      return typeof localStorage !== "undefined" && localStorage && typeof localStorage.getItem === "function";
    } catch (_e) {
      return false;
    }
  }

  async function getItem(key) {
    if (!key) return null;

    if (canUseOfficeRuntimeStorage()) {
      try {
        var value = await OfficeRuntime.storage.getItem(String(key));
        return value == null ? null : String(value);
      } catch (_e) {}
    }

    if (canUseLocalStorage()) {
      try {
        var localValue = localStorage.getItem(String(key));
        return localValue == null ? null : String(localValue);
      } catch (_e2) {}
    }

    return null;
  }

  async function setItem(key, value) {
    if (!key) return;
    var stringValue = value == null ? "" : String(value);

    if (canUseOfficeRuntimeStorage()) {
      try {
        await OfficeRuntime.storage.setItem(String(key), stringValue);
        return;
      } catch (_e) {}
    }

    if (canUseLocalStorage()) {
      try {
        localStorage.setItem(String(key), stringValue);
      } catch (_e2) {}
    }
  }

  async function removeItem(key) {
    if (!key) return;

    if (canUseOfficeRuntimeStorage()) {
      try {
        await OfficeRuntime.storage.removeItem(String(key));
      } catch (_e) {}
    }

    if (canUseLocalStorage()) {
      try {
        localStorage.removeItem(String(key));
      } catch (_e2) {}
    }
  }

  async function getJson(key) {
    var raw = await getItem(key);
    if (!raw) return null;
    try {
      return JSON.parse(raw);
    } catch (_e) {
      return null;
    }
  }

  async function setJson(key, obj) {
    try {
      await setItem(key, JSON.stringify(obj == null ? null : obj));
    } catch (_e) {}
  }

  ns.storage = {
    getItem: getItem,
    setItem: setItem,
    removeItem: removeItem,
    getJson: getJson,
    setJson: setJson,
    _canUseOfficeRuntimeStorage: canUseOfficeRuntimeStorage,
  };
})();

